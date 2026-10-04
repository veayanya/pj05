#!/usr/bin/env python3
"""
Web server pengisi formulir Word (.docx)

Alur:
  1. Upload file .docx
  2. Server mendeteksi otomatis semua isian: "titik-titik" (.....), kotak centang (☐ / kotak
     ActiveX), isian "Label : nilai" (nilai lama bisa diedit), dan sel tabel yang masih kosong
  3. Isi / edit formulir di browser (tampilan mirip dokumen asli)
  4. Klik "Unduh Word" -> file .docx dengan isian & centang sudah terisi

Jalankan:   python app.py      lalu buka  http://127.0.0.1:5000
"""
import base64
import html as H
import io
import math
import os
import re
import time
import uuid
import zipfile
from copy import deepcopy
from datetime import date
from pathlib import Path

from flask import (Flask, abort, redirect, render_template_string, request,
                   send_file, url_for)
from lxml import etree

# --------------------------------------------------------------------------
# Konfigurasi
# --------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
FORMS_DIR = BASE_DIR / "formulir"
TEMPLATES = {                      # formulir bawaan untuk yang belum punya file kosongannya
    "pendaftaran-magang": ("Formulir Pendaftaran Peserta Magang",
                           "Identitas mahasiswa, kelengkapan berkas, pernyataan, dan verifikasi program studi.",
                           "FORMULIR_PENDAFTARAN_PESERTA_MAGANG.docx"),
    "konversi-nilai": ("Form Permohonan Konversi Nilai MBKM",
                       "Identitas, informasi kegiatan, tabel konversi mata kuliah, dokumen pendukung, dan tanda tangan.",
                       "Form_Permohonan_Konversi_Nilai.docx"),
}
MAX_UPLOAD = 3 * 1024 * 1024           # 3 MB (batas body Vercel 4,5 MB; file dikirim bolak-balik sebagai base64)
MAX_XML = 40 * 1024 * 1024             # batas ukuran document.xml setelah dekompresi
KEEP_SECONDS = 12 * 3600               # file sementara dihapus setelah 12 jam
CHECKED = "☑"
UNCHECKED = "☐"
BULAN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli",
         "Agustus", "September", "Oktober", "November", "Desember"]

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
MAX_ROWS = 30                          # batas baris tabel dinamis
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
V_NS = "urn:schemas-microsoft-com:vml"
NS = {"w": W_NS}
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def q(name):
    return f"{{{W_NS}}}{name}"


# Lebar huruf Times New Roman (per 1000 em) untuk memperkirakan lebar teks
_TW = {' ': 250, '!': 333, '"': 408, '#': 500, '$': 500, '%': 833, '&': 778, "'": 180,
       '(': 333, ')': 333, '*': 500, '+': 564, ',': 250, '-': 333, '.': 250, '/': 278,
       ':': 278, ';': 278, '<': 564, '=': 564, '>': 564, '?': 444, '@': 921, '[': 333,
       '\\': 278, ']': 333, '^': 469, '_': 500, '`': 333, '{': 480, '|': 200, '}': 480, '~': 541}
_TW.update({d: 500 for d in "0123456789"})
_TW.update(zip("ABCDEFGHIJKLMNOPQRSTUVWXYZ",
               [722, 667, 667, 722, 611, 556, 722, 722, 333, 389, 722, 611, 889, 722, 722, 556,
                722, 667, 556, 611, 722, 722, 944, 722, 722, 611]))
_TW.update(zip("abcdefghijklmnopqrstuvwxyz",
               [444, 500, 444, 500, 444, 333, 500, 500, 278, 278, 500, 278, 778, 500, 500, 500,
                500, 333, 389, 278, 500, 500, 722, 500, 500, 444]))
_TW_OTHER = {"☐": 900, "☑": 900, "☒": 900, "✓": 800, "≥": 549, "≤": 549, "×": 564, "…": 1000}


def tw(text, size, bold=False):
    """Perkiraan lebar teks dalam pt (sedikit dilebihkan agar aman)."""
    total = 0
    for ch in text:
        w = _TW.get(ch) or _TW_OTHER.get(ch) or (1000 if ord(ch) >= 0x2E80 else 556)
        total += w
    return total / 1000 * size * (1.05 if bold else 1.0) * 1.02


PPR_AFTER_SPACING = ["ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc",
                     "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId",
                     "cnfStyle", "rPr", "sectPr", "pPrChange"]
PPR_AFTER_TABS = ["suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct", "topLinePunct",
                  "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid",
                  "spacing"] + PPR_AFTER_SPACING
RPR_AFTER_SZ = ["highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs",
                "em", "lang", "eastAsianLayout", "specVanish", "oMath"]


def set_sz(node, hp, cs=True):
    """Set ukuran font (half-point) pada <w:r> atau <w:rPr>, menjaga urutan elemen skema."""
    rpr = node if node.tag == q("rPr") else node.find(q("rPr"))
    if rpr is None:
        rpr = etree.Element(q("rPr"))
        node.insert(0, rpr)
    for tag in (("sz", "szCs") if cs else ("sz",)):
        e = rpr.find(q(tag))
        if e is None:
            e = etree.Element(q(tag))
            idx = len(rpr)
            if tag == "szCs" and rpr.find(q("sz")) is not None:
                idx = list(rpr).index(rpr.find(q("sz"))) + 1
            else:
                for i, c in enumerate(rpr):
                    if etree.QName(c).localname in RPR_AFTER_SZ + (["szCs"] if tag == "sz" else []):
                        idx = i
                        break
            rpr.insert(idx, e)
        e.set(q("val"), str(hp))


# --------------------------------------------------------------------------
# Model "slot" = bagian dokumen yang bisa diisi
# --------------------------------------------------------------------------
TOKEN = re.compile(r"(\.{4,}|…{2,}|☐)")


class Slot:
    def __init__(self, kind, orig):
        self.kind = kind          # text | check | date
        self.orig = orig
        self.id = None
        self.label = ""
        self.link = None          # kunci sinkron otomatis (nama / nim)
        self.group = None         # grup centang saling-eksklusif
        self.full = False         # input selebar baris
        self.blank = False        # isian tanpa titik, mis. "Nama :"
        self.before = ""
        self.sub = None           # untuk date: (titik_hari_bulan, teks_tengah, titik_tahun)
        self.preset = False       # label sudah ditetapkan saat deteksi
        self.prefill = False      # isian "Label : nilai" yang sudah ada nilainya (bisa diedit)
        self.nospace = False      # isian blank tanpa spasi pemisah (sel tabel)
        self.iso = None           # tanggal awal (YYYY-MM-DD) untuk date yang sudah terisi
        self.default_today = True # date kosong diisi tanggal hari ini sebagai awal


def tokenize(text):
    parts = []
    for piece in TOKEN.split(text):
        if not piece:
            continue
        if piece == UNCHECKED:
            parts.append(Slot("check", piece))
        elif TOKEN.fullmatch(piece):
            parts.append(Slot("text", piece))
        else:
            parts.append(piece)
    return parts


def merge_dates(parts):
    """'............... 20.....' -> satu slot tanggal."""
    out, i = [], 0
    while i < len(parts):
        if (i + 2 < len(parts)
                and isinstance(parts[i], Slot) and parts[i].kind == "text"
                and isinstance(parts[i + 1], str) and re.fullmatch(r"\s*20", parts[i + 1])
                and isinstance(parts[i + 2], Slot) and parts[i + 2].kind == "text"
                and len(parts[i + 2].orig) <= 8):
            s = Slot("date", parts[i].orig)
            s.sub = (parts[i], parts[i + 1], parts[i + 2])
            out.append(s)
            i += 3
        else:
            out.append(parts[i])
            i += 1
    return out


def merge_runs(root):
    """Gabungkan run bersebelahan yang formatnya sama agar titik-titik tidak terpotong."""
    def simple(r):
        kids = [c for c in r if c.tag != q("rPr")]
        return len(kids) == 1 and kids[0].tag == q("t")

    def key(r):
        rp = r.find(q("rPr"))
        return b"" if rp is None else etree.tostring(rp, method="c14n")

    # penanda spell-check Word memecah teks jadi banyak run; aman dihapus
    for pe in list(root.iter(q("proofErr"))):
        pe.getparent().remove(pe)

    for p in root.iter(q("p")):
        prev = None
        for r in list(p):
            if r.tag == q("r") and simple(r):
                if prev is not None and key(prev) == key(r):
                    pt = prev.find(q("t"))
                    pt.text = (pt.text or "") + (r.find(q("t")).text or "")
                    p.remove(r)
                    continue
                prev = r
            else:
                prev = None


def clean_value(v):
    v = (v or "").replace("\r", " ").replace("\n", " ")
    v = "".join(ch for ch in v if ch == "\t" or ord(ch) >= 32)
    return v.strip()[:300]


MONTH_RE = "|".join(BULAN)
DATE_RE = re.compile(rf"\b\d{{1,2}}\s+(?:{MONTH_RE})\s+\d{{4}}\b", re.I)
NEXT_FIELD = re.compile(r"\t[^:\t]{1,32}:")
MULTI_CUE = re.compile(r"✓|✔|centang|lampir|dokumen|berkas|checklist|tanda", re.I)


def parse_id_date(txt):
    m = re.fullmatch(r"(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})", txt.strip())
    if not m:
        return None
    names = [b.lower() for b in BULAN]
    if m.group(2).lower() not in names:
        return None
    try:
        return date(int(m.group(3)), names.index(m.group(2).lower()) + 1, int(m.group(1))).isoformat()
    except ValueError:
        return None


# --------------------------------------------------------------------------
# Dokumen formulir
# --------------------------------------------------------------------------
class FormDoc:
    def __init__(self, data: bytes, rows=None):
        self.data = data
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            names = z.namelist()
            if "word/document.xml" not in names:
                raise ValueError("Bukan file Word (.docx) yang valid.")
            if z.getinfo("word/document.xml").file_size > MAX_XML:
                raise ValueError("Isi dokumen terlalu besar.")
            xml = z.read("word/document.xml")
            styles = z.read("word/styles.xml") if "word/styles.xml" in names else None
        parser = etree.XMLParser(resolve_entities=False, no_network=True)
        self.root = etree.fromstring(xml, parser)
        self.styles = etree.fromstring(styles, parser) if styles else None
        merge_runs(self.root)
        self._read_defaults()
        self.removed_rids = set()
        self.cell_ts = set()
        self._convert_activex()
        self._setup_dynamic(rows or {})
        self._add_cell_runs()
        self._mark_editable()
        self.analyze()

    # ----- default gaya dari styles.xml --------------------------------
    def _read_defaults(self):
        d = {"before": 0, "after": 0, "line": 240}
        if self.styles is not None:
            nodes = self.styles.xpath("//w:docDefaults/w:pPrDefault/w:pPr/w:spacing", namespaces=NS)
            nodes += self.styles.xpath(
                "//w:style[@w:type='paragraph' and @w:default='1']/w:pPr/w:spacing", namespaces=NS)
            for sp in nodes:
                for k in ("before", "after", "line"):
                    v = sp.get(q(k))
                    if v is not None and v.lstrip("-").isdigit():
                        d[k] = int(v)
        self.defaults = d
        # ukuran font default (half-point)
        self.default_hp = 22
        if self.styles is not None:
            for path in ("//w:style[@w:type='paragraph' and @w:default='1']/w:rPr/w:sz",
                         "//w:docDefaults/w:rPrDefault/w:rPr/w:sz"):
                n = self.styles.xpath(path, namespaces=NS)
                if n and (n[0].get(q("val")) or "").isdigit():
                    self.default_hp = int(n[0].get(q("val")))
                    break
        self._style_hp_cache = {}
        # ukuran halaman & margin (twips)
        pgd = {"w": 12240, "h": 15840, "top": 1440, "bottom": 1440, "left": 1440, "right": 1440}
        body = self.root.find(q("body"))
        sect = body.find(q("sectPr")) if body is not None else None
        if sect is None:
            sect = self.root.find(f".//{q('sectPr')}")
        self.sect = sect
        if sect is not None:
            pg = sect.find(q("pgSz"))
            mar = sect.find(q("pgMar"))
            try:
                if pg is not None:
                    pgd["w"], pgd["h"] = int(pg.get(q("w"))), int(pg.get(q("h")))
                if mar is not None:
                    for k in ("top", "bottom", "left", "right"):
                        pgd[k] = int(mar.get(q(k)))
            except (TypeError, ValueError):
                pass
        self.pg = pgd
        self.text_width = max(pgd["w"] - pgd["left"] - pgd["right"], 1000)

    # ----- pra-proses: ubah elemen yang tidak bisa diisi langsung -------------
    def _convert_activex(self):
        """Kotak centang ActiveX (<w:object>/<w:control>) -> karakter ☐ biasa."""
        for r in list(self.root.iter(q("r"))):
            obj = r.find(q("object"))
            if obj is None or obj.find(q("control")) is None:
                continue
            for e in obj.iter():
                for k, v in e.attrib.items():
                    if k in (f"{{{R_NS}}}id", f"{{{R_NS}}}embed"):
                        self.removed_rids.add(v)
            parent = r.getparent()
            src = None
            nxt = r.getnext()
            while nxt is not None and src is None:
                if nxt.tag == q("r") and nxt.find(q("rPr")) is not None:
                    src = nxt.find(q("rPr"))
                nxt = nxt.getnext()
            if src is None:
                ppr = parent.find(q("pPr")) if parent.tag == q("p") else None
                src = ppr.find(q("rPr")) if ppr is not None else None
            nr = etree.Element(q("r"))
            if src is not None:
                nr.append(deepcopy(src))
            t = etree.SubElement(nr, q("t"))
            t.text = UNCHECKED
            parent.replace(r, nr)

    # ----- tabel dinamis (baris bisa ditambah / dikurangi) ---------------
    @classmethod
    def _row_first(cls, tr):
        tcs = tr.findall(q("tc"))
        return cls._cell_text(tcs[0]) if tcs else ""

    def _data_rows(self, tbl):
        """Baris isian tabel: baris setelah header yang kolom pertamanya bernomor."""
        trs = tbl.findall(q("tr"))
        num = [tr for tr in trs[1:] if re.fullmatch(r"\d+", self._row_first(tr))]
        return num or trs[1:]

    def _setup_dynamic(self, rows):
        """Tabel berheader 'No' dianggap tabel dinamis. Jika jumlah baris diminta (rows[idx]),
        baris disesuaikan: kelebihan dibuang, kekurangan disalin dari baris terakhir (dikosongkan)."""
        self.dyn = []
        self._dyn_ids = {}
        for tbl in self.root.iter(q("tbl")):
            if any(a.tag == q("tc") for a in tbl.iterancestors()):
                continue
            trs = tbl.findall(q("tr"))
            tcs = trs[0].findall(q("tc")) if trs else []
            if len(trs) < 2 or not tcs or not re.fullmatch(r"no\.?", self._cell_text(tcs[0]).strip(), re.I):
                continue
            idx = len(self.dyn)
            self.dyn.append(tbl)
            if idx in rows:
                self._resize_rows(tbl, max(1, min(int(rows[idx]), MAX_ROWS)))
            for ri, tr in enumerate(self._data_rows(tbl)):
                self._dyn_ids[tr] = (idx, ri)

    def _resize_rows(self, tbl, n):
        data = self._data_rows(tbl)
        if not data:
            return
        for tr in data[n:]:
            tbl.remove(tr)
        data = data[:n]
        last = data[-1]
        while len(data) < n:
            new = self._blank_row(data[-1])
            last.addnext(new)
            last = new
            data.append(new)
        for i, tr in enumerate(data, 1):               # nomor urut ulang
            tcs = tr.findall(q("tc"))
            ts = list(tcs[0].iter(q("t"))) if tcs else []
            if ts and re.fullmatch(r"\s*\d+\s*", "".join(t.text or "" for t in ts)):
                ts[0].text = str(i)
                for t in ts[1:]:
                    t.text = ""

    @staticmethod
    def _blank_row(tmpl):
        new = deepcopy(tmpl)
        for el in new.iter(etree.Element):
            for a in ("paraId", "textId"):
                el.attrib.pop(f"{{{W14_NS}}}{a}", None)
        for ci, tc in enumerate(new.findall(q("tc"))):
            if ci == 0:
                continue
            ps = tc.findall(q("p"))
            for extra in ps[1:]:
                tc.remove(extra)
            for p in ps[:1]:
                for ch in list(p):
                    if ch.tag != q("pPr"):
                        p.remove(ch)
        return new

    def _dyn_id(self, p):
        tc = next((a for a in p.iterancestors() if a.tag == q("tc")), None)
        return self._dyn_ids.get(tc.getparent()) if tc is not None else None

    def _add_cell_runs(self):
        """Sel tabel yang masih kosong (di baris yang sebagian sudah terisi) -> isian."""
        for tbl in self.root.iter(q("tbl")):
            rows = tbl.findall(q("tr"))
            if len(rows) < 2:
                continue
            head = [self._cell_text(c) for c in rows[0].findall(q("tc"))]
            for tr in rows[1:]:
                cells = tr.findall(q("tc"))
                texts = [self._cell_text(c) for c in cells]
                if not any(texts) or all(texts):
                    continue
                for ci, (tc, txt) in enumerate(zip(cells, texts)):
                    if txt or ci >= len(head) or not head[ci]:
                        continue
                    vm = tc.find(f"{q('tcPr')}/{q('vMerge')}")
                    if vm is not None:
                        continue
                    ps = tc.findall(q("p"))
                    if len(ps) != 1 or tc.find(q("tbl")) is not None:
                        continue
                    p = ps[0]
                    if any(True for _ in p.iter(q("t"))) or any(
                            c.tag not in (q("rPr"),) for r in p.iter(q("r")) for c in r):
                        continue
                    nr = etree.SubElement(p, q("r"))
                    ppr = p.find(q("pPr"))
                    mk = ppr.find(q("rPr")) if ppr is not None else None
                    if mk is not None:
                        nr.append(deepcopy(mk))
                    t = etree.SubElement(nr, q("t"))
                    t.text = ""
                    t.set(XML_SPACE, "preserve")
                    self.cell_ts.add(t)

    def _mark_editable(self):
        """Teks tetap yang juga boleh diubah: isi tabel isian (nama mata kuliah, SKS, ...)
        dan teks di blok tanda tangan (jabatan, nama pejabat)."""
        self.edit_ps = {}                       # paragraf -> label
        for tbl in self.root.iter(q("tbl")):
            rows = tbl.findall(q("tr"))
            if len(rows) < 2:
                continue
            hc = rows[0].findall(q("tc"))
            head = [self._cell_text(c) for c in hc]
            made = any(t in self.cell_ts for t in tbl.iter(q("t")))
            if not (made or (head and re.fullmatch(r"no\.?", head[0].strip(), re.I))):
                continue
            for tr in rows[1:]:
                for ci, tc in enumerate(tr.findall(q("tc"))):
                    ps = tc.findall(q("p"))
                    if len(ps) != 1 or tc.find(q("tbl")) is not None:
                        continue
                    txt = self._cell_text(tc)
                    if not txt or (ci == 0 and re.fullmatch(r"[\d.\s]+", txt)):
                        continue
                    if self._editable_ok(ps[0]):
                        self.edit_ps[ps[0]] = head[ci] if ci < len(head) and head[ci] else "Isi"
        start = re.compile(r"^\s*(Mengetahui|Menyetujui|Pemohon)\b|^[A-Za-z .]{2,30},\s*\.{4,}")
        on = not self.edit_ps and None          # blok tanda tangan hanya untuk formulir bertabel isian
        if on is None:
            return
        on = False
        for p in self.root.iter(q("p")):
            if any(a.tag == q("tc") for a in p.iterancestors()):
                continue
            txt = self._ptext(p).strip()
            if not on and start.match(txt):
                on = True
            if on and txt and self._editable_ok(p) and not TOKEN.search(txt) and not re.search(r"\w\s*:", txt):
                self.edit_ps[p] = "Teks tanda tangan"

    def _editable_ok(self, p):
        for r in p.iter(q("r")):
            for c in r:
                if c.tag not in (q("rPr"), q("t")):
                    return False
        return bool(self._ptext(p).strip())

    def _whole_parts(self, p, label):
        ts = list(p.iter(q("t")))
        V = "".join(t.text or "" for t in ts)
        lead = V[:len(V) - len(V.lstrip())]
        trail = V[len(V.rstrip()):]
        s = Slot("text", V.strip())
        s.prefill = s.preset = True
        s.label = label
        res, first = {}, True
        for t in ts:
            res[t] = []
        res[ts[0]] = ([lead] if lead else []) + [s] + ([trail] if trail else [])
        return res, s

    def _tabbed_neighbor(self, p):
        for nb in (p.getprevious(), p.getnext()):
            if nb is not None and nb.tag == q("p"):
                t = self._ptext_tabs(nb)
                if "\t" in t and re.search(r":\s*", t):
                    return True
        return False

    def _field_parts(self, p):
        """Isian gaya 'Label : nilai' (termasuk 'NIM : 123 <tab> NIDN :').
        Mengembalikan {w:t: [teks | Slot, ...]} atau None."""
        text, owner, tinfo = [], [], {}
        for r in p.iter(q("r")):
            for c in r:
                if c.tag == q("t"):
                    tinfo[c] = len(text)
                    for ch in c.text or "":
                        text.append(ch)
                        owner.append(c)
                elif c.tag == q("tab"):
                    text.append("\t")
                    owner.append(None)
                elif c.tag in (q("br"), q("object"), q("drawing"), q("pict")):
                    return None
        text = "".join(text)
        if not text.strip() or TOKEN.search(text):
            return None

        fields, i, n = [], 0, len(text)
        while i < n:
            c = text.find(":", i)
            if c < 0:
                break
            label = text[i:c].replace("\t", " ").strip()
            if (not label or len(label) > 32 or len(label.split()) > 5
                    or not re.search(r"[^\W\d_]", label) or text[c + 1:c + 2].isdigit()):
                break
            m = NEXT_FIELD.search(text, c + 1)
            end = m.start() if m else n
            fields.append((label, c + 1, end))
            i = end
        if not fields:
            return None

        ops = []                                   # (awal, akhir, slot)
        for label, rs, re_ in fields:
            R = text[rs:re_]
            vs = rs + len(R) - len(R.lstrip())
            ve = re_ - (len(R) - len(R.rstrip()))
            if ve <= vs:
                vs = ve = re_
            V = text[vs:ve]
            if "\t" in V or (V == "" and "\t" not in text and not self._tabbed_neighbor(p)):
                continue
            dates = list(DATE_RE.finditer(V))
            if dates:
                for k, m in enumerate(dates):
                    s = Slot("date", m.group(0))
                    s.iso = parse_id_date(m.group(0))
                    s.preset = True
                    s.label = f"{label} – " + (("awal", "akhir")[k] if len(dates) == 2 else f"tanggal {k + 1}")
                    ops.append((vs + m.start(), vs + m.end(), s))
            else:
                s = Slot("text", V)
                s.blank, s.prefill, s.preset, s.label = (V == ""), (V != ""), True, label
                ops.append((vs, ve, s))
        if not ops:
            return None

        starts, ends, removed = {}, {}, []
        for a, b, s in ops:
            if b > a:
                starts.setdefault(a, []).append(s)
                removed.append((a, b))
            else:
                j = a - 1
                while j >= 0 and owner[j] is None:
                    j -= 1
                if j >= 0:
                    ends.setdefault(owner[j], []).append(s)
        res = {}
        for t, g0 in tinfo.items():
            parts, buf = [], ""
            for k, ch in enumerate(t.text or ""):
                gi = g0 + k
                if gi in starts:
                    if buf:
                        parts.append(buf)
                        buf = ""
                    parts.extend(starts[gi])
                if any(a <= gi < b for a, b in removed):
                    continue
                buf += ch
            if buf:
                parts.append(buf)
            parts.extend(ends.get(t, []))
            res[t] = parts
        return res

    # ----- analisis: temukan semua slot ---------------------------------
    @staticmethod
    def _ptext(p):
        return "".join(t.text or "" for t in p.iter(q("t")))

    @staticmethod
    def _ptext_tabs(p):
        out = []
        for r in p.iter(q("r")):
            for c in r:
                if c.tag == q("t"):
                    out.append(c.text or "")
                elif c.tag == q("tab"):
                    out.append("\t")
        return "".join(out)

    @classmethod
    def _cell_text(cls, tc):
        txt = " ".join(cls._ptext(p) for p in tc.iter(q("p")))
        return re.sub(r"\.{3,}|…+", "", txt).strip()

    def analyze(self):
        self.tparts = {}
        self.para_ts = {}
        self.slots = []
        pmap = {}
        rowctr = {}
        prev_text, prev_raw, section, counter = "", "", "", 0
        paras = list(self.root.iter(q("p")))

        for pi, p in enumerate(paras):
            ts = list(p.iter(q("t")))
            has_br = p.find(f".//{q('br')}") is not None
            text_all = "".join(t.text or "" for t in ts)
            in_tbl = any(a.tag == q("tc") for a in p.iterancestors())

            # konteks tabel
            row_label = header = ""
            tc = next((a for a in p.iterancestors() if a.tag == q("tc")), None)
            if tc is not None:
                tr, tbl = tc.getparent(), tc.getparent().getparent()
                cells = tr.findall(q("tc"))
                rows = tbl.findall(q("tr"))
                idx = cells.index(tc)
                if idx > 0:
                    row_label = next((self._cell_text(c) for c in cells[:idx]
                                      if re.search(r"[^\W\d_]", self._cell_text(c))),
                                     self._cell_text(cells[0]))
                if tr is not rows[0]:
                    hc = rows[0].findall(q("tc"))
                    if idx < len(hc):
                        header = self._cell_text(hc[idx])

            over = None
            if p in self.edit_ps:
                over, es = self._whole_parts(p, self.edit_ps[p])
                es.full = tc is not None
            elif tc is None and not has_br:
                over = self._field_parts(p)
            pslots = []
            for t in ts:
                txt = t.text or ""
                if t in self.cell_ts:
                    parts = [Slot("text", "")]
                    parts[0].blank = parts[0].nospace = parts[0].full = True
                elif over is not None and t in over:
                    parts = over[t]
                    if not any(isinstance(x, Slot) for x in parts):
                        t.text = "".join(parts)
                        t.set(XML_SPACE, "preserve")
                        parts = [t.text] if t.text else []
                else:
                    parts = merge_dates(tokenize(txt))
                    if (has_br and not any(isinstance(x, Slot) for x in parts)
                            and txt.rstrip().endswith(":")):
                        blank = Slot("text", "")
                        blank.blank = True
                        parts.append(blank)
                seen = ""
                for x in parts:
                    if isinstance(x, str):
                        seen += x
                    else:
                        x.before = seen
                        pslots.append(x)
                self.tparts[t] = parts
            if pslots:
                self.para_ts[p] = ts

            # penomoran + label
            flat = [x for t in ts for x in self.tparts[t]]
            ptxt = text_all.strip()
            sig = bool(re.fullmatch(r"\(.*\)", ptxt))

            # "Cirebon, ........." -> pilih tanggal
            if (tc is None and len(pslots) == 1 and pslots[0].kind == "text"
                    and not pslots[0].blank and not pslots[0].preset
                    and re.fullmatch(r"[A-Za-z .]{2,30},\s*", pslots[0].before)
                    and ptxt == (pslots[0].before + pslots[0].orig).strip()):
                pslots[0].kind = "date"
            # "Mulai : ..... s.d. ....." -> dua pemilih tanggal (awal & akhir), tanpa isi awal
            if (tc is None and len(pslots) == 2
                    and all(x.kind == "text" and not x.blank and not x.preset for x in pslots)
                    and re.fullmatch(r"\s*(s\.?\s?d\.?|sampai(\s+dengan)?|-|–)\s*",
                                     pslots[1].before[len(pslots[0].before):] or "", re.I)):
                head = re.match(r"\s*([^:]{1,32}?)\s*(?:\t|:)", self._ptext_tabs(p))
                base = head.group(1).strip() if head else "Tanggal"
                for x, lab in zip(pslots, ("awal", "akhir")):
                    x.kind, x.preset, x.default_today = "date", True, False
                    x.label = f"{base} – {lab}"
            # beberapa titik-titik sejajar (kolom tanda tangan) -> nama kolom dari baris di atasnya
            if (tc is None and len(pslots) >= 2
                    and all(x.kind == "text" and not x.blank and not x.preset for x in pslots)):
                cols = [c.strip() for c in re.split(r"\t+| {3,}", prev_raw) if c.strip()]
                if len(cols) == len(pslots):
                    for x, c in zip(pslots, cols):
                        x.label = "Nama " + re.sub(r"[\s,:.]+$", "", c)
                        x.preset = True

            for si, s in enumerate(pslots):
                did = self._dyn_id(p)
                if did:        # id berdasarkan posisi (tabel, baris, isian) agar tak bergeser saat baris berubah
                    s.id = f"d{did[0]}_{did[1]}_{rowctr.get(did, 0)}"
                    rowctr[did] = rowctr.get(did, 0) + 1
                else:
                    s.id = f"{'c' if s.kind == 'check' else 'f'}{counter}"
                    counter += 1
                before = re.sub(r"[\s:.,;(]+$", "", s.before).strip()
                if s.kind == "check":
                    fi = flat.index(s)
                    follow = ""
                    for y in flat[fi + 1:]:
                        if isinstance(y, Slot):
                            break
                        follow += y
                    follow = follow.strip()
                    if follow:
                        s.label = follow
                    elif row_label:
                        s.label = row_label + (f" – {header}" if header else "")
                    else:
                        s.label = prev_text
                else:
                    if s.preset:
                        pass
                    elif s.kind == "date":
                        s.label = "Tanggal" + (f" ({section})" if section else "")
                    elif sig:
                        s.label = "Nama " + re.sub(r"[\s,:.]+$", "", prev_text)
                    elif row_label:
                        s.label = row_label + (f" – {header}" if header and s.nospace else "")
                    elif len(before) >= 2:
                        s.label = before[-60:]
                    else:
                        s.label = re.sub(r"[\s:.]+$", "", prev_text)
                    s.full = s.full or (not s.blank and s.kind == "text" and ptxt == s.orig)
                    low = s.label.lower().strip()
                    if low in ("nama lengkap", "nama") or (low.startswith("nama") and ("mahasiswa" in low or "pemohon" in low)):
                        s.link = "nama"
                    elif low == "nim":
                        s.link = "nim"
                s.label = s.label.strip()[:80]
                self.slots.append(s)

            pmap[p] = pslots
            if not in_tbl and re.match(r"^\s*[A-Z]\.\s+\S", text_all):
                section = text_all.strip()
            if re.sub(r"[.\s…]", "", text_all):
                prev_text = text_all.strip()
                prev_raw = self._ptext_tabs(p)

        # ---- grup centang saling-eksklusif ----
        for pi, p in enumerate(paras):
            checks = [s for s in pmap[p] if s.kind == "check"]
            if len(checks) >= 2:
                for s in checks:
                    s.group = f"p{pi}"
        for ri, tr in enumerate(self.root.iter(q("tr"))):
            checks = [s for p in tr.iter(q("p")) for s in pmap.get(p, []) if s.kind == "check"]
            if len(checks) >= 2 and not any(s.group for s in checks):
                for s in checks:
                    s.group = f"r{ri}"
        run, run_prev, last_text, n = [], "", "", 0

        def close():
            nonlocal run, n
            if len(run) >= 2 and run_prev.endswith(":") and not MULTI_CUE.search(run_prev):
                n += 1
                for s in run:
                    s.group = s.group or f"k{n}"
            run = []

        for p in paras:
            in_tbl = any(a.tag == q("tc") for a in p.iterancestors())
            txt = self._ptext(p).strip()
            checks = [s for s in pmap[p] if s.kind == "check"]
            if not in_tbl and len(checks) == 1 and txt.startswith(UNCHECKED):
                if not run:
                    run_prev = last_text
                run.append(checks[0])
            elif txt:
                close()
                last_text = txt

        close()

    # ----- helper ukuran/ lebar ------------------------------------------
    def _style_hp(self, p):
        """Ukuran font (half-point) dari gaya paragraf, atau default dokumen."""
        if self.styles is None:
            return self.default_hp
        ppr = p.find(q("pPr"))
        sid = None
        if ppr is not None and ppr.find(q("pStyle")) is not None:
            sid = ppr.find(q("pStyle")).get(q("val"))
        if sid in self._style_hp_cache:
            return self._style_hp_cache[sid]
        cur, hp = sid, None
        for _ in range(6):
            if cur is None:
                break
            st = self.styles.xpath(f"//w:style[@w:styleId='{cur}']", namespaces=NS)
            if not st:
                break
            sz = st[0].find(f"{q('rPr')}/{q('sz')}")
            if sz is not None and (sz.get(q("val")) or "").isdigit():
                hp = int(sz.get(q("val")))
                break
            bo = st[0].find(q("basedOn"))
            cur = bo.get(q("val")) if bo is not None else None
        hp = hp or self.default_hp
        self._style_hp_cache[sid] = hp
        return hp

    def _rfont(self, r, p):
        rpr = r.find(q("rPr"))
        hp, bold = None, False
        if rpr is not None:
            sz = rpr.find(q("sz"))
            if sz is not None and (sz.get(q("val")) or "").isdigit():
                hp = int(sz.get(q("val")))
            b = rpr.find(q("b"))
            bold = b is not None and b.get(q("val")) not in ("0", "false", "off")
        return (hp or self._style_hp(p)) / 2, bold

    def _cellw(self, tc, text_w):
        """Lebar sel (twips)."""
        tr = tc.getparent()
        tbl = tr.getparent()
        grid = [int(g.get(q("w"), 0) or 0) for g in tbl.findall(f"{q('tblGrid')}/{q('gridCol')}")]
        cells = tr.findall(q("tc"))

        def span(c):
            gs = c.find(f"{q('tcPr')}/{q('gridSpan')}")
            return int(gs.get(q("val"))) if gs is not None and (gs.get(q("val")) or "").isdigit() else 1

        if not grid or sum(grid) == 0:
            return text_w / max(len(cells), 1)
        idx = 0
        for c in cells:
            if c is tc:
                break
            idx += span(c)
        w = sum(grid[idx: idx + span(tc)])
        tot = sum(grid)
        target = tot
        tw_ = tbl.find(f"{q('tblPr')}/{q('tblW')}")
        if tw_ is not None:
            try:
                v, ty = int(tw_.get(q("w"), 0)), tw_.get(q("type"))
                if ty == "pct" and v:
                    target = v / 5000 * text_w
                elif ty == "dxa" and v:
                    target = v
            except ValueError:
                pass
        target = min(target, text_w)
        return w * target / tot

    def _avail_pt(self, p, text_w=None):
        """Lebar area tulis paragraf (pt), sudah dikurangi indent & margin sel."""
        text_w = text_w or self.text_width
        tc = next((a for a in p.iterancestors() if a.tag == q("tc")), None)
        width = text_w if tc is None else self._cellw(tc, text_w) - 216
        ind = p.find(f"{q('pPr')}/{q('ind')}")
        if ind is not None:
            for k in ("left", "start", "right", "end"):
                v = ind.get(q(k))
                if v and v.lstrip("-").isdigit():
                    width -= int(v)
        return max(width, 200) / 20

    # ----- isi nilai -> docx --------------------------------------------
    def _value(self, s, form):
        if s.kind == "check":
            return CHECKED if form.get(s.id) else UNCHECKED
        v = clean_value(form.get(s.id, ""))
        if s.kind == "date" and s.sub is None:
            try:
                d = date.fromisoformat(v)
            except ValueError:
                return s.orig
            return f"{d.day} {BULAN[d.month - 1]} {d.year}"
        if s.kind == "date":
            a, mid, b = s.sub
            try:
                d = date.fromisoformat(v)
            except ValueError:
                return a.orig + mid + b.orig
            if 2000 <= d.year <= 2099:
                return f"{d.day} {BULAN[d.month - 1]}{mid}{d.year % 100:02d}"
            return f"{d.day} {BULAN[d.month - 1]} {d.year}"
        if s.prefill:
            return v
        return v if v else s.orig

    def _pieces(self, t, form, fit_dots, size, bold, cont, st):
        """Pecah satu <w:t> berisi slot menjadi potongan teks / tab penuntun titik."""
        parts = self.tparts[t]
        res, buf, pos = [], "", st["pos"]
        lead_trim = False
        for i, x in enumerate(parts):
            if isinstance(x, str):
                if lead_trim and not buf:           # setelah titik-titik dibuang, rapikan spasi di depan
                    x = x.lstrip()
                lead_trim = False
                buf += x
                pos += tw(x, size, bold)
                continue
            if x.kind == "text" and x.prefill:
                v = clean_value(form[x.id]) if x.id in form else x.orig
                buf += v
                pos += tw(v, size, bold)
            elif x.kind == "text" and not x.blank:
                v = clean_value(form.get(x.id, ""))
                if not v:
                    if getattr(self, "drop_empty", False):      # titik-titik kosong dibuang
                        lead_trim = True
                        continue
                    buf += x.orig
                    pos += tw(x.orig, size, bold)
                    continue
                after = ""
                for y in parts[i + 1:]:
                    if isinstance(y, Slot):
                        break
                    after += y
                dot_w = (1.0 if "…" in x.orig else 0.25) * size * len(x.orig)
                field = min(dot_w, cont - pos - tw(after, size, bold) - 4)
                vs, tab = None, False
                if fit_dots and field >= 24:
                    vw, target = tw(v, size, bold), field - 3
                    if vw <= target:
                        tab = True
                    else:                                  # terlalu panjang -> kecilkan font
                        s2 = int(size * target / vw * 2) / 2
                        if s2 >= max(7.5, size * 0.6):
                            vs, tab = s2, True
                if tab:
                    if buf:
                        res.append(("text", buf, None))
                        buf = ""
                    res.append(("text", v, vs))
                    res.append(("tab", int((pos + field) * 20)))
                    st["tabs"].append(int((pos + field) * 20))
                    pos += field
                else:
                    buf += v
                    pos += tw(v, size, bold)
            elif x.blank:
                v = clean_value(form.get(x.id, ""))
                if v:
                    buf = v if x.nospace else buf.rstrip() + " " + v
                    pos += tw(v if x.nospace else " " + v, size, bold)
            else:
                v = self._value(x, form)
                buf += v
                pos += tw(v, size, bold)
        if buf:
            res.append(("text", buf, None))
        st["pos"] = pos
        return res

    @staticmethod
    def _next_stop(st):
        """Posisi (pt) tab-stop berikutnya setelah posisi tulis saat ini."""
        stops = sorted(st["stops"] + [x / 20 for x in st["tabs"]])
        nxt = [x for x in stops if x > st["pos"] + 0.5]
        return nxt[0] if nxt else (int(st["pos"] // 36) + 1) * 36

    def _fill_run(self, r, p, form, fit_dots, cont, st):
        size, bold = self._rfont(r, p)
        kids = [c for c in r if c.tag != q("rPr")]
        rpr = r.find(q("rPr"))
        has = any(c.tag == q("t") and any(isinstance(x, Slot) for x in self.tparts.get(c, []))
                  for c in kids)
        if not has:
            for c in kids:
                if c.tag == q("t"):
                    st["pos"] += tw(c.text or "", size, bold)
                elif c.tag == q("br") and c.get(q("type")) != "page":
                    st["pos"] = 0.0
                elif c.tag == q("tab"):
                    st["pos"] = self._next_stop(st)
            return
        out, cur = [], []                       # out: [(ukuran_baru|None, [elemen])]

        def flush():
            nonlocal cur
            if cur:
                out.append((None, cur))
                cur = []

        def mk_t(text):
            e = etree.Element(q("t"))
            e.text = text
            e.set(XML_SPACE, "preserve")
            return e

        for c in kids:
            if c.tag == q("t") and any(isinstance(x, Slot) for x in self.tparts.get(c, [])):
                for kind, a, *b in self._pieces(c, form, fit_dots, size, bold, cont, st):
                    if kind == "text" and b[0]:
                        flush()
                        out.append((b[0], [mk_t(a)]))
                    elif kind == "text":
                        cur.append(mk_t(a))
                    else:                        # tab penuntun titik: run sendiri agar titik berukuran normal
                        flush()
                        out.append((None, [etree.Element(q("tab"))]))
            else:
                if c.tag == q("br") and c.get(q("type")) != "page":
                    st["pos"] = 0.0
                elif c.tag == q("tab"):
                    st["pos"] = self._next_stop(st)
                cur.append(c)
        flush()
        parent, idx = r.getparent(), r.getparent().index(r)
        parent.remove(r)
        for n, (szov, kids2) in enumerate(out):
            nr = etree.Element(q("r"))
            if rpr is not None:
                nr.append(deepcopy(rpr))
            if szov:
                set_sz(nr, int(szov * 2))
            for k in kids2:
                nr.append(k)
            parent.insert(idx + n, nr)

    def _add_tabs(self, p, positions):
        ppr = p.find(q("pPr"))
        if ppr is None:
            ppr = etree.Element(q("pPr"))
            p.insert(0, ppr)
        tabs = ppr.find(q("tabs"))
        if tabs is None:
            tabs = etree.Element(q("tabs"))
            pos_i = len(ppr)
            for i, c in enumerate(ppr):
                if etree.QName(c).localname in PPR_AFTER_TABS:
                    pos_i = i
                    break
            ppr.insert(pos_i, tabs)
        have = {tb.get(q("pos")) for tb in tabs}
        for pos in sorted(set(positions)):
            if str(pos) in have:
                continue
            tb = etree.SubElement(tabs, q("tab"))
            tb.set(q("val"), "left")
            tb.set(q("leader"), "dot")
            tb.set(q("pos"), str(pos))

    def fill(self, form, fit_dots=True, drop_empty=False):
        self.drop_empty = drop_empty
        for p, ts in self.para_ts.items():
            cont = self._avail_pt(p)
            ppr = p.find(q("pPr"))
            stops = [int(tb.get(q("pos"))) / 20 for tb in
                     (ppr.findall(f"{q('tabs')}/{q('tab')}") if ppr is not None else [])
                     if tb.get(q("val")) != "clear" and (tb.get(q("pos")) or "").isdigit()]
            st = {"pos": 0.0, "tabs": [], "stops": stops}
            for r in list(p.iter(q("r"))):
                self._fill_run(r, p, form, fit_dots, cont, st)
            if st["tabs"]:
                self._add_tabs(p, st["tabs"])

    # ----- pas satu halaman ------------------------------------------------
    def _tbl_spacing(self, tbl):
        """Spasi paragraf dari gaya tabel (mengalahkan default dokumen)."""
        if self.styles is None:
            return None
        pr = tbl.find(q("tblPr"))
        sid = pr.find(q("tblStyle")).get(q("val")) if pr is not None and pr.find(q("tblStyle")) is not None else None
        for _ in range(6):
            if not sid:
                return None
            st = self.styles.xpath(f"//w:style[@w:styleId='{sid}']", namespaces=NS)
            if not st:
                return None
            sp = st[0].find(f"{q('pPr')}/{q('spacing')}")
            if sp is not None:
                return sp
            bo = st[0].find(q("basedOn"))
            sid = bo.get(q("val")) if bo is not None else None
        return None

    def _pinfo(self, p):
        ppr = p.find(q("pPr"))
        sp = ppr.find(q("spacing")) if ppr is not None else None
        tc = next((a for a in p.iterancestors() if a.tag == q("tc")), None)
        tsp = self._tbl_spacing(tc.getparent().getparent()) if tc is not None else None

        def num(e, k, default):
            v = e.get(q(k)) if e is not None else None
            if v is None or not v.lstrip("-").isdigit():
                v = tsp.get(q(k)) if tsp is not None else None
            return int(v) if v is not None and v.lstrip("-").isdigit() else default

        stops = []
        if ppr is not None:
            for tb in ppr.findall(f"{q('tabs')}/{q('tab')}"):
                if tb.get(q("val")) != "clear" and (tb.get(q("pos")) or "").isdigit():
                    stops.append(int(tb.get(q("pos"))) / 20)
        stops.sort()
        lines, sizes = [0.0], []
        for r in p.iter(q("r")):
            size, bold = self._rfont(r, p)
            for c in r:
                if c.tag == q("t") and (c.text or ""):
                    lines[-1] += tw(c.text, size, bold)
                    sizes.append(size)
                elif c.tag == q("br") and c.get(q("type")) != "page":
                    lines.append(0.0)
                elif c.tag == q("tab"):
                    nxt = [x for x in stops if x > lines[-1] + 0.5]
                    lines[-1] = nxt[0] if nxt else (int(lines[-1] // 36) + 1) * 36
        if not sizes:
            mk = ppr.find(f"{q('rPr')}/{q('sz')}") if ppr is not None else None
            hp = num(mk, "val", self._style_hp(p)) if mk is not None else self._style_hp(p)
            sizes = [hp / 2]
        return {
            "lines": lines, "s0": max(sizes),
            "empty": not any((c.text or "").strip() for c in p.iter(q("t"))),
            "before": num(sp, "before", self.defaults["before"]),
            "after": num(sp, "after", self.defaults["after"]),
            "line": num(sp, "line", self.defaults["line"]),
            "rule": (sp.get(q("lineRule")) if sp is not None else None) or "auto",
            "bauto": sp is not None and sp.get(q("beforeAutospacing")) in ("1", "true"),
            "aauto": sp is not None and sp.get(q("afterAutospacing")) in ("1", "true"),
        }

    def _pheight(self, inf, avail, f, g):
        s0 = inf["s0"]
        s = max(min(s0, 8.0), s0 * f)
        fe = s / s0
        n = sum(max(1, math.ceil(w * fe / (avail * 1.04) - 1e-6)) for w in inf["lines"])
        if inf["rule"] == "auto":
            lh = s * 1.15 * inf["line"] / 240
        elif inf["rule"] == "exact":
            lh = inf["line"] / 20 * f
        else:
            lh = max(inf["line"] / 20 * f, s * 1.15)
        if inf["empty"]:
            lh *= max(g, 0.5)
        before = 0 if inf["bauto"] else inf["before"] / 20
        after = 14 if inf["aauto"] else inf["after"] / 20
        return n * lh + (before + after) * g

    def _blocks_h(self, parent, text_w, f, g, infos, avail=None):
        total = 0.0
        for ch in parent:
            if ch.tag == q("p"):
                total += self._pheight(infos[ch], avail or self._avail_pt(ch, text_w), f, g)
            elif ch.tag == q("tbl"):
                for tr in ch.findall(q("tr")):
                    rowh = 0.0
                    trh = tr.find(f"{q('trPr')}/{q('trHeight')}")
                    if trh is not None and (trh.get(q("val")) or "").isdigit():
                        rowh = int(trh.get(q("val"))) / 20 * f
                    for tc in tr.findall(q("tc")):
                        cw = (self._cellw(tc, text_w) - 216) / 20
                        rowh = max(rowh, self._blocks_h(tc, text_w, f, g, infos, cw) + 1)
                    total += rowh
            elif ch.tag == q("sdt"):
                c = ch.find(q("sdtContent"))
                if c is not None:
                    total += self._blocks_h(c, text_w, f, g, infos, avail)
        return total

    def fit_one_page(self):
        """Padatkan dokumen agar muat satu halaman. Mengembalikan ringkasan pengaturan."""
        root = self.root
        # hapus pemisah halaman paksa
        for br in list(root.iter(q("br"))):
            if br.get(q("type")) == "page":
                br.getparent().remove(br)
        for tag in ("pageBreakBefore", "lastRenderedPageBreak"):
            for e in list(root.iter(q(tag))):
                e.getparent().remove(e)

        body = root.find(q("body"))
        infos = {p: self._pinfo(p) for p in root.iter(q("p"))}
        pg = self.pg
        orig = (pg["top"], pg["bottom"], pg["left"], pg["right"])
        cap = lambda lim: tuple(min(x, lim) for x in orig)

        cands = [(1.0, round(1 - 0.05 * i, 2), orig) for i in range(0, 15)]       # 1) rapatkan spasi
        cands += [(1.0, 0.3, cap(850))]                                             # 2) margin 1,5 cm
        cands += [(f / 100, 0.3, cap(850)) for f in (97, 94, 91, 88, 85, 82)]       # 3) kecilkan font
        cands += [(f / 100, 0.25, cap(720)) for f in (85, 82, 79, 76, 73, 70, 67, 64)]

        chosen = cands[-1]
        for f, g, mar in cands:
            text_w = pg["w"] - mar[2] - mar[3]
            avail_h = (pg["h"] - mar[0] - mar[1]) / 20
            if self._blocks_h(body, text_w, f, g, infos) <= avail_h * 0.95:
                chosen = (f, g, mar)
                break
        f, g, mar = chosen
        self._apply_compact(f, g, mar, infos)
        return {"font": f, "spacing": g, "margins": mar}

    def _apply_compact(self, f, g, mar, infos):
        root = self.root
        if self.sect is not None:
            pm = self.sect.find(q("pgMar"))
            if pm is not None:
                for k, v in zip(("top", "bottom", "left", "right"), mar):
                    pm.set(q(k), str(v))

        if f < 1:
            for r in root.iter(q("r")):
                p = next((a for a in r.iterancestors() if a.tag == q("p")), None)
                hp = int(self._rfont(r, p)[0] * 2) if p is not None else self.default_hp
                set_sz(r, max(min(hp, 16), int(round(hp * f))))
            for trh in root.iter(q("trHeight")):
                if (trh.get(q("val")) or "").isdigit():
                    trh.set(q("val"), str(int(int(trh.get(q("val"))) * f)))

        # tanda paragraf (menentukan tinggi baris kosong); baris kosong ikut dirapatkan
        if f < 1 or g < 1:
            for p in root.iter(q("p")):
                inf = infos.get(p) or self._pinfo(p)
                factor = f * (max(g, 0.5) if inf["empty"] and g < 1 else 1.0)
                if factor >= 1:
                    continue
                ppr = p.find(q("pPr"))
                if ppr is None:
                    ppr = etree.Element(q("pPr"))
                    p.insert(0, ppr)
                mk = ppr.find(q("rPr"))
                sz = mk.find(q("sz")) if mk is not None else None
                hp = int(sz.get(q("val"))) if sz is not None and (sz.get(q("val")) or "").isdigit() \
                    else self._style_hp(p)
                if mk is None:
                    mk = etree.Element(q("rPr"))
                    idx = len(ppr)
                    for i, c in enumerate(ppr):
                        if etree.QName(c).localname in ("sectPr", "pPrChange"):
                            idx = i
                            break
                    ppr.insert(idx, mk)
                set_sz(mk, max(min(hp, 16) if not inf["empty"] else 4, int(round(hp * factor))), cs=False)

        if g < 1:
            for p in root.iter(q("p")):
                inf = infos.get(p) or self._pinfo(p)
                ppr = p.find(q("pPr"))
                if ppr is None:
                    ppr = etree.Element(q("pPr"))
                    p.insert(0, ppr)
                sp = ppr.find(q("spacing"))
                if sp is None:
                    sp = etree.Element(q("spacing"))
                    idx = len(ppr)
                    for i, c in enumerate(ppr):
                        if etree.QName(c).localname in PPR_AFTER_SPACING:
                            idx = i
                            break
                    ppr.insert(idx, sp)
                for a in ("beforeAutospacing", "afterAutospacing", "beforeLines", "afterLines"):
                    if q(a) in sp.attrib:
                        del sp.attrib[q(a)]
                sp.set(q("before"), str(int((0 if inf["bauto"] else inf["before"]) * g)))
                sp.set(q("after"), str(int((280 if inf["aauto"] else inf["after"]) * g)))
                if inf["rule"] in ("exact", "atLeast") and f < 1:
                    sp.set(q("line"), str(int(inf["line"] * f)))

    def _drop_unused_parts(self, zin, xml):
        """Kotak ActiveX sudah diganti ☐/☑ -> buang relasi & part ActiveX yang tak terpakai.
        Mengembalikan (daftar nama part dibuang, {nama: isi baru})."""
        if not self.removed_rids:
            return set(), {}
        rels_name = "word/_rels/document.xml.rels"
        if rels_name not in zin.namelist():
            return set(), {}
        used = set(re.findall(rb'r:(?:id|embed|link|pict)="([^"]+)"', xml))
        used = {u.decode() for u in used}
        rels = etree.fromstring(zin.read(rels_name))
        drop, new = set(), {}
        for rel in list(rels):
            rid = rel.get("Id")
            if rid in self.removed_rids and rid not in used:
                tgt = rel.get("Target") or ""
                if rel.get("TargetMode") != "External" and not tgt.startswith("/"):
                    path = os.path.normpath("word/" + tgt).replace("\\", "/")
                    drop.add(path)
                    sub = f"{os.path.dirname(path)}/_rels/{os.path.basename(path)}.rels"
                    if sub in zin.namelist():
                        drop.add(sub)
                        for r2 in etree.fromstring(zin.read(sub)):
                            t2 = os.path.normpath(os.path.dirname(path) + "/" + (r2.get("Target") or ""))
                            drop.add(t2.replace("\\", "/"))
                rels.remove(rel)
        # part yang masih dipakai relasi lain jangan ikut dibuang
        for rel in rels:
            tgt = rel.get("Target") or ""
            drop.discard(os.path.normpath("word/" + tgt).replace("\\", "/"))
        new[rels_name] = etree.tostring(rels, xml_declaration=True, encoding="UTF-8", standalone=True)
        ct_name = "[Content_Types].xml"
        if ct_name in zin.namelist():
            ct = etree.fromstring(zin.read(ct_name))
            for e in list(ct):
                if e.tag.endswith("}Override") and (e.get("PartName") or "").lstrip("/") in drop:
                    ct.remove(e)
            new[ct_name] = etree.tostring(ct, xml_declaration=True, encoding="UTF-8", standalone=True)
        return drop, new

    def to_docx(self):
        xml = etree.tostring(self.root, xml_declaration=True, encoding="UTF-8", standalone=True)
        out = io.BytesIO()
        with zipfile.ZipFile(io.BytesIO(self.data)) as zin, \
                zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
            drop, repl = self._drop_unused_parts(zin, xml)
            for info in zin.infolist():
                n = info.filename
                if n.startswith("/") or ".." in n.split("/") or n in drop:
                    continue
                if n == "word/document.xml":
                    zout.writestr(n, xml)
                elif n in repl:
                    zout.writestr(n, repl[n])
                else:
                    zout.writestr(info, zin.read(n), compress_type=zipfile.ZIP_DEFLATED)
        return out.getvalue()

    # ----- render HTML (pratinjau yang bisa diisi) -----------------------
    def render_html(self):
        out = []
        self._children(self.root.find(q("body")), out)
        return "".join(out)

    def _children(self, parent, out):
        for ch in parent:
            if ch.tag == q("p"):
                self._p(ch, out)
            elif ch.tag == q("tbl"):
                self._tbl(ch, out)
            elif ch.tag == q("sdt"):
                c = ch.find(q("sdtContent"))
                if c is not None:
                    self._children(c, out)

    def _p(self, p, out):
        ppr = p.find(q("pPr"))
        st = []
        if ppr is not None:
            jc = ppr.find(q("jc"))
            if jc is not None:
                st.append("text-align:" + {"both": "justify", "center": "center", "right": "right",
                                           "end": "right"}.get(jc.get(q("val")), "left"))
            ind = ppr.find(q("ind"))
            if ind is not None:
                for k, css in (("left", "padding-left"), ("start", "padding-left"),
                               ("firstLine", "text-indent")):
                    v = ind.get(q(k))
                    if v and v.lstrip("-").isdigit():
                        st.append(f"{css}:{int(v) / 15:.0f}px")
        d = dict(self.defaults)
        sp = ppr.find(q("spacing")) if ppr is not None else None
        auto_b = auto_a = False
        rule = "auto"
        if sp is not None:
            for k in ("before", "after", "line"):
                v = sp.get(q(k))
                if v is not None and v.lstrip("-").isdigit():
                    d[k] = int(v)
            auto_b = sp.get(q("beforeAutospacing")) in ("1", "true")
            auto_a = sp.get(q("afterAutospacing")) in ("1", "true")
            rule = sp.get(q("lineRule")) or "auto"
        before = 14 if auto_b else d["before"] / 20
        after = 14 if auto_a else d["after"] / 20
        lh = f"{d['line'] / 240 * 1.15:.2f}" if rule == "auto" else f"{d['line'] / 20:.1f}pt"
        st.append(f"margin:{before:.1f}pt 0 {after:.1f}pt")
        st.append(f"line-height:{lh}")
        stops = sorted(int(tb.get(q("pos"))) / 20 for tb in (ppr.findall(f"{q('tabs')}/{q('tab')}")
                       if ppr is not None else []) if tb.get(q("val")) != "clear" and (tb.get(q("pos")) or "").isdigit())
        left = 0.0
        if ppr is not None and ppr.find(q("ind")) is not None:
            v = ppr.find(q("ind")).get(q("left")) or ppr.find(q("ind")).get(q("start")) or ""
            left = int(v) / 20 if v.lstrip("-").isdigit() else 0.0
        tst = {"pos": left, "stops": stops, "avail": self._avail_pt(p) + left}
        inner = "".join(self._run(r, p, tst) for r in p.iter(q("r")))
        if not inner.strip():
            inner = "&nbsp;"
        out.append(f'<p style="{";".join(st)}">{inner}</p>')

    def _run(self, r, p=None, tst=None):
        rpr = r.find(q("rPr"))
        st = []
        if rpr is not None:
            def on(tag):
                e = rpr.find(q(tag))
                return e is not None and e.get(q("val")) not in ("0", "false", "off")
            if on("b"):
                st.append("font-weight:bold")
            if on("i"):
                st.append("font-style:italic")
            u = rpr.find(q("u"))
            if u is not None and u.get(q("val")) not in (None, "none"):
                st.append("text-decoration:underline")
            sz = rpr.find(q("sz"))
            if sz is not None and (sz.get(q("val")) or "").isdigit():
                st.append(f"font-size:{int(sz.get(q('val'))) / 2}pt")
            rf = rpr.find(q("rFonts"))
            if rf is not None and rf.get(q("ascii")):
                st.append(f"font-family:'{rf.get(q('ascii'))}','Times New Roman',serif")
        size, bold = self._rfont(r, p) if p is not None else (12, False)
        buf = []
        for c in r:
            if c.tag == q("t"):
                buf.append(self._t(c))
                if tst is not None:
                    tst["pos"] += self._t_width(c, size, bold, tst["avail"] - tst["pos"])
            elif c.tag == q("br") and c.get(q("type")) != "page":
                buf.append("<br>")
                if tst is not None:
                    tst["pos"] = 0.0
            elif c.tag == q("tab"):
                if tst is None:
                    buf.append("&emsp;")
                else:
                    nxt = [x for x in tst["stops"] if x > tst["pos"] + 0.5]
                    stop = nxt[0] if nxt else (int(tst["pos"] // 36) + 1) * 36
                    buf.append(f'<span style="display:inline-block;width:{(stop - tst["pos"]) * 1.333:.0f}px"></span>')
                    tst["pos"] = stop
        return f'<span style="{";".join(st)}">{"".join(buf)}</span>' if st else "".join(buf)

    def _slot_width(self, x, size, room):
        """Perkiraan lebar isian (pt) di pratinjau, untuk menghitung posisi tab."""
        if x.kind == "date":
            return 11.5 * size * 0.9
        if x.kind == "check":
            return size * 1.4
        if x.full:
            return max(room, 40)
        return min(self._field_ch(x) * size * 0.5, max(room, 40))

    def _t_width(self, t, size, bold, room):
        w = 0.0
        for x in self.tparts.get(t, []):
            w += tw(x, size, bold) if isinstance(x, str) else self._slot_width(x, size, room - w)
        return w

    @staticmethod
    def _field_ch(s):
        if s.prefill:
            return min(max(len(s.orig) + 3, 14), 64)
        if s.blank:
            return 18
        return min(max(len(s.orig) * 0.5, 14), 40)

    def _t(self, t):
        res = []
        for x in self.tparts.get(t, []):
            res.append(H.escape(x) if isinstance(x, str) else self._input(x))
        return "".join(res)

    def _input(self, s):
        lab = H.escape(s.label, quote=True)
        if s.kind == "check":
            return (f'<input type="checkbox" class="cb" name="{s.id}" '
                    f'data-group="{s.group or ""}" title="{lab}" aria-label="{lab}">')
        if s.kind == "date":
            dv = s.iso or (date.today().isoformat() if s.default_today else "")
            return (f'<input type="date" class="fld date" name="{s.id}" data-def="{"" if s.iso else dv}" '
                    f'value="{dv}" title="{lab}" aria-label="{lab}">')
        ph = H.escape(s.label.split(" – ")[-1], quote=True) if s.nospace else lab
        if s.full:
            style, cls = "", "fld full"
        elif s.prefill and len(s.orig) > 45:
            style, cls = "width:100%", "fld full"
        else:
            style, cls = f"width:{self._field_ch(s):.0f}ch", "fld"
        val = f' value="{H.escape(s.orig, quote=True)}"' if s.prefill else ""
        return (f'<input type="text" class="{cls}" name="{s.id}" placeholder="{ph}"{val} '
                f'title="{lab}" aria-label="{lab}" maxlength="300" autocomplete="off" '
                f'data-link="{s.link or ""}" style="{style}">')

    def _has_borders(self, tblpr):
        def check(pr):
            b = pr.find(q("tblBorders")) if pr is not None else None
            if b is None:
                return None
            return any(c.get(q("val")) not in ("nil", "none") for c in b)

        r = check(tblpr)
        if r is not None:
            return r
        sid = tblpr.find(q("tblStyle")) if tblpr is not None else None
        sid = sid.get(q("val")) if sid is not None else None
        for _ in range(5):
            if not sid or self.styles is None:
                break
            st = self.styles.xpath(f"//w:style[@w:styleId='{sid}']", namespaces=NS)
            if not st:
                break
            r = check(st[0].find(q("tblPr")))
            if r is not None:
                return r
            bo = st[0].find(q("basedOn"))
            sid = bo.get(q("val")) if bo is not None else None
        return False

    def _tbl(self, tbl, out):
        grid = [int(g.get(q("w"), 0) or 0) for g in tbl.findall(f"{q('tblGrid')}/{q('gridCol')}")]
        tot = sum(grid) or 1
        pr = tbl.find(q("tblPr"))
        tw = pr.find(q("tblW")) if pr is not None else None
        wpct = min(tot / self.text_width * 100, 100)
        if tw is not None:
            try:
                v, ty = int(tw.get(q("w"), 0)), tw.get(q("type"))
                if ty == "pct" and v:
                    wpct = min(v / 50, 100)
                elif ty == "dxa" and v:
                    wpct = min(v / self.text_width * 100, 100)
            except ValueError:
                pass
        cls = "doc bordered" if self._has_borders(pr) else "doc"
        dyn_i = next((i for i, d in enumerate(self.dyn) if d is tbl), None)
        data_ids = {id(r) for r in self._data_rows(tbl)} if dyn_i is not None else set()
        attr = ""
        if dyn_i is not None:
            cls += " dyn"
            attr = f' data-dyn="{dyn_i}"'
        out.append(f'<table class="{cls}"{attr} style="width:{wpct:.1f}%"><colgroup>')
        for g in grid:
            out.append(f'<col style="width:{g / tot * 100:.1f}%">')
        out.append("</colgroup>")
        for tr in tbl.findall(q("tr")):
            out.append('<tr class="dr">' if id(tr) in data_ids else "<tr>")
            for tc in tr.findall(q("tc")):
                tcpr = tc.find(q("tcPr"))
                span, va = 1, "top"
                if tcpr is not None:
                    gs = tcpr.find(q("gridSpan"))
                    if gs is not None and (gs.get(q("val")) or "").isdigit():
                        span = int(gs.get(q("val")))
                    v = tcpr.find(q("vAlign"))
                    if v is not None:
                        va = {"center": "middle", "bottom": "bottom"}.get(v.get(q("val")), "top")
                out.append(f'<td colspan="{span}" style="vertical-align:{va}">')
                self._children(tc, out)
                out.append("</td>")
            out.append("</tr>")
        out.append("</table>")
        if dyn_i is not None:
            out.append(f'<div class="dctl" data-for="{dyn_i}">'
                       '<button type="button" class="btn sec dadd">＋ Tambah baris</button>'
                       '<button type="button" class="btn ghost ddel">− Kurangi baris</button>'
                       '<span class="dcnt"></span></div>')


# --------------------------------------------------------------------------
# Penyimpanan sementara
# --------------------------------------------------------------------------
# --------------------------------------------------------------------------
# Halaman
# --------------------------------------------------------------------------
HOME = """<!doctype html>
<html lang="id"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Isi Formulir Word</title>
<style>
:root{--bg:#f4f6fb;--card:#fff;--ink:#1b2333;--mut:#667085;--line:#d9e0ee;--acc:#2f5fe3;--acc-d:#2349b8;--acc-l:#eef3ff;--err:#b42318;--errbg:#fef3f2}
@media(prefers-color-scheme:dark){:root{--bg:#0f1420;--card:#171e2e;--ink:#e8ecf5;--mut:#9aa5bd;--line:#2a3550;--acc:#6c93ff;--acc-d:#8aa9ff;--acc-l:#1b2644;--err:#ffb4a8;--errbg:#3a1a17}}
*{box-sizing:border-box}
body{margin:0;font-family:system-ui,-apple-system,"Segoe UI",Roboto,sans-serif;background:var(--bg);color:var(--ink);
min-height:100vh;display:flex;align-items:center;justify-content:center;padding:24px 16px}
.card{background:var(--card);max-width:580px;width:100%;border-radius:20px;padding:36px 32px 28px;
box-shadow:0 1px 2px rgba(16,24,40,.06),0 12px 40px rgba(16,24,40,.10);border:1px solid var(--line)}
.logo{width:44px;height:44px;border-radius:12px;background:var(--acc-l);color:var(--acc);display:grid;place-items:center;margin-bottom:14px}
h1{margin:0 0 6px;font-size:1.55rem;letter-spacing:-.01em}
p.sub{margin:0 0 22px;color:var(--mut);line-height:1.55}
.err{display:flex;gap:8px;background:var(--errbg);color:var(--err);padding:11px 14px;border-radius:10px;margin-bottom:16px;font-size:.93rem;line-height:1.45}
.drop{display:block;border:2px dashed var(--line);border-radius:14px;padding:34px 18px;text-align:center;cursor:pointer;
background:var(--acc-l);transition:border-color .15s,background .15s,transform .15s}
.drop:hover,.drop.over,.drop:focus-within{border-color:var(--acc)}
.drop.over{transform:scale(1.01)}
.drop svg{color:var(--acc);margin-bottom:8px}
.drop b{display:block;font-size:1.05rem;margin-bottom:4px}
.drop span{color:var(--mut);font-size:.88rem}
input[type=file]{position:absolute;width:1px;height:1px;opacity:0;pointer-events:none}
.steps{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin:22px 0 0;padding:0;list-style:none;counter-reset:s}
.steps li{counter-increment:s;font-size:.86rem;color:var(--mut);line-height:1.4;background:var(--bg);border-radius:10px;padding:10px 10px 10px 38px;position:relative}
.steps li::before{content:counter(s);position:absolute;left:10px;top:10px;width:20px;height:20px;border-radius:50%;background:var(--acc);color:#fff;
font-size:.75rem;font-weight:700;display:grid;place-items:center}
.tpl{margin-top:26px;border-top:1px solid var(--line);padding-top:20px}
.tpl h2{font-size:1.02rem;margin:0 0 2px}
.sub2{margin:0 0 12px;color:var(--mut);font-size:.88rem}
.t-item{display:flex;gap:12px;align-items:center;flex-wrap:wrap;border:1px solid var(--line);border-radius:12px;padding:12px;margin-bottom:10px}
.t-ico{width:38px;height:38px;border-radius:9px;background:#2b579a;color:#fff;font-weight:800;display:grid;place-items:center;flex:none}
.t-txt{flex:1;min-width:200px;display:flex;flex-direction:column;gap:2px;font-size:.92rem}
.t-txt span{color:var(--mut);font-size:.82rem;line-height:1.4}
.t-act{display:flex;gap:8px;flex-wrap:wrap}
.t-act a{text-decoration:none;font-size:.85rem;font-weight:600;padding:8px 12px;border-radius:8px;white-space:nowrap}
.t-act .b1{background:var(--acc);color:#fff}.t-act .b1:hover{background:var(--acc-d)}
.t-act .b2{background:var(--acc-l);color:var(--acc)}
.row{display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap;margin-top:20px;font-size:.85rem;color:var(--mut)}
a.sample{color:var(--acc);text-decoration:none;font-weight:600;font-size:.92rem}
a.sample:hover{text-decoration:underline}
.busy{display:none;position:fixed;inset:0;background:rgba(15,20,32,.55);z-index:50;align-items:center;justify-content:center;flex-direction:column;gap:14px;color:#fff;font-weight:600}
.busy.on{display:flex}
.spin{width:38px;height:38px;border-radius:50%;border:4px solid rgba(255,255,255,.3);border-top-color:#fff;animation:sp .8s linear infinite}
@keyframes sp{to{transform:rotate(360deg)}}
@media(max-width:520px){.card{padding:26px 18px 22px}.steps{grid-template-columns:1fr}}
</style></head><body>
<main class="card">
<div class="logo"><svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5M9 13h6M9 17h4"/></svg></div>
<h1>Isi Formulir Word</h1>
<p class="sub">Upload formulir <b>.docx</b>, isi atau edit di browser, lalu unduh kembali sebagai file Word dengan format yang tetap sama.</p>
{% if error %}<div class="err" role="alert"><span>⚠️</span><span>{{ error }}</span></div>{% endif %}
<form id="up" method="post" action="{{ url_for('upload') }}" enctype="multipart/form-data">
<label class="drop" id="drop" for="file">
<svg width="34" height="34" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 16V4m0 0L7 9m5-5 5 5"/><path d="M4 16v3a1 1 0 0 0 1 1h14a1 1 0 0 0 1-1v-3"/></svg>
<b id="dt">Klik atau seret file .docx ke sini</b><span id="ds">Maksimal 3 MB</span>
</label>
<input type="file" id="file" name="file" accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document">
</form>
<ol class="steps"><li>Upload file Word</li><li>Isi &amp; centang isiannya</li><li>Klik <b>Unduh Word</b></li></ol>
{% if templates %}
<section class="tpl"><h2>Belum punya file formulirnya?</h2>
<p class="sub2">Pakai formulir bawaan: isi langsung di sini, atau unduh file kosongnya.</p>
{% for key, title, desc in templates %}
<div class="t-item"><div class="t-ico">W</div>
<div class="t-txt"><b>{{ title }}</b><span>{{ desc }}</span></div>
<div class="t-act"><a class="b1" href="{{ url_for('template_fill', key=key) }}">Isi sekarang</a>
<a class="b2" href="{{ url_for('template_download', key=key) }}">⬇ Unduh kosong</a></div></div>
{% endfor %}</section>{% endif %}
<div class="row"><span>🔒 File tidak disimpan di server.</span></div>
</main>
<div class="busy" id="busy"><div class="spin"></div><div>Membaca formulir…</div></div>
<script>
const f=document.getElementById('file'),d=document.getElementById('drop'),form=document.getElementById('up'),busy=document.getElementById('busy');
function go(){
  const file=f.files[0]; if(!file)return;
  const t=document.getElementById('dt'),s=document.getElementById('ds');
  if(!/\\.docx$/i.test(file.name)){t.textContent='Hanya file .docx yang didukung';s.textContent='Simpan ulang .doc sebagai .docx di Word';f.value='';return}
  if(file.size>3*1024*1024){t.textContent='File terlalu besar';s.textContent='Maksimal 3 MB';f.value='';return}
  t.textContent=file.name;s.textContent='Membaca…';busy.classList.add('on');form.submit();
}
f.addEventListener('change',go);
['dragenter','dragover'].forEach(e=>d.addEventListener(e,ev=>{ev.preventDefault();d.classList.add('over')}));
['dragleave','drop'].forEach(e=>d.addEventListener(e,ev=>{ev.preventDefault();d.classList.remove('over')}));
d.addEventListener('drop',ev=>{if(ev.dataTransfer.files.length){f.files=ev.dataTransfer.files;go()}});
window.addEventListener('pageshow',()=>busy.classList.remove('on'));
</script></body></html>"""

FORM = """<!doctype html>
<html lang="id"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Isi: {{ name }}</title>
<style>
:root{--bg:#e9edf5;--bar:#ffffff;--ink:#1b2333;--mut:#667085;--line:#d9e0ee;--acc:#2f5fe3;--acc-l:#eef3ff;--ok:#16a34a;--okbg:#e3f6e8;--warnbg:#fff4c2;--warnbg2:#ffe98a}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);font-family:system-ui,-apple-system,"Segoe UI",Roboto,sans-serif}
.bar{position:sticky;top:0;z-index:10;background:var(--bar);border-bottom:1px solid var(--line);box-shadow:0 2px 12px rgba(16,24,40,.07)}
.bar-in{max-width:980px;margin:0 auto;padding:10px 16px;display:flex;gap:10px 14px;align-items:center;flex-wrap:wrap}
.back{display:grid;place-items:center;width:36px;height:36px;border-radius:9px;color:var(--mut);text-decoration:none;border:1px solid var(--line)}
.back:hover{background:var(--acc-l);color:var(--acc)}
.t{font-weight:600;max-width:26ch;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.prog{flex:1;min-width:170px}
.prog small{display:block;font-size:.8rem;color:var(--mut);margin-bottom:4px}
.meter{height:6px;border-radius:99px;background:var(--line);overflow:hidden}
.meter i{display:block;height:100%;width:0;background:var(--ok);transition:width .25s}
.btn{border:0;border-radius:9px;padding:9px 14px;font:inherit;font-weight:600;cursor:pointer;text-decoration:none;font-size:.9rem;display:inline-flex;align-items:center;gap:6px;white-space:nowrap}
.btn:focus-visible,.back:focus-visible,summary:focus-visible{outline:3px solid #9db4ff;outline-offset:2px}
.btn.main{background:var(--acc);color:#fff}.btn.main:hover{background:#2349b8}
.btn.main[disabled]{opacity:.7;cursor:wait}
.btn.sec{background:var(--acc-l);color:var(--acc)}.btn.sec:hover{background:#dde7ff}
.btn.ghost{background:transparent;color:var(--mut);border:1px solid var(--line)}.btn.ghost:hover{background:#f3f5fa}
details.opts{position:relative}
details.opts summary{list-style:none;cursor:pointer;border:1px solid var(--line);border-radius:9px;padding:8px 12px;font-size:.9rem;font-weight:600;color:var(--mut);user-select:none}
details.opts summary::-webkit-details-marker{display:none}
details.opts[open] summary{background:var(--acc-l);color:var(--acc)}
.pop{position:absolute;right:0;top:calc(100% + 6px);width:290px;background:#fff;border:1px solid var(--line);border-radius:12px;padding:12px;box-shadow:0 10px 30px rgba(16,24,40,.16);z-index:20}
.pop label{display:flex;gap:9px;align-items:flex-start;padding:7px 4px;cursor:pointer;font-size:.88rem;line-height:1.35}
.pop label span small{display:block;color:var(--mut);font-size:.78rem}
.pop input{accent-color:var(--acc);width:1.1em;height:1.1em;margin-top:2px}
.hint{max-width:816px;margin:14px auto 0;padding:0 14px;color:var(--mut);font-size:.88rem;line-height:1.5}
.hint mark{background:var(--warnbg);padding:0 4px;border-radius:3px}
.paper{background:#fff;max-width:816px;margin:12px auto 48px;padding:76px;box-shadow:0 4px 24px rgba(16,24,40,.14);border-radius:3px;
font-family:"Times New Roman",Times,serif;font-size:12pt;color:#000}
.paper p{white-space:pre-wrap;word-wrap:break-word}
table.doc{border-collapse:collapse;table-layout:fixed}
table.doc td{padding:1px 4px;overflow-wrap:anywhere}
table.bordered td{border:1px solid #000}
.fld{font:inherit;color:#0b3d91;border:0;border-bottom:1.5px dotted #8a8f99;background:var(--warnbg);padding:1px 5px;outline:none;
min-width:8ch;max-width:100%;border-radius:4px 4px 0 0;transition:background .12s}
.fld:hover{background:var(--warnbg2)}
.fld::placeholder{color:#a8923a;font-size:.78em;font-style:italic}
.fld:focus{background:#fff;border-bottom:2px solid var(--acc);box-shadow:0 0 0 3px rgba(47,95,227,.2)}
.fld.filled{background:var(--okbg);border-bottom-color:var(--ok)}
.fld.filled:focus{background:#fff}
.fld.full{width:100%;display:block}
.paper .fld{text-align:inherit}
.fld.date{min-width:0;width:11.5em;font-size:.9em}
.cb{width:1.15em;height:1.15em;vertical-align:-.18em;accent-color:var(--acc);cursor:pointer;margin:0 3px}
.toast{position:fixed;left:50%;bottom:22px;transform:translateX(-50%) translateY(20px);background:#1b2333;color:#fff;padding:10px 16px;border-radius:10px;
font-size:.9rem;opacity:0;pointer-events:none;transition:.25s;z-index:60}
.toast.on{opacity:1;transform:translateX(-50%)}
@media(max-width:700px){.paper{padding:22px 14px;font-size:11pt;margin-top:8px}.t{display:none}.btn .lbl{display:none}.pop{right:auto;left:0}}
table.dyn tr.dr>td:first-child{position:relative}
.rm{position:absolute;left:-30px;top:50%;transform:translateY(-50%);width:22px;height:22px;border-radius:50%;border:1px solid var(--line);background:#fff;color:#b42318;cursor:pointer;font:700 14px/1 system-ui,sans-serif;padding:0;opacity:.55}
.rm:hover,.rm:focus-visible{opacity:1;background:#fef3f2}
.dctl{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin:8px 0 14px}
.dctl .btn{font-family:system-ui,-apple-system,"Segoe UI",Roboto,sans-serif;font-size:.85rem;padding:7px 12px}
.dctl .btn[disabled]{opacity:.4;cursor:not-allowed}
.dcnt{color:var(--mut);font:.82rem system-ui,sans-serif}
@media(max-width:700px){.rm{left:-13px;width:18px;height:18px;font-size:12px}}
@media print{.bar,.hint,.toast,.dctl,.rm{display:none}.paper{box-shadow:none;margin:0}}
</style></head><body>
<form id="f" method="post" action="{{ url_for('download') }}" autocomplete="off">
<input type="hidden" name="_doc" value="{{ doc_b64 }}">
<input type="hidden" name="_name" value="{{ name }}">
<header class="bar"><div class="bar-in">
  <a class="back" href="{{ url_for('home') }}" title="Ganti file" aria-label="Ganti file"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M15 18l-6-6 6-6"/></svg></a>
  <div class="t" title="{{ name }}">{{ name }}</div>
  <div class="prog"><small id="prog"></small><div class="meter"><i id="meter"></i></div></div>
  <button type="button" class="btn sec" id="nextEmpty" title="Lompat ke isian yang belum terisi">↓ <span class="lbl">Isian kosong</span></button>
  <details class="opts"><summary>⚙ <span class="lbl">Pengaturan</span></summary>
    <div class="pop">
      <label><input type="checkbox" name="fit_dots" value="1" checked><span>Pas di titik-titik<small>Tulisan dipaskan pada garis titik; font mengecil otomatis jika terlalu panjang.</small></span></label>
      <label><input type="checkbox" name="drop_empty" value="1" checked><span>Hapus titik-titik yang kosong<small>Titik-titik yang tidak diisi dihilangkan di file Word. Matikan jika formulir akan ditulis tangan.</small></span></label>
      <label><input type="checkbox" name="fit_page" value="1"><span>Muat 1 halaman<small>Padatkan spasi/margin/font agar semua muat satu halaman. Biarkan mati untuk mengikuti halaman Word asli.</small></span></label>
    </div></details>
  <button type="button" class="btn ghost" id="reset">Kosongkan</button>
  <button type="submit" class="btn main" id="dl">⬇ <span class="lbl">Unduh Word</span></button>
</div></header>
<div class="hint">Klik bagian <mark>berwarna kuning</mark> untuk mengisi atau mengubah; <b>Enter</b> pindah ke isian berikutnya.
Ditemukan <b>{{ n_text }}</b> isian dan <b>{{ n_check }}</b> kotak centang.</div>
<div class="paper">{{ body|safe }}</div>
</form>
<div class="toast" id="toast"></div>
<script>
const form=document.getElementById('f');
let texts=[],checks=[];
function collect(){texts=[...form.querySelectorAll('input.fld')];checks=[...form.querySelectorAll('input.cb')]}
collect();
const prog=document.getElementById('prog'),meter=document.getElementById('meter');
const toast=document.getElementById('toast');
function say(m){toast.textContent=m;toast.classList.add('on');clearTimeout(say.t);say.t=setTimeout(()=>toast.classList.remove('on'),2200)}
function update(){
  let a=0,b=0;
  texts.forEach(e=>{const on=e.value.trim()!=='';e.classList.toggle('filled',on);if(on)a++});
  checks.forEach(e=>{if(e.checked)b++});
  prog.textContent=a+' dari '+texts.length+' isian terisi'+(checks.length?' · '+b+'/'+checks.length+' dicentang':'');
  meter.style.width=(texts.length?a/texts.length*100:100)+'%';
}
form.addEventListener('change',ev=>{
  const c=ev.target;if(!c.classList||!c.classList.contains('cb'))return;
  const g=c.dataset.group;
  if(g&&c.checked)checks.forEach(o=>{if(o!==c&&o.dataset.group===g)o.checked=false});
  update();
});
const links={};
texts.forEach(e=>{if(e.dataset.link){(links[e.dataset.link]=links[e.dataset.link]||[]).push(e)}});
form.addEventListener('input',ev=>{
  const e=ev.target;if(!e.classList||!e.classList.contains('fld'))return;
  const g=links[e.dataset.link];
  if(g){
    if(e===g[0]){g.slice(1).forEach(o=>{if(!o.dataset.touched)o.value=e.value})}
    else{if(e.value==='')delete e.dataset.touched;else e.dataset.touched='1'}
  }
  update();
});
Object.values(links).forEach(g=>g.slice(1).forEach(o=>{
  if(o.value==='')o.value=g[0].value;else if(o.value!==g[0].value)o.dataset.touched='1';
}));
form.addEventListener('keydown',ev=>{
  const e=ev.target;if(ev.key!=='Enter'||!e.classList||!e.classList.contains('fld'))return;
  ev.preventDefault();const i=texts.indexOf(e);(texts[i+1]||e).focus();
});
document.getElementById('nextEmpty').addEventListener('click',()=>{
  const cur=texts.indexOf(document.activeElement);
  const order=[...texts.slice(cur+1),...texts.slice(0,Math.max(cur,0)+1)];
  const t=order.find(e=>e.value.trim()==='');
  if(!t){say('Semua isian sudah terisi ✓');return}
  t.scrollIntoView({block:'center',behavior:'smooth'});t.focus({preventScroll:true});
});
document.getElementById('reset').addEventListener('click',()=>{
  if(!confirm('Kosongkan semua isian dan centang?'))return;
  texts.forEach(e=>{e.value=e.type==='date'?(e.dataset.def||''):'';delete e.dataset.touched});
  checks.forEach(e=>e.checked=false);update();say('Semua isian dikosongkan');
});
// tutup menu pengaturan saat klik di luar
document.addEventListener('click',ev=>{const d=document.querySelector('details.opts');if(d&&d.open&&!d.contains(ev.target))d.open=false});
// umpan balik saat mengunduh
// ---- tabel dinamis: tambah / kurangi baris ----
const MAXR=30;
function setNum(td,n){
  const w=document.createTreeWalker(td,NodeFilter.SHOW_TEXT);let t,done=false;
  while(t=w.nextNode()){
    if(t.parentNode.closest('button'))continue;
    if(!done&&t.nodeValue.trim()){t.nodeValue=String(n);done=true}
    else if(done)t.nodeValue='';
  }
}
function dynRows(tb){return [...tb.querySelectorAll('tr.dr')]}
function addRm(tr){
  const td=tr.cells[0];if(!td||td.querySelector('.rm'))return;
  const b=document.createElement('button');b.type='button';b.className='rm';b.textContent='×';
  b.title='Hapus baris ini';b.setAttribute('aria-label','Hapus baris ini');td.appendChild(b);
}
function refreshDyn(tb){
  const rows=dynRows(tb),ctl=document.querySelector('.dctl[data-for="'+tb.dataset.dyn+'"]');
  rows.forEach((tr,i)=>{if(tb.dataset.num==='1')setNum(tr.cells[0],i+1);addRm(tr);
    const rm=tr.querySelector('.rm');if(rm)rm.disabled=rows.length<=1});
  if(ctl){
    ctl.querySelector('.dcnt').textContent=rows.length+' baris';
    ctl.querySelector('.ddel').disabled=rows.length<=1;
    ctl.querySelector('.dadd').disabled=rows.length>=MAXR;
  }
  collect();update();
}
function addRow(tb){
  const rows=dynRows(tb);if(!rows.length||rows.length>=MAXR)return;
  const nr=rows[rows.length-1].cloneNode(true);
  nr.querySelectorAll('input').forEach(i=>{
    if(i.type==='checkbox')i.checked=false;
    else if(i.type==='date')i.value=i.dataset.def||'';
    else i.value='';
    delete i.dataset.touched;
  });
  rows[rows.length-1].after(nr);refreshDyn(tb);
  const f=nr.querySelector('input.fld');if(f){f.scrollIntoView({block:'center',behavior:'smooth'});f.focus({preventScroll:true})}
}
function delRow(tb,tr){
  const rows=dynRows(tb);if(rows.length<=1){say('Minimal 1 baris');return}
  tr.remove();refreshDyn(tb);say('Baris dihapus');
}
document.querySelectorAll('table.dyn').forEach(tb=>{
  const first=dynRows(tb)[0];
  tb.dataset.num=first&&/^\\s*\\d+\\s*$/.test(first.cells[0].textContent)?'1':'0';
  refreshDyn(tb);
});
form.addEventListener('click',ev=>{
  const t=ev.target;
  if(t.classList.contains('rm')){const tr=t.closest('tr'),tb=t.closest('table.dyn');if(tr&&tb)delRow(tb,tr);return}
  const ctl=t.closest('.dctl');if(!ctl)return;
  const tb=document.querySelector('table.dyn[data-dyn="'+ctl.dataset.for+'"]');if(!tb)return;
  if(t.classList.contains('dadd'))addRow(tb);
  else if(t.classList.contains('ddel')){const r=dynRows(tb);if(r.length>1)delRow(tb,r[r.length-1])}
});
form.addEventListener('submit',()=>{
  document.querySelectorAll('table.dyn').forEach(tb=>{   // beri nama isian menurut posisi (tabel, baris, isian)
    const T=tb.dataset.dyn,rows=dynRows(tb);
    rows.forEach((tr,i)=>[...tr.querySelectorAll('input')].forEach((el,j)=>{el.name='d'+T+'_'+i+'_'+j}));
    let h=form.querySelector('input[name="_dt'+T+'_n"]');
    if(!h){h=document.createElement('input');h.type='hidden';h.name='_dt'+T+'_n';form.appendChild(h)}
    h.value=rows.length;
  });
  const b=document.getElementById('dl'),old=b.innerHTML;
  b.disabled=true;b.innerHTML='Menyiapkan…';
  setTimeout(()=>{b.disabled=false;b.innerHTML=old;say('File Word diunduh ✓')},2500);
});
update();
</script></body></html>"""


# --------------------------------------------------------------------------
# Aplikasi Flask
# --------------------------------------------------------------------------
app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = int(MAX_UPLOAD * 1.5)
app.config["MAX_FORM_MEMORY_SIZE"] = 8 * 1024 * 1024
app.config["MAX_FORM_PARTS"] = 2000


def home_page(error=None, status=200):
    items = [(k, v[0], v[1]) for k, v in TEMPLATES.items() if (FORMS_DIR / v[2]).exists()]
    return render_template_string(HOME, error=error, templates=items), status


@app.get("/")
def home():
    return home_page()


def open_doc(data, name):
    """Validasi lalu tampilkan halaman isian. Tanpa penyimpanan di server (cocok untuk Vercel):
    file .docx dibawa bolak-balik oleh browser sebagai field tersembunyi."""
    try:
        doc = FormDoc(data)
    except (zipfile.BadZipFile, ValueError, etree.XMLSyntaxError) as e:
        return home_page(str(e) if isinstance(e, ValueError)
                         else "File tidak bisa dibaca. Pastikan formatnya .docx (bukan .doc).", 400)
    if not doc.slots:
        return home_page("Tidak ditemukan isian (titik-titik, Label :, sel tabel kosong) maupun kotak centang di dokumen ini.", 400)
    return render_template_string(
        FORM, name=name, doc_b64=base64.b64encode(data).decode("ascii"), body=doc.render_html(),
        today=date.today().isoformat(),
        n_text=sum(s.kind != "check" for s in doc.slots),
        n_check=sum(s.kind == "check" for s in doc.slots))


@app.errorhandler(413)
def too_big(_e):
    return home_page("File terlalu besar (maksimal 3 MB).", 413)


@app.post("/upload")
def upload():
    f = request.files.get("file")
    if not f or not f.filename:
        return home_page("Pilih file .docx terlebih dahulu.", 400)
    if not f.filename.lower().endswith(".docx"):
        return home_page("Hanya file .docx yang didukung (simpan ulang .doc sebagai .docx).", 400)
    return open_doc(f.read(), os.path.basename(f.filename))


def template_file(key):
    if key not in TEMPLATES:
        abort(404)
    f = FORMS_DIR / TEMPLATES[key][2]
    if not f.exists():
        abort(404)
    return f


@app.get("/formulir/<key>")
def template_fill(key):
    f = template_file(key)
    return open_doc(f.read_bytes(), f.name)


@app.get("/formulir/<key>/unduh")
def template_download(key):
    f = template_file(key)
    return send_file(f, as_attachment=True, download_name=f.name,
                     mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


@app.post("/download")
def download():
    try:
        data = base64.b64decode(request.form.get("_doc", ""), validate=True)
        rows = {}
        for k, v in request.form.items():       # jumlah baris tabel dinamis dari browser
            m = re.fullmatch(r"_dt(\d+)_n", k)
            if m and v.isdigit():
                rows[int(m.group(1))] = int(v)
        doc = FormDoc(data, rows)
    except Exception:
        return home_page("Sesi formulir tidak valid. Upload ulang file .docx.", 400)
    name = os.path.basename(request.form.get("_name") or "formulir.docx")
    doc.fill(request.form, fit_dots=bool(request.form.get("fit_dots")),
             drop_empty=bool(request.form.get("drop_empty")))
    if request.form.get("fit_page"):
        doc.fit_one_page()
    stem = re.sub(r"\.docx$", "", name, flags=re.I)
    return send_file(io.BytesIO(doc.to_docx()), as_attachment=True,
                     download_name=f"{stem}_TERISI.docx",
                     mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


if __name__ == "__main__":
    host = os.environ.get("HOST", "127.0.0.1")
    port = int(os.environ.get("PORT", "5000"))
    print(f"\n  Buka di browser:  http://{'127.0.0.1' if host == '0.0.0.0' else host}:{port}\n")
    app.run(host=host, port=port, debug=False)
