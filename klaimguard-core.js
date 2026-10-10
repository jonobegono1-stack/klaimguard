/* KlaimGuard core — parsing & aturan cek klaim (browser + node). Tanpa DOM. */
(function (global) {
  "use strict";

  const MONTHS = {
    "JANUARI": 1, "JANUARY": 1, "JAN": 1,
    "FEBRUARI": 2, "FEBRUARY": 2, "FEB": 2,
    "MARET": 3, "MARCH": 3, "MAR": 3,
    "APRIL": 4, "APR": 4,
    "MEI": 5, "MAY": 5,
    "JUNI": 6, "JUNE": 6, "JUN": 6,
    "JULI": 7, "JULY": 7, "JUL": 7,
    "AGUSTUS": 8, "AUGUST": 8, "AGU": 8, "AUG": 8,
    "SEPTEMBER": 9, "SEP": 9,
    "OKTOBER": 10, "OCTOBER": 10, "OKT": 10, "OCT": 10,
    "NOVEMBER": 11, "NOV": 11,
    "DESEMBER": 12, "DECEMBER": 12, "DES": 12, "DEC": 12,
  };

  function pad(n) { return (n < 10 ? "0" : "") + n; }

  // Terima: Date, number (serial Excel), atau string tanggal Indonesia/ISO.
  function parseTanggal(v) {
    if (v == null) return null;
    if (v instanceof Date && !isNaN(v)) return v;
    if (typeof v === "number") {
      const d = new Date(Math.round((v - 25569) * 86400 * 1000));
      return isNaN(d) ? null : d;
    }
    const s = String(v).trim();
    if (!s || s === "-" || s === "0") return null;
    let m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/); // dd/mm/yyyy
    if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
    m = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);    // yyyy-mm-dd
    if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
    m = s.match(/^(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})$/);        // 1 Mei 2026
    if (m) {
      const mo = MONTHS[m[2].toUpperCase()];
      if (mo) return new Date(+m[3], mo - 1, +m[1]);
    }
    return null;
  }

  function isTanggalValid(d) { return d instanceof Date && !isNaN(d); }
  function tglStr(d) { return d ? pad(d.getDate()) + "/" + pad(d.getMonth() + 1) + "/" + d.getFullYear() : "-"; }

  function parseFloatNum(v) {
    if (v == null) return 0;
    if (typeof v === "number") return v;
    const s = String(v).replace(/[^0-9.\-]/g, "");
    const n = parseFloat(s);
    return isNaN(n) ? 0 : n;
  }

  // ============ KAMUS ICD-10 ============
  let ICD = null, ICD_PREFIX = null;
  function initICD() {
    if (ICD) return;
    ICD = (typeof global.ICD10 === "object" && global.ICD10) ? global.ICD10 : {};
    ICD_PREFIX = new Set();
    for (const k in ICD) { if (k.length >= 3) ICD_PREFIX.add(k.slice(0, 3)); }
  }
  function icdInfo(code) {
    code = String(code == null ? "" : code).trim().toUpperCase();
    if (!code) return null;
    if (!/^[A-Z]\d{2}(\.\d{1,2})?$/.test(code)) return { code, status: "format", desc: null };
    if (ICD[code] !== undefined) return { code, status: "ok", desc: ICD[code] };
    const p = code.slice(0, 3);
    if (ICD[p] !== undefined) return { code, status: "sub", desc: ICD[p] };
    if (ICD_PREFIX.has(p)) return { code, status: "sub", desc: null };
    return { code, status: "unknown", desc: null };
  }
  function saranSpesifik(code) {
    code = String(code == null ? "" : code).trim().toUpperCase();
    if (code.length < 4) return [];
    const p = code.slice(0, 3);
    const out = [];
    for (const k in ICD) {
      if (k.startsWith(p) && k !== code && !/\.9$/.test(k)) out.push([k, ICD[k]]);
    }
    out.sort((a, b) => a[0].localeCompare(b[0]));
    return out.slice(0, 6);
  }
  function cariICD(q) {
    initICD();
    q = String(q == null ? "" : q).trim();
    if (!q) return [];
    const ql = q.toLowerCase();
    const out = [];
    if (/^[A-Za-z]\d/.test(q)) {
      const info = icdInfo(q);
      if (info && info.desc) out.push([q.toUpperCase(), info.desc + (info.status === "sub" ? " (kategori " + q.slice(0, 3) + ")" : "")]);
    }
    for (const k in ICD) {
      const d = ICD[k];
      if (d && d.toLowerCase().includes(ql)) out.push([k, d]);
      if (out.length >= 40) break;
    }
    return out.slice(0, 30);
  }

  // ============ DETEKSI FORMAT ============
  function deteksiFormat(textOrAoA, namaFile) {
    const nama = (namaFile || "").toLowerCase();
    if (typeof textOrAoA === "string") return "mix"; // teks = ekspor MIX (^I)
    if (nama.endsWith(".txt")) return "mix";
    return "rekap"; // xlsx = rekap bulanan
  }

  // ============ REKAP (XLSX, 11 kolom) ============
  function norm(s) { return String(s == null ? "" : s).toLowerCase().replace(/\s+/g, " ").trim(); }

  function cariHeaderRekap(aoa) {
    for (let i = 0; i < aoa.length; i++) {
      const row = aoa[i] || [];
      const cells = row.map(c => norm(c));
      if (cells.includes("no.") && cells.some(c => c.includes("inacbg"))) return i;
      if (cells.includes("no") && cells.some(c => c.includes("inacbg"))) return i;
    }
    return -1;
  }

  function petaKolom(header) {
    const map = {};
    header.forEach((h, i) => {
      const n = norm(h);
      if (!n) return;
      if (n === "no." || n === "no") { if (map.no === undefined) map.no = i; return; }
      if (n.includes("sep") || n.includes("no klaim") || n.includes("no. klaim")) { map.sep = i; return; }
      if (n.includes("inacbg")) { map.inacbg = i; return; }
      if (n.includes("tgl masuk") || (n.includes("tgl") && n.includes("masuk"))) { map.masuk = i; return; }
      if (n.includes("tgl pulang") || (n.includes("tgl") && n.includes("pulang"))) { map.pulang = i; return; }
      if (n.includes("no rm") || n === "no. rm" || n.includes("rekam")) { map.rm = i; return; }
      if (n.includes("nama")) { map.nama = i; return; }
      if (n.includes("top")) { map.topup = i; return; }
      if (n.includes("total tarif")) { map.total = i; return; }
      if (n.includes("tarif rs")) { map.tarif_rs = i; return; }
      if (n.includes("jenis")) { map.jenis = i; return; }
    });
    return map;
  }

  function parseRekap(aoa) {
    const hi = cariHeaderRekap(aoa);
    if (hi < 0) return { format: "rekap", error: "Baris header ('No.' & 'INACBG') tidak ditemukan." };
    const header = aoa[hi].map(c => c == null ? "" : String(c));
    const m = petaKolom(header);
    const judul = aoa.slice(0, hi).map(r => (r && r[0]) ? String(r[0]) : "").filter(Boolean);
    const claims = [];
    for (let i = hi + 1; i < aoa.length; i++) {
      const r = aoa[i] || [];
      const no = m.no !== undefined ? String(r[m.no] == null ? "" : r[m.no]).trim() : "";
      if (!/^\d+$/.test(no)) continue; // lewati baris total/judul
      const g = (k) => (m[k] !== undefined && r[m[k]] != null) ? r[m[k]] : "";
      claims.push({
        baris: i + 1,
        no: parseInt(no, 10),
        sep: String(g("sep")).trim(),
        inacbg: String(g("inacbg")).trim(),
        rm: String(g("rm")).trim(),
        nama: String(g("nama")).trim(),
        masuk: parseTanggal(g("masuk")),
        pulang: parseTanggal(g("pulang")),
        masukRaw: String(g("masuk")).trim(),
        pulangRaw: String(g("pulang")).trim(),
        topup: parseFloatNum(g("topup")),
        topupRaw: String(g("topup")).trim(),
        total: parseFloatNum(g("total")),
        tarif_rs: parseFloatNum(g("tarif_rs")),
        jenis: String(g("jenis")).trim().toUpperCase(),
      });
    }
    return { format: "rekap", judul, header, claims, count: claims.length };
  }

  // ============ MIX (TXT, delimiter ^I) ============
  function parseMix(text) {
    const lines = String(text).split(/\r?\n/).filter(l => l.trim().length > 0);
    if (lines.length === 0) return { format: "mix", error: "File kosong." };
    const header = lines[0].split("\t").map(h => String(h).trim());
    if (!header.includes("KODE_RS") && !header.includes("DIAGLIST")) {
      return { format: "mix", error: "Header MIX (KODE_RS/DIAGLIST) tidak ditemukan — pastikan ekspor INA-CBG MIX (.TXT)." };
    }
    const idx = {};
    header.forEach((h, i) => { idx[h] = i; });
    const get = (r, k) => (idx[k] !== undefined && r[idx[k]] != null) ? String(r[idx[k]]).trim() : "";
    const claims = [];
    for (let i = 1; i < lines.length; i++) {
      const parts = lines[i].split("\t");
      if (parts.length < 5) continue;
      const inacbg = get(parts, "INACBG");
      const sep = get(parts, "SEP");
      if (!inacbg && !sep) continue;
      const kelas = inacbg.split("-").pop() || "";
      const diag = (get(parts, "DIAGLIST") || "").split(";").map(s => s.trim()).filter(s => s && s !== "-");
      const proc = (get(parts, "PROCLIST") || "").split(";").map(s => s.trim()).filter(s => s && s !== "-");
      claims.push({
        baris: i + 1,
        sep: sep,
        inacbg: inacbg,
        kelas: kelas,
        sev: (inacbg.split("-")[1] || ""),
        jenis: kelas === "0" ? "RJ" : "RI",
        nama: get(parts, "NAMA_PASIEN"),
        rm: get(parts, "MRN"),
        sex: get(parts, "SEX"),
        umur: parseFloatNum(get(parts, "UMUR_TAHUN")),
        los: parseFloatNum(get(parts, "LOS")),
        diag: diag,
        proc: proc,
        diagUtama: diag[0] || "",
        jmlDiag: diag.length,
        cw: parseFloatNum(get(parts, "IDRG_COST_WEIGHT")),
        cwRaw: get(parts, "IDRG_COST_WEIGHT"),
        tarifIdrg: parseFloatNum(get(parts, "IDRG_TOTAL_TARIF")),
        tarifInacbg: parseFloatNum(get(parts, "TARIF_INACBG")),
        deskripsi: get(parts, "DESKRIPSI_INACBG"),
        drgDesc: get(parts, "IDRG_DRG_DESCRIPTION"),
        masukRaw: get(parts, "ADMISSION_DATE"),
        pulangRaw: get(parts, "DISCHARGE_DATE"),
      });
    }
    return { format: "mix", header, claims, count: claims.length };
  }

  // ============ CEK ============
  const RE = /^[A-Z]-\d+-\d+-(0|I|II|III)$/i;

  function cekRekap(claims) {
    const sepCount = {};
    claims.forEach(c => { if (c.sep) sepCount[c.sep] = (sepCount[c.sep] || 0) + 1; });

    const results = [];
    const checks = { dup: 0, tgl: 0, tglkosong: 0, inacbg: 0, kelas: 0, nilai: 0, topup: 0 };
    for (const c of claims) {
      const masalah = [];
      const add = (level, label) => masalah.push({ level, label });

      if (!c.sep) { add("merah", "Nomor klaim/SEP kosong"); checks.dup++; }
      else if (sepCount[c.sep] > 1) { add("merah", "No SEP duplikat (muncul " + sepCount[c.sep] + "×)"); checks.dup++; }

      if (!c.masukRaw && !c.pulangRaw) { checks.tglkosong++; add("kuning", "Tanggal masuk & pulang kosong"); }
      else if (!c.masuk || !c.pulang) { add("kuning", "Tanggal tidak bisa dibaca"); checks.tgl++; }
      else if (c.pulang < c.masuk) { add("merah", "Tgl pulang < tgl masuk (" + tglStr(c.pulang) + " < " + tglStr(c.masuk) + ")"); checks.tgl++; }

      if (!c.inacbg) { add("kuning", "Kode INACBG kosong"); checks.inacbg++; }
      else if (!RE.test(c.inacbg)) { add("kuning", "Format kode INACBG tidak standar: " + c.inacbg); checks.inacbg++; }
      else {
        const suf = c.inacbg.split("-").pop();
        const cocok = (c.jenis === "RI" && /^I+$/.test(suf)) || (c.jenis === "RJ" && suf === "0");
        if (c.jenis === "RI" && suf === "0") { add("kuning", "Jenis RI tapi kode berkelas rawat jalan (0)"); checks.kelas++; }
        else if (c.jenis === "RJ" && /^I+$/.test(suf)) { add("kuning", "Jenis RJ tapi kode berkelas rawat inap (" + suf + ")"); checks.kelas++; }
        else if (!cocok && c.jenis) { add("kuning", "Kelas kode vs jenis tidak cocok"); checks.kelas++; }
      }

      if (c.total <= 0) { add("merah", "Nilai klaim ≤ 0"); checks.nilai++; }
      if (c.topupRaw && parseFloatNum(c.topup) !== 0) { add("kuning", "Ada top-up: " + c.topupRaw); checks.topup++; }

      const level = masalah.some(x => x.level === "merah") ? "merah"
        : masalah.some(x => x.level === "kuning") ? "kuning" : "hijau";
      results.push({
        baris: c.baris, no: c.no, sep: c.sep, jenis: c.jenis, nama: c.nama, rm: c.rm,
        inacbg: c.inacbg, masuk: tglStr(c.masuk), pulang: tglStr(c.pulang),
        topup: c.topupRaw, total: c.total, tarif_rs: c.tarif_rs, masalah, level,
      });
    }
    return { results, summary: ringkas(results, claims, checks) };
  }

  function cekMix(claims) {
    initICD();
    const results = [];
    const checks = { gagal: 0, satuDx: 0, dx9: 0, sex: 0, umur: 0, invalid: 0, under: 0 };
    for (const c of claims) {
      const masalah = [];
      const add = (level, label) => masalah.push({ level, label });

      if (c.cwRaw === "" || c.cwRaw === "-" || c.cw === 0 || c.tarifIdrg === 0) {
        add("merah", "Gagal grouping iDRG (cost weight 0 / tarif 0)");
        checks.gagal++;
      }

      const ranap = /^I+$/.test(c.kelas);
      if (ranap && c.jmlDiag === 1) { add("kuning", "Hanya 1 diagnosis — cek komorbiditas"); checks.satuDx++; }

      if (c.diagUtama && /\.9$/.test(c.diagUtama)) {
        const s = saranSpesifik(c.diagUtama);
        const sS = s.length ? " → coba: " + s.map(x => x[0]).slice(0, 4).join(", ") : "";
        add("kuning", "Diagnosa utama terlalu umum (" + c.diagUtama + ")" + sS);
        checks.dx9++;
      }

      // validitas kode vs kamus ICD-10 (induk = fallback subdivisi valid)
      for (const d of c.diag) {
        const info = icdInfo(d);
        if (info && info.status === "format") {
          add("kuning", "Format kode " + d + " tidak standar — periksa");
          checks.invalid++;
        } else if (info && info.status === "unknown") {
          add("kuning", "Kode " + d + " tidak ada di kamus lokal — verifikasi");
          checks.invalid++;
        }
      }

      if (c.sex === "1" && c.diag.some(d => /^O/i.test(d))) {
        add("merah", "Kode obstetri (" + c.diag.filter(d => /^O/i.test(d))[0] + ") di pasien laki-laki");
        checks.sex++;
      }

      if (c.umur >= 1 && c.diag.some(d => /^P/i.test(d))) {
        add("kuning", "Kode neonatal (P) di pasien usia " + c.umur + " th");
        checks.umur++;
      }

      if (ranap && c.sev === "1" && c.jmlDiag === 1 && c.los >= 3) {
        add("kuning", "Potensi under-coding: severity I + 1 diagnosis + LOS " + c.los + " hari");
        checks.under++;
      }

      const level = masalah.some(x => x.level === "merah") ? "merah"
        : masalah.some(x => x.level === "kuning") ? "kuning" : "hijau";
      results.push({
        baris: c.baris, sep: c.sep, jenis: c.jenis, nama: c.nama, rm: c.rm,
        inacbg: c.inacbg, deskripsi: c.deskripsi, diag: c.diag.join("; "), proc: c.proc.join("; "),
        cw: c.cwRaw, los: c.los, umur: c.umur, sex: c.sex, masalah, level,
      });
    }
    return { results, summary: ringkas(results, claims, checks) };
  }

  function ringkas(results, claims, checks) {
    const rajal = claims.filter(c => c.jenis === "RJ").length;
    const ranap = claims.filter(c => c.jenis === "RI").length;
    const merah = results.filter(r => r.level === "merah").length;
    const kuning = results.filter(r => r.level === "kuning").length;
    const hijau = results.filter(r => r.level === "hijau").length;
    return { total: claims.length, rajal, ranap, merah, kuning, hijau, checks };
  }

  function proses(textOrAoA, namaFile) {
    const fmt = deteksiFormat(textOrAoA, namaFile);
    if (fmt === "mix") {
      const p = parseMix(textOrAoA);
      if (p.error) return p;
      const c = cekMix(p.claims);
      return { format: "mix", judul: [], header: p.header, count: p.count, ...c };
    }
    const p = parseRekap(textOrAoA);
    if (p.error) return p;
    const c = cekRekap(p.claims);
    return { format: "rekap", judul: p.judul, header: p.header, count: p.count, ...c };
  }

  const KG = { VERSION: "0.2.0", parseTanggal, parseRekap, parseMix, cekRekap, cekMix, proses, tglStr, cariICD, icdInfo, saranSpesifik };

  if (typeof module !== "undefined" && module.exports) module.exports = KG;
  global.KG = KG;
})(typeof window !== "undefined" ? window : globalThis);
