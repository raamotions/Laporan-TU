/* =========================================================
   Reren Recon
   Menggabungkan rekening koran (CSV bank) dengan laporan top up
   (xlsx) menjadi satu tabel hasil, sama seperti sheet "MICO".
   ========================================================= */

/* ---------------------------------------------------------
   BAGIAN 1 – LOGIKA INTI (tanpa DOM, bisa diuji di Node)
   --------------------------------------------------------- */
const Core = (() => {
  const HEADERS = [
    'Agency Mico ID', 'Agency name', 'Target Mico ID', 'Target name', 'Coins',
    'Currency', 'Currency amount', 'Value to USD', 'Value (1USD)', 'Target is agent',
    'Note1', 'Note2', 'Note3', 'Country', 'Recharge time(GMT+7)', 'Transaction Status'
  ];

  const pad = (n, l = 2) => String(n).padStart(l, '0');

  /* ---- waktu: disimpan sebagai "jam dinding" (tanpa zona waktu) ---- */
  function mk(y, mo, d, h = 0, mi = 0, s = 0, ms = 0) {
    return { y, mo, d, h, mi, s, ms, ts: Date.UTC(y, mo - 1, d, h, mi, s, ms) };
  }
  function fromUTCms(ms) {
    const dt = new Date(ms);
    return mk(dt.getUTCFullYear(), dt.getUTCMonth() + 1, dt.getUTCDate(),
      dt.getUTCHours(), dt.getUTCMinutes(), dt.getUTCSeconds(), dt.getUTCMilliseconds());
  }
  function parseTime(v) {
    if (v == null || v === '') return null;
    if (typeof v === 'number' && isFinite(v)) {
      // nomor seri Excel -> ms sejak epoch
      return fromUTCms(Math.round((v - 25569) * 86400000));
    }
    if (v instanceof Date) return fromUTCms(v.getTime());
    const s = String(v).trim();
    let m = s.match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2}):(\d{2})(?:[.,](\d+))?/);
    if (m) {
      const ms = m[7] ? Math.round(parseFloat('0.' + m[7]) * 1000) : 0;
      return mk(+m[1], +m[2], +m[3], +m[4], +m[5], +m[6], ms);
    }
    m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:[ ,]+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
    if (m) return mk(+m[3], +m[2], +m[1], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0));
    return null;
  }
  const fmtDateTime = t => t ? `${t.y}-${pad(t.mo)}-${pad(t.d)} ${pad(t.h)}:${pad(t.mi)}:${pad(t.s)}` : '';
  const fmtDate = t => t ? `${pad(t.d)}/${pad(t.mo)}/${t.y}` : '';
  const fmtDateFull = t => t ? `${fmtDate(t)} ${pad(t.h)}:${pad(t.mi)}:${pad(t.s)}` : '';

  /* ---- rekening koran (CSV bank) ---- */
  // Satu transaksi bisa terpecah ke beberapa baris fisik (Keterangan memuat enter),
  // jadi baris yang tidak diawali tanggal disambung ke transaksi sebelumnya sampai polanya lengkap.
  const DATE_START = /^\s*'?\d{2}\/\d{2}\/\d{4}\s*,/;
  const LINE_RE = /^\s*'?(\d{2})\/(\d{2})\/(\d{4})\s*,(.*?),\s*["']?\d*["']?\s*,\s*"?([\d.,]+)"?\s*,\s*(CR|DB)\s*(?:,\s*"?([\d.,]+)"?)?\s*$/i;

  // "1,000,000.00" / "1.000.000,00" / "1000000.00" -> number
  function parseAmt(raw) {
    let s = String(raw).trim();
    const hasC = s.includes(','), hasD = s.includes('.');
    if (hasC && hasD) {
      s = s.lastIndexOf(',') > s.lastIndexOf('.') ? s.replace(/\./g, '').replace(',', '.') : s.replace(/,/g, '');
    } else if (hasC) {
      s = ((s.match(/,/g) || []).length === 1 && /,\d{1,2}$/.test(s)) ? s.replace(',', '.') : s.replace(/,/g, '');
    } else if (hasD && (s.match(/\./g) || []).length > 1) {
      s = s.replace(/\./g, '');
    }
    return parseFloat(s);
  }

  function parseRekening(text) {
    text = String(text).replace(/^\uFEFF/, '');
    const lines = text.split(/\r?\n/);
    const credits = [];
    const skippedLines = [];
    let debit = 0, skipped = 0, buf = null;

    const tryCommit = () => {
      const m = buf.match(LINE_RE);
      if (!m) return false;
      const amt = parseAmt(m[5]);
      if (isNaN(amt)) return false;
      if (m[6].toUpperCase() === 'DB') debit++;
      else credits.push({ idx: credits.length, date: mk(+m[3], +m[2], +m[1]), ket: m[4].trim(), amt });
      buf = null;
      return true;
    };
    const drop = () => { skipped++; if (skippedLines.length < 5) skippedLines.push(buf.slice(0, 160)); buf = null; };

    for (const line of lines) {
      if (DATE_START.test(line)) {
        if (buf !== null) drop();
        buf = line;
        tryCommit();
      } else if (buf !== null && line.trim()) {
        buf += ' ' + line.trim();
        tryCommit();
      }
    }
    if (buf !== null) drop();

    const nama = (text.match(/^Nama,=,(.*)$/m) || [])[1] || '';
    const rek = (text.match(/^No\. Rekening,=,'?(.*)$/m) || [])[1] || '';
    return { credits, debit, skipped, skippedLines, nama: nama.trim(), rek: rek.trim() };
  }

  // Gabungkan beberapa file rekening koran. Transaksi identik yang muncul di dua file
  // (periode tumpang tindih) hanya dihitung sekali.
  function mergeCredits(list) {
    const seen = new Map();
    const out = [];
    list.forEach(p => {
      const local = new Map();
      p.credits.forEach(c => {
        const k = `${c.date.ts}|${c.amt}|${c.ket}`;
        const n = (local.get(k) || 0) + 1;
        local.set(k, n);
        if (n <= (seen.get(k) || 0)) return;
        out.push({ idx: out.length, date: c.date, ket: c.ket, amt: c.amt });
      });
      local.forEach((n, k) => seen.set(k, Math.max(seen.get(k) || 0, n)));
    });
    return out;
  }

  /* ---- laporan top up ---- */
  function parseTopupAOA(aoa) {
    let h = -1;
    for (let i = 0; i < Math.min(aoa.length, 30); i++) {
      if ((aoa[i] || []).some(c => String(c).trim().toLowerCase() === 'agency mico id')) { h = i; break; }
    }
    if (h < 0) return null;
    const map = {};
    aoa[h].forEach((c, i) => { map[String(c).trim().toLowerCase()] = i; });
    const rows = [];
    for (let i = h + 1; i < aoa.length; i++) {
      const r = aoa[i] || [];
      const get = name => { const j = map[name.toLowerCase()]; return j == null ? '' : r[j]; };
      if (get('Agency Mico ID') === '' && get('Target Mico ID') === '') continue;
      const src = {};
      HEADERS.forEach(name => { src[name] = get(name); });
      const t = parseTime(src['Recharge time(GMT+7)']);
      rows.push({
        src,
        t,
        note3: src['Note3'] == null ? '' : String(src['Note3']),
        notes: ''
      });
    }
    // urut naik berdasarkan waktu top up (seperti sheet MICO)
    rows.sort((a, b) => (a.t ? a.t.ts : 0) - (b.t ? b.t.ts : 0));
    return rows;
  }

  function mergeTopups(list) {
    const seen = new Map();
    const out = [];
    list.forEach(rows => {
      const local = new Map();
      rows.forEach(r => {
        const k = JSON.stringify(HEADERS.map(h => r.src[h]));
        const n = (local.get(k) || 0) + 1;
        local.set(k, n);
        if (n <= (seen.get(k) || 0)) return;
        out.push(r);
      });
      local.forEach((n, k) => seen.set(k, Math.max(seen.get(k) || 0, n)));
    });
    out.sort((a, b) => (a.t ? a.t.ts : 0) - (b.t ? b.t.ts : 0));
    return out;
  }

  /* ---- pencocokan ---- */
  const squash = s => String(s).replace(/\s+/g, '').toLowerCase();
  const norm = s => String(s == null ? '' : s).normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
  const round2 = n => Math.round(n * 100) / 100;
  const nfRp = n => Number(n).toLocaleString('id-ID');
  const dayNum = t => Date.UTC(t.y, t.mo - 1, t.d) / 86400000;

  // kata-kata standar bank / gelar yang tidak membantu mengenali pengirim
  const BANKSTOP = new Set(['trsf', 'trf', 'transfer', 'tr', 'e', 'banking', 'ebanking', 'cr', 'db', 'ftscy', 'ke', 'dari',
    'bifast', 'bi', 'fast', 'qris', 'kliring', 'sknbi', 'rtgs', 'switching', 'setor', 'tunai', 'tarikan', 'mutasi',
    'pembayaran', 'bayar', 'pemindahbukuan', 'pbk', 'mbanking', 'myb', 'internet', 'atm']);
  const TITLES = new Set(['pt', 'cv', 'ud', 'bpk', 'bapak', 'ibu', 'ibuk', 'tn', 'tuan', 'ny', 'nyonya', 'sdr', 'sdri',
    'mr', 'mrs', 'ms', 'dr', 'drs', 'ir', 'hj', 'haji', 'toko', 'kios', 'konter', 'agen', 'agent', 'reseller', 'an']);

  function buildIndex(credits) {
    const all = credits.map(c => {
      const low = c.ket.toLowerCase();
      const n = norm(c.ket);
      const tok = n.split(' ').filter(t => /^[a-z]{2,}$/.test(t) && !BANKSTOP.has(t));
      return {
        idx: c.idx, amt: c.amt, date: c.date, ket: c.ket,
        low, sq: squash(c.ket),
        words: new Set(low.split(/[^a-z0-9]+/).filter(Boolean)),
        nsq: n.replace(/ /g, ''), tok, tokSet: new Set(tok)
      };
    });
    const sorted = all.slice().sort((a, b) => a.amt - b.amt || a.idx - b.idx);
    let minDay = Infinity, maxDay = -Infinity;
    all.forEach(c => { const d = dayNum(c.date); if (d < minDay) minDay = d; if (d > maxDay) maxDay = d; });
    return { sorted, all, minDay, maxDay };
  }

  function lowerBound(arr, v) {
    let lo = 0, hi = arr.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (arr[mid].amt < v) lo = mid + 1; else hi = mid;
    }
    return lo;
  }

  // Levenshtein dengan batas; > max berarti "terlalu jauh"
  function lev(a, b, max) {
    if (Math.abs(a.length - b.length) > max) return max + 1;
    let prev = Array.from({ length: b.length + 1 }, (_, i) => i);
    for (let i = 1; i <= a.length; i++) {
      const cur = [i];
      let rowMin = i;
      for (let j = 1; j <= b.length; j++) {
        const v = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
        cur[j] = v;
        if (v < rowMin) rowMin = v;
      }
      if (rowMin > max) return max + 1;
      prev = cur;
    }
    return prev[b.length];
  }

  function nameProfile(note3) {
    const raw = norm(String(note3).replace(/\n/g, ' '));
    const all = raw.split(' ').filter(Boolean);
    let sig = all.filter(t => t.length >= 2 && !TITLES.has(t));
    if (!sig.length) sig = all.filter(t => t.length >= 2);
    return { compact: raw.replace(/ /g, ''), sig };
  }

  // seberapa mirip satu kata nama dengan keterangan rekening (0..1)
  function tokStrength(t, c) {
    if (c.tokSet.has(t)) return 1;
    if (t.length >= 4 && c.nsq.includes(t)) return 0.9;     // "BUDISANTOSO"
    let best = 0;
    for (const k of c.tok) {
      if (k.length >= 3 && t.length >= 3 && (k.startsWith(t) || t.startsWith(k))) best = Math.max(best, 0.8); // MUH ~ MUHAMMAD
      else if (t.length >= 4 && k.length >= 4) {
        const lim = t.length >= 7 ? 2 : 1;                   // MOHAMMAD ~ MUHAMMAD
        if (lev(t, k, lim) <= lim) best = Math.max(best, 0.7);
      }
    }
    return best;
  }

  function nameScore(nm, c) {
    if (!nm.sig.length) return { score: 0, anchor: false };
    if (nm.compact.length >= 4 && c.nsq.includes(nm.compact)) return { score: 1, anchor: true };
    let sum = 0, anchor = false;
    for (const t of nm.sig) {
      const s = tokStrength(t, c);
      sum += s;
      if (s >= 0.7 && t.length >= 3) anchor = true;
    }
    return { score: sum / nm.sig.length, anchor };
  }

  // Format Notes: "3.000.000 (bagi 2 akun)" -> satu transfer rekening dibagi ke beberapa akun.
  // nominal = jumlah transfer asli di rekening koran, n = jumlah akun yang membagi transfer itu.
  function parseSplit(notes) {
    const m = String(notes == null ? '' : notes)
      .match(/^\s*(?:rp\.?\s*)?(\d[\d.,]*)\s*\(\s*bagi\s*(\d+)\s*(?:akun|account|acc|rek|rekening)?\s*\)/i);
    if (!m) return null;
    const nominal = parseAmt(m[1].replace(/[.,]$/, ''));
    const n = parseInt(m[2], 10);
    if (!isFinite(nominal) || nominal <= 0 || !(n >= 1)) return null;
    return { nominal, n };
  }

  function targetOf(notes, coins) {
    const sp = parseSplit(notes);
    if (sp) return sp.nominal;
    if (notes) {
      const m = String(notes).match(/[\d.]+/);
      const digits = m ? m[0].replace(/\./g, '') : '';
      return digits ? parseFloat(digits) : 0;
    }
    return coins * 100;
  }

  // Aturan awal (ketat): nama utuh, kata pertama, kata terakhir
  function strictName(cands, note3) {
    const full = squash(note3);
    if (full) {
      const hit = cands.find(c => c.sq.includes(full));
      if (hit) return hit;
    }
    const words = String(note3).replace(/\n/g, ' ').trim().split(/\s+/).filter(Boolean);
    for (const w of [words[0], words[words.length - 1]]) {
      if (!w) continue;
      const lw = w.toLowerCase();
      // kata 1-2 huruf hanya dihitung bila berdiri sendiri, supaya "A" tidak cocok ke semua baris
      const hit = cands.find(c => lw.length < 3 ? c.words.has(lw) : c.low.includes(lw));
      if (hit) return hit;
    }
    return null;
  }

  function bestByName(cands, nm, minScore, t, used) {
    let best = null;
    for (const c of cands) {
      if (used && used.has(c.idx)) continue;
      const s = nameScore(nm, c);
      if (!s.anchor || s.score < minScore) continue;
      const dd = t ? Math.abs(dayNum(c.date) - dayNum(t)) : 0;
      if (!best || s.score > best.s || (s.score === best.s && (dd < best.dd || (dd === best.dd && c.idx < best.c.idx)))) {
        best = { c, s: s.score, dd };
      }
    }
    return best && best.c;
  }

  // 2 atau 3 transfer dari pengirim yang sama yang jumlahnya pas
  function findSum(index, nm, target, used) {
    const hi = target + 999;
    const pool = [];
    for (let i = 0; i < index.sorted.length && index.sorted[i].amt <= hi; i++) {
      const c = index.sorted[i];
      if (c.amt <= 0 || used.has(c.idx)) continue;
      const s = nameScore(nm, c);
      if (s.anchor && s.score >= 0.5) pool.push({ c, s: s.score });
    }
    pool.sort((a, b) => b.s - a.s || a.c.idx - b.c.idx);
    pool.length = Math.min(pool.length, 40);
    const ok = v => { v = round2(v); return v >= target && v <= hi; };
    const make = parts => ({
      idx: parts[0].idx, idxs: parts.map(p => p.idx), ket: parts.map(p => p.ket).join(' + '),
      amt: round2(parts.reduce((s, p) => s + p.amt, 0)), date: parts[0].date, tier: 'sum'
    });
    for (let i = 0; i < pool.length; i++) {
      for (let j = i + 1; j < pool.length; j++) {
        if (ok(pool[i].c.amt + pool[j].c.amt)) return make([pool[i].c, pool[j].c]);
      }
    }
    for (let i = 0; i < pool.length; i++) {
      for (let j = i + 1; j < pool.length; j++) {
        for (let k = j + 1; k < pool.length; k++) {
          if (ok(pool[i].c.amt + pool[j].c.amt + pool[k].c.amt)) return make([pool[i].c, pool[j].c, pool[k].c]);
        }
      }
    }
    return null;
  }

  // Cari baris rekening koran. Bertingkat, dari yang paling pasti:
  //  1 exact  : nominal [target, target+999] + nama (aturan awal)
  //  2 name   : nominal sama + nama fleksibel (urutan/ejaan/singkatan/typo)
  //  3 amount : nominal agak berbeda (+-10%) + nama cocok
  //  4 sum    : beberapa transfer dari pengirim yang sama yang jumlahnya pas
  //  5 loose  : nama sangat cocok + tanggal berdekatan, nominal bebas
  // Tingkat 3-5 hanya dipakai bila deep = true dan tidak memakai transfer yang sudah dipakai baris lain.
  function findMatch(index, note3, notes, coins, opts = {}) {
    const target = targetOf(notes, coins);
    const arr = index.sorted;
    const win = (lo, hi) => {
      const out = [];
      for (let i = lowerBound(arr, lo); i < arr.length && arr[i].amt <= hi; i++) out.push(arr[i]);
      return out.sort((a, b) => a.idx - b.idx); // urutan asli rekening koran
    };
    const tag = (c, tier, amtOff) => Object.assign(c, { tier, amtOff: !!amtOff });

    const exact = win(target, target + 999);
    if (exact.length) {
      const h = strictName(exact, note3);
      if (h) return tag(h, 'exact', false);
    }
    const nm = nameProfile(note3);
    if (exact.length) {
      const h = bestByName(exact, nm, 0.5, opts.t, null);
      if (h) return tag(h, 'name', false);
    }
    if (!opts.deep || !nm.sig.length) return null;

    const used = opts.used || new Set();
    const tol = Math.max(target * 0.1, 2000);
    const wide = win(target - tol, target + 999 + tol);
    let h = bestByName(wide, nm, 0.6, opts.t, used);
    if (h) return tag(h, 'amount', !(h.amt >= target && h.amt <= target + 999));

    const sum = findSum(index, nm, target, used);
    if (sum) return sum;

    let best = null;
    for (const c of index.all) {
      if (used.has(c.idx)) continue;
      const s = nameScore(nm, c);
      if (!s.anchor || s.score < 0.8) continue;
      const dd = opts.t ? Math.abs(dayNum(c.date) - dayNum(opts.t)) : 0;
      if (opts.t && dd > 5) continue;
      const ad = Math.abs(c.amt - target);
      if (!best || s.score > best.s || (s.score === best.s && (dd < best.dd || (dd === best.dd && ad < best.ad)))) {
        best = { c, s: s.score, dd, ad };
      }
    }
    if (best) return tag(best.c, 'loose', true);
    return null;
  }

  // Petunjuk untuk baris yang tetap tidak ketemu
  function suggest(index, note3, target, t) {
    const nm = nameProfile(note3);
    const fmt = x => `- ${fmtDate(x.c.date)} | ${nfRp(x.c.amt)} | ${x.c.ket.slice(0, 90)}`;
    const scored = index.all.map(c => ({ c, s: nm.sig.length ? nameScore(nm, c).score : 0, ad: Math.abs(c.amt - target) }));
    const byName = scored.filter(x => x.s >= 0.34).sort((a, b) => b.s - a.s || a.ad - b.ad).slice(0, 3);
    const byAmt = scored.filter(x => x.ad <= Math.max(target * 0.2, 3000)).sort((a, b) => a.ad - b.ad).slice(0, 3);
    const out = [];
    if (t && isFinite(index.minDay) && (dayNum(t) < index.minDay - 1 || dayNum(t) > index.maxDay + 1)) {
      const f = d => fmtDate(fromUTCms(d * 86400000));
      out.push(`Tanggal top up di luar periode rekening koran yang diunggah (${f(index.minDay)} sampai ${f(index.maxDay)}). Unggah rekening koran periode tersebut.`);
    }
    if (byName.length) out.push('Nama mirip:\n' + byName.map(fmt).join('\n'));
    if (byAmt.length) out.push('Nominal dekat:\n' + byAmt.map(fmt).join('\n'));
    if (!out.length) out.push('Tidak ada transfer dengan nama maupun nominal yang mirip di rekening koran.');
    return 'Tidak ditemukan di rekening koran.\n' + out.join('\n');
  }

  // Hasil satu baris: status, tanggal rekening, keterangan, nominal, coin, selisih
  function computeRow(row, index, opts = {}) {
    const deep = opts.deep !== false;
    const E = Number(row.src['Coins']) || 0;
    const note3 = String(row.note3 == null ? '' : row.note3);
    const key = note3.trim().toLowerCase();
    const notes = String(row.notes == null ? '' : row.notes).trim();

    // Tanggal Rekening Koran selalu = tanggal Recharge time (kolom O), format DD/MM/YYYY
    const R = row.t ? mk(row.t.y, row.t.mo, row.t.d) : null;

    if (!key) return { status: 'empty', R: null, S: '', T: '', V: '', W: '', hits: [] };
    if (key === 'komisi' || key === 'ro reseller') {
      return { status: 'special', R, Rfull: false, S: note3, T: 0, V: 0, W: E, hits: [] };
    }
    // Notes "nominal (bagi N akun)": transfer di rekening koran dipakai bersama beberapa akun,
    // jadi transfer yang sudah dipakai akun lain tetap boleh dipilih lagi.
    const split = parseSplit(notes);
    const hit = findMatch(index, note3, notes, E, { t: row.t, deep, used: split ? new Set() : opts.used });
    if (!hit) {
      const out = { status: 'miss', R, Rfull: false, S: note3, T: 0, V: 0, W: E, hits: [] };
      if (deep) out.hint = suggest(index, note3, targetOf(notes, E), row.t);
      return out;
    }
    let T = E * 100;
    if (split) {
      T = E * 100;               // kolom Nominal selalu Coins x 100; Keterangan mengikuti nominal awal di Notes
    } else if (hit.amtOff) {
      T = Math.round(hit.amt);   // nominal sebenarnya di rekening koran, supaya selisihnya jujur
    } else if (notes) {
      const q = notes.match(/\(([\d.]+)\)/);
      if (q) { const n = parseFloat(q[1].replace(/\./g, '')); if (!isNaN(n)) T = n; }
    }
    const V = T / 100;
    const how = {
      name: 'Nama dicocokkan secara fleksibel (urutan, ejaan, atau singkatan berbeda). Mohon dicek.',
      amount: `Nama cocok, tetapi nominal di rekening koran ${nfRp(hit.amt)} berbeda dari coin x 100. Kolom Nominal diisi nominal sebenarnya.`,
      sum: `Gabungan ${(hit.idxs || []).length} transfer dengan total ${nfRp(hit.amt)}.`,
      loose: `Hanya nama dan tanggal yang cocok, nominal di rekening koran ${nfRp(hit.amt)}. Mohon dicek.`
    }[hit.tier] || '';
    const splitNote = split ? `Transfer ${nfRp(split.nominal)} dibagi ${split.n} akun. Kolom Nominal = Coins x 100.` : '';
    return {
      status: 'match', R, Rfull: false, S: hit.ket, T, V, W: round2(E - V), split: !!split, splitNote,
      hits: hit.idxs || [hit.idx], tier: hit.tier, verify: hit.tier !== 'exact', how
    };
  }

  // transfer rekening koran yang sudah dipakai baris lain
  function usedBy(rows, skip) {
    const used = new Set();
    rows.forEach(r => { if (r !== skip && r.res && r.res.hits) r.res.hits.forEach(i => used.add(i)); });
    return used;
  }

  // Dua putaran: putaran 1 aturan ketat untuk semua baris, putaran 2 pencarian lanjutan
  // hanya untuk yang belum ketemu (dengan transfer yang masih bebas).
  function computeAll(rows, index) {
    rows.forEach(r => { r.res = computeRow(r, index, { deep: false }); });
    const used = usedBy(rows, null);
    rows.forEach(r => {
      if (r.res.status !== 'miss') return;
      r.res = computeRow(r, index, { deep: true, used });
      r.res.hits.forEach(i => used.add(i));
    });
  }

  const api = {
    HEADERS, pad, mk, parseTime, fmtDateTime, fmtDate, fmtDateFull,
    parseRekening, mergeCredits, parseTopupAOA, mergeTopups,
    buildIndex, findMatch, computeRow, computeAll, usedBy, parseSplit
  };
  return api;
})();

if (typeof module !== 'undefined' && module.exports) module.exports = Core;

/* ---------------------------------------------------------
   BAGIAN 2 – ANTARMUKA (hanya jalan di browser)
   --------------------------------------------------------- */
if (typeof document !== 'undefined') {
  (() => {
    'use strict';

    const { HEADERS, fmtDateTime, fmtDate, fmtDateFull } = Core;
    const PAGE_SIZE = 100;
    const $ = s => document.querySelector(s);

    const state = {
      slots: { rekening: null, topup: null },
      index: null,
      rows: [],
      filter: 'all',
      q: '',
      page: 1,
      admin: 'REREN'
    };

    /* ---------- util ---------- */
    const esc = v => String(v == null ? '' : v)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    const nf = n => (n === '' || n == null || Number.isNaN(Number(n))) ? '' : Number(n).toLocaleString('id-ID');
    const bytes = n => n > 1048576 ? (n / 1048576).toFixed(1) + ' MB' : Math.max(1, Math.round(n / 1024)) + ' KB';

    function toast(msg, type = '') {
      const el = document.createElement('div');
      el.className = 'toast ' + type;
      el.textContent = msg;
      const box = $('#toasts');
      box.appendChild(el);
      while (box.children.length > 3) box.firstChild.remove();
      setTimeout(() => el.remove(), type === 'err' ? 7000 : 3500);
    }

    /* ---------- membaca & mengenali file ---------- */
    const readBuffer = file => new Promise((resolve, reject) => {
      const fr = new FileReader();
      fr.onload = () => resolve(fr.result);
      fr.onerror = () => reject(fr.error);
      fr.readAsArrayBuffer(file);
    });

    function sheetToAOA(wb) {
      const out = [];
      for (const name of wb.SheetNames) {
        out.push(XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: '' }));
      }
      return out;
    }

    function tryTopup(wbOrText, isText) {
      if (typeof XLSX === 'undefined') throw new Error('Library XLSX belum termuat. Periksa koneksi internet lalu muat ulang halaman.');
      const wb = isText ? XLSX.read(wbOrText, { type: 'string' }) : wbOrText;
      for (const aoa of sheetToAOA(wb)) {
        const rows = Core.parseTopupAOA(aoa);
        if (rows) return rows;
      }
      return null;
    }

    async function classify(file) {
      const buf = await readBuffer(file);
      const u8 = new Uint8Array(buf.slice(0, 4));
      const isZip = u8[0] === 0x50 && u8[1] === 0x4b;
      const isOle = u8[0] === 0xd0 && u8[1] === 0xcf;

      if (isZip || isOle) {
        if (typeof XLSX === 'undefined') throw new Error('Library XLSX belum termuat. Periksa koneksi internet lalu muat ulang halaman.');
        const wb = XLSX.read(buf, { type: 'array' });
        const rows = tryTopup(wb, false);
        if (rows) return { kind: 'topup', rows };
        return { kind: 'unknown' };
      }

      const text = new TextDecoder('utf-8').decode(buf).replace(/^\uFEFF/, '');
      if (/^\s*'?\d{2}\/\d{2}\/\d{4}\s*,/m.test(text) || /No\. Rekening/i.test(text)) {
        const parsed = Core.parseRekening(text);
        if (parsed.credits.length) return { kind: 'rekening', parsed };
      }
      if (/agency mico id/i.test(text)) {
        const rows = tryTopup(text, true);
        if (rows) return { kind: 'topup', rows };
      }
      return { kind: 'unknown' };
    }

    /* ---------- slot upload ---------- */
    function slotEl(slot) { return $('#drop-' + slot); }

    function paintSlot(slot) {
      const el = slotEl(slot);
      const data = state.slots[slot];
      const stateEl = el.querySelector('.drop-state');
      const clear = el.querySelector('.drop-clear');
      el.classList.remove('err');
      if (!data) {
        el.classList.remove('has');
        stateEl.textContent = '';
        clear.hidden = true;
        return;
      }
      el.classList.add('has');
      clear.hidden = false;
      const names = data.files.map(f => `${f.name} (${bytes(f.size)})`).join('\n');
      if (slot === 'rekening') {
        const credits = Core.mergeCredits(data.files.map(f => f.parsed));
        const debit = data.files.reduce((n, f) => n + f.parsed.debit, 0);
        const skipped = data.files.reduce((n, f) => n + f.parsed.skipped, 0);
        const nama = (data.files.find(f => f.parsed.nama) || { parsed: {} }).parsed.nama;
        stateEl.textContent = `${names}\n${credits.length} transaksi kredit, ${debit} debit diabaikan` +
          (skipped ? `\nPERHATIAN: ${skipped} baris tidak terbaca (lihat Console)` : '') +
          (nama ? `\nRekening ${nama}` : '');
        data.files.forEach(f => { if (f.parsed.skipped) console.warn(`${f.name}: baris tidak terbaca`, f.parsed.skippedLines); });
      } else {
        const rows = Core.mergeTopups(data.files.map(f => f.rows));
        const first = rows[0] && rows[0].t, last = rows[rows.length - 1] && rows[rows.length - 1].t;
        stateEl.textContent = `${names}\n${rows.length} baris top up` +
          (first && last ? `\n${fmtDateTime(first).slice(0, 16)} sampai ${fmtDateTime(last).slice(0, 16)}` : '');
      }
      stateEl.style.whiteSpace = 'pre-line';
    }

    async function handleFiles(fileList, hintSlot) {
      const files = Array.from(fileList || []);
      if (!files.length) return;
      for (const file of files) {
        try {
          const res = await classify(file);
          if (res.kind === 'unknown') {
            const el = slotEl(hintSlot || 'rekening');
            el.classList.add('err');
            el.querySelector('.drop-state').textContent = `${file.name}: format tidak dikenali`;
            toast(`${file.name}: bukan rekening koran (CSV bank) maupun laporan top up (kolom "Agency Mico ID" tidak ditemukan).`, 'err');
            continue;
          }
          if (hintSlot && res.kind !== hintSlot) {
            toast(`${file.name} dikenali sebagai ${res.kind === 'topup' ? 'laporan top up' : 'rekening koran'}, dipindahkan ke kotak yang sesuai.`);
          }
          const entry = { name: file.name, size: file.size, ...(res.kind === 'topup' ? { rows: res.rows } : { parsed: res.parsed }) };
          const cur = state.slots[res.kind] || { files: [] };
          const k = cur.files.findIndex(f => f.name === file.name);
          if (k >= 0) cur.files[k] = entry; else cur.files.push(entry);   // nama sama = ganti, nama beda = gabung
          state.slots[res.kind] = cur;
          paintSlot(res.kind);
        } catch (err) {
          console.error(err);
          toast(`${file.name}: ${err.message || 'gagal dibaca'}`, 'err');
        }
      }
      process();
    }

    function clearSlot(slot) {
      state.slots[slot] = null;
      slotEl(slot).querySelector('input').value = '';
      paintSlot(slot);
      process();
    }

    /* ---------- proses gabungan ---------- */
    function process() {
      const { rekening, topup } = state.slots;
      if (!rekening || !topup) {
        state.rows = [];
        state.index = null;
        state.page = 1;
        render(false);
        return;
      }
      state.index = Core.buildIndex(Core.mergeCredits(rekening.files.map(f => f.parsed)));
      state.rows = Core.mergeTopups(topup.files.map(f => f.rows))
        .map(r => ({ src: r.src, t: r.t, note3: r.note3, notes: '', res: null }));
      Core.computeAll(state.rows, state.index);
      state.page = 1;
      render(true);
      const s = summary();
      toast(`${s.total} baris diproses: ${s.match} cocok` + (s.approx ? ` (${s.approx} perlu verifikasi)` : '') +
        `, ${s.special} komisi/RO, ${s.check} perlu dicek.`, 'ok');
      if (s.empty > 0) openFill();
    }

    /* ---------- pop up: isi Note3 yang kosong ---------- */
    const emptyIdx = () => state.rows.map((r, i) => (r.res.status === 'empty' ? i : -1)).filter(i => i >= 0);

    function updateFillBtn() {
      const n = emptyIdx().length;
      const b = $('#btnFill');
      b.hidden = n === 0;
      b.textContent = `Isi Note3 kosong (${n})`;
    }

    function openFill() {
      const list = emptyIdx();
      if (!list.length) return;
      $('#fillBody').innerHTML = list.map(i => {
        const r = state.rows[i];
        return `<tr data-i="${i}">` +
          `<td class="ro">${esc(r.src['Target Mico ID'])}</td>` +
          `<td class="ro">${esc(r.src['Target name'])}</td>` +
          `<td><input class="fill-in" type="text" autocomplete="off" spellcheck="false" placeholder="nama pengirim" aria-label="Note3 untuk ${esc(r.src['Target name'])}"></td>` +
          `<td class="ro">${esc(fmtDateTime(r.t))}</td></tr>`;
      }).join('');
      $('#fillCount').textContent = `${list.length} baris perlu diisi`;
      $('#fillModal').hidden = false;
      const first = $('#fillBody input');
      if (first) first.focus();
    }

    function closeFill() { $('#fillModal').hidden = true; }

    function submitFill() {
      let filled = 0, matched = 0, notFound = 0;
      document.querySelectorAll('#fillBody tr').forEach(tr => {
        const v = tr.querySelector('input').value.trim();
        if (!v) return;
        const r = state.rows[Number(tr.dataset.i)];
        r.note3 = v;                                  // masuk ke kolom Note3 (M)
        r.res = Core.computeRow(r, state.index, { used: Core.usedBy(state.rows, r) });   // cari lagi di rekening koran
        filled++;
        if (r.res.status === 'match' || r.res.status === 'special') matched++; else notFound++;
      });
      closeFill();
      if (!filled) { toast('Belum ada Note3 yang diisi.', 'err'); return; }
      render(true);
      const left = emptyIdx().length;
      toast(`${filled} Note3 diisi: ${matched} ketemu di rekening koran, ${notFound} belum ketemu.` + (left ? ` Masih ada ${left} kosong.` : ''), 'ok');
    }

    function summary() {
      const s = { total: state.rows.length, match: 0, special: 0, miss: 0, empty: 0, check: 0, diff: 0, approx: 0 };
      for (const r of state.rows) {
        const st = r.res.status;
        s[st]++;
        if (st === 'match' && r.res.verify) s.approx++;
        if (st !== 'empty') s.diff += Number(r.res.W) || 0;
      }
      s.check = s.miss + s.empty;
      s.diff = Math.round(s.diff * 100) / 100;
      return s;
    }

    function updateReadout() {
      const s = summary();
      $('#m-total').textContent = nf(s.total) || '0';
      $('#m-match').textContent = nf(s.match) || '0';
      $('#m-special').textContent = nf(s.special) || '0';
      $('#m-miss').textContent = nf(s.check) || '0';
      $('#m-diff').textContent = nf(s.diff) || '0';
    }

    /* ---------- tabel ---------- */
    function passes(r) {
      const st = r.res.status;
      if (state.filter === 'match' && st !== 'match') return false;
      if (state.filter === 'special' && st !== 'special') return false;
      if (state.filter === 'approx' && !(st === 'match' && r.res.verify)) return false;
      if (state.filter === 'check' && st !== 'miss' && st !== 'empty') return false;
      if (state.filter === 'diff' && !(st !== 'empty' && Number(r.res.W) !== 0)) return false;
      if (state.q) {
        const hay = [r.src['Target name'], r.src['Target Mico ID'], r.note3, r.res.S, r.notes].join(' ').toLowerCase();
        if (!hay.includes(state.q)) return false;
      }
      return true;
    }

    function resCells(r) {
      const x = r.res;
      const rTxt = x.R ? (x.Rfull ? fmtDateFull(x.R) : fmtDate(x.R)) : '';
      const w = x.W === '' ? '' : nf(x.W);
      const wCls = x.status === 'empty' ? '' : (Number(x.W) === 0 ? 'z' : 'nz');
      return { rTxt, w, wCls };
    }

    function sTitle(x) {
      if (x.status === 'miss' && x.hint) return x.hint;
      if (x.verify && x.how) return x.S + (x.splitNote ? '\n\n' + x.splitNote : '') + '\n\nPERLU VERIFIKASI: ' + x.how;
      if (x.splitNote) return x.S + '\n\n' + x.splitNote;
      return x.S;
    }
    const rowCls = x => 'is-' + x.status + (x.verify ? ' is-verify' : '');

    function rowHTML(r, i) {
      const s = r.src;
      const { rTxt, w, wCls } = resCells(r);
      const x = r.res;
      return `<tr data-i="${i}" class="${rowCls(x)}">` +
        `<td class="x num dim">${esc(s['Agency Mico ID'])}</td>` +
        `<td class="x">${esc(s['Agency name'])}</td>` +
        `<td class="num dim">${esc(s['Target Mico ID'])}</td>` +
        `<td>${esc(s['Target name'])}</td>` +
        `<td class="num">${esc(nf(s['Coins']))}</td>` +
        `<td class="x">${esc(s['Currency'])}</td>` +
        `<td class="num">${esc(nf(s['Currency amount']))}</td>` +
        `<td class="x num">${esc(s['Value to USD'])}</td>` +
        `<td class="x num">${esc(nf(s['Value (1USD)']))}</td>` +
        `<td class="x">${esc(s['Target is agent'])}</td>` +
        `<td class="x">${esc(s['Note1'])}</td>` +
        `<td class="x">${esc(s['Note2'])}</td>` +
        `<td><input class="cell-in note3" data-f="note3" value="${esc(r.note3)}" placeholder="isi nama" aria-label="Note3"></td>` +
        `<td class="x">${esc(s['Country'])}</td>` +
        `<td>${esc(fmtDateTime(r.t))}</td>` +
        `<td class="x">${esc(s['Transaction Status'])}</td>` +
        `<td class="gap"></td>` +
        `<td data-c="R">${esc(rTxt)}</td>` +
        `<td data-c="S" title="${esc(sTitle(x))}">${esc(x.S)}</td>` +
        `<td class="num" data-c="T">${esc(nf(x.T))}</td>` +
        `<td class="gap"></td>` +
        `<td class="num" data-c="V">${esc(nf(x.V))}</td>` +
        `<td class="num ${wCls}" data-c="W">${esc(w)}</td>` +
        `<td class="admin">${esc(state.admin)}</td>` +
        `<td><input class="cell-in" data-f="notes" value="${esc(r.notes)}" placeholder="" aria-label="Notes"></td>` +
        `</tr>`;
    }

    function paintRow(tr, r) {
      const x = r.res;
      const { rTxt, w, wCls } = resCells(r);
      tr.className = rowCls(x);
      const set = (c, txt) => { const td = tr.querySelector(`td[data-c="${c}"]`); if (td) { td.textContent = txt; return td; } };
      set('R', rTxt);
      const sTd = set('S', x.S); if (sTd) sTd.title = sTitle(x);
      set('T', nf(x.T));
      set('V', nf(x.V));
      const wTd = set('W', w); if (wTd) wTd.className = 'num ' + wCls;
    }

    function render(sweep) {
      const wrap = $('#tableWrap');
      const hasRows = state.rows.length > 0;
      wrap.classList.toggle('filled', hasRows);
      $('#btnDownload').disabled = !hasRows;
      updateReadout();
      updateFillBtn();

      const list = [];
      state.rows.forEach((r, i) => { if (passes(r)) list.push(i); });
      const pages = Math.max(1, Math.ceil(list.length / PAGE_SIZE));
      if (state.page > pages) state.page = pages;
      const start = (state.page - 1) * PAGE_SIZE;
      const slice = list.slice(start, start + PAGE_SIZE);

      $('#tbody').innerHTML = slice.map(i => rowHTML(state.rows[i], i)).join('');

      const pager = $('#pager');
      pager.hidden = !hasRows;
      if (hasRows) {
        $('#pageInfo').textContent = list.length
          ? `Menampilkan ${start + 1}–${start + slice.length} dari ${list.length} baris` + (list.length !== state.rows.length ? ` (total ${state.rows.length})` : '')
          : 'Tidak ada baris yang cocok dengan filter';
        $('#pageNum').textContent = `${state.page} / ${pages}`;
        $('#prev').disabled = state.page <= 1;
        $('#next').disabled = state.page >= pages;
      }

      if (sweep && hasRows) {
        wrap.style.setProperty('--sweep-to', wrap.clientWidth + 'px');
        wrap.classList.remove('sweep');
        void wrap.offsetWidth;
        wrap.classList.add('sweep');
        setTimeout(() => wrap.classList.remove('sweep'), 1000);
      }
    }

    /* ---------- ekspor Excel ---------- */
    async function downloadXlsx() {
      if (typeof ExcelJS === 'undefined') {
        toast('Library ExcelJS belum termuat. Periksa koneksi internet lalu muat ulang halaman.', 'err');
        return;
      }
      if (!state.rows.length) return;

      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('MICO', { views: [{ state: 'frozen', ySplit: 2 }] });
      const widths = [15.5, 13.75, 14.5, 24.5, 9, 10.25, 16.63, 13.38, 13.5, 14.5, 8.13, 8.13, 27.88, 9.63, 21, 17.38,
        3, 21.5, 80, 13.13, 3, 10.63, 8.38, 18.75, 29.88];
      ws.columns = widths.map(width => ({ width }));

      const thin = { style: 'thin', color: { argb: 'FF999999' } };
      const box = { top: thin, left: thin, bottom: thin, right: thin };
      const center = { horizontal: 'center', vertical: 'middle' };

      // baris 1: judul grup
      const groups = [
        ['A1:P1', 'DASHBOARD MICO', 'FFFFFF00'],
        ['R1:T1', 'REKENING KORAN', 'FFFFC000'],
        ['V1:W1', 'Check Selisih', 'FF92D050'],
        ['X1:Y1', 'PIC (Personal In Charge)', 'FF9FC5E8']
      ];
      for (const [range, label, color] of groups) {
        ws.mergeCells(range);
        const c = ws.getCell(range.split(':')[0]);
        c.value = label;
        c.font = { bold: true };
        c.alignment = center;
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: color } };
        c.border = box;
      }

      // baris 2: header kolom
      const head = [...HEADERS];
      const heads = {};
      head.forEach((h, i) => { heads[i + 1] = h; });
      Object.assign(heads, { 18: 'Tanggal Rekening Koran', 19: 'Keterangan', 20: 'Nominal', 22: 'Coin', 23: 'Selisih', 24: 'Nama Admin', 25: 'Notes' });
      for (const [col, label] of Object.entries(heads)) {
        const c = ws.getCell(2, Number(col));
        c.value = label;
        c.font = { bold: true };
        c.alignment = center;
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC0C0C0' } };
        c.border = box;
      }

      // data
      const numCols = ['Agency Mico ID', 'Target Mico ID', 'Coins', 'Currency amount', 'Value to USD', 'Value (1USD)'];
      const asDate = t => new Date(t.ts); // ts sudah "jam dinding" dalam UTC, jadi Excel menampilkan jam yang sama
      state.rows.forEach((r, i) => {
        const row = ws.getRow(i + 3);
        HEADERS.forEach((h, j) => {
          let v = r.src[h];
          if (h === 'Note3') v = r.note3;
          if (h === 'Recharge time(GMT+7)') v = r.t ? asDate(r.t) : '';
          else if (numCols.includes(h)) { const n = Number(v); v = v === '' || isNaN(n) ? v : n; }
          row.getCell(j + 1).value = v === '' ? null : v;
        });
        row.getCell(15).numFmt = 'yyyy-mm-dd h:mm:ss';

        const x = r.res;
        if (x.status !== 'empty') {
          if (x.R) {
            row.getCell(18).value = asDate(x.R);
            row.getCell(18).numFmt = x.Rfull ? 'dd/mm/yyyy hh:mm:ss' : 'dd/mm/yyyy';
          }
          row.getCell(19).value = x.S;
          if (x.verify) {
            row.getCell(19).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
            if (x.how) row.getCell(19).note = x.how;
          }
          if (x.status === 'miss' && x.hint) row.getCell(19).note = x.hint;
          row.getCell(20).value = x.T;
          row.getCell(22).value = x.V;
          row.getCell(23).value = x.W;
          [20, 22, 23].forEach(c => { row.getCell(c).numFmt = '#,##0'; });
          if (x.status === 'miss') row.getCell(19).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF4CCCC' } };
        } else {
          row.getCell(13).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
        }
        row.getCell(24).value = state.admin;
        row.getCell(25).value = r.notes || null;
      });

      ws.autoFilter = { from: { row: 2, column: 1 }, to: { row: state.rows.length + 2, column: 25 } };

      // O3 = Recharge time baris pertama -> "30 Oktober 2026"
      // (tanda "/" tidak boleh dipakai di nama file, jadi dipisah spasi)
      const BULAN = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
      const first = state.rows[0] && state.rows[0].t;
      const stamp = first ? `${Core.pad(first.d)} ${BULAN[first.mo - 1]} ${first.y}` : 'hasil';
      const buf = await wb.xlsx.writeBuffer();
      const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `Lembar Kerja Reren - ${stamp}.xlsx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 2000);
      toast('File Excel siap diunduh.', 'ok');
    }

    /* ---------- event ---------- */
    document.querySelectorAll('.drop').forEach(el => {
      const slot = el.dataset.slot;
      const input = el.querySelector('input[type="file"]');
      input.addEventListener('change', () => { handleFiles(input.files, slot); input.value = ''; });
      ['dragenter', 'dragover'].forEach(ev => el.addEventListener(ev, e => { e.preventDefault(); el.classList.add('over'); }));
      ['dragleave', 'drop'].forEach(ev => el.addEventListener(ev, e => { e.preventDefault(); el.classList.remove('over'); }));
      el.addEventListener('drop', e => { handleFiles(e.dataTransfer.files, slot); });
      el.querySelector('.drop-clear').addEventListener('click', e => { e.preventDefault(); e.stopPropagation(); clearSlot(slot); });
    });
    // cegah browser membuka file yang dijatuhkan di luar kotak upload
    ['dragover', 'drop'].forEach(ev => window.addEventListener(ev, e => e.preventDefault()));

    $('#btnReset').addEventListener('click', () => {
      state.slots.rekening = null;
      state.slots.topup = null;
      ['rekening', 'topup'].forEach(paintSlot);
      state.filter = 'all';
      state.q = '';
      $('#search').value = '';
      document.querySelectorAll('.tab').forEach(t => { const on = t.dataset.f === 'all'; t.classList.toggle('on', on); t.setAttribute('aria-pressed', on); });
      process();
    });

    $('#adminName').addEventListener('input', e => {
      state.admin = e.target.value.trim();
      document.querySelectorAll('#tbody td.admin').forEach(td => { td.textContent = state.admin; });
    });

    $('#tabs').addEventListener('click', e => {
      const b = e.target.closest('.tab');
      if (!b) return;
      state.filter = b.dataset.f;
      state.page = 1;
      document.querySelectorAll('.tab').forEach(t => { const on = t === b; t.classList.toggle('on', on); t.setAttribute('aria-pressed', on); });
      render(false);
    });

    let searchTimer;
    $('#search').addEventListener('input', e => {
      clearTimeout(searchTimer);
      searchTimer = setTimeout(() => { state.q = e.target.value.trim().toLowerCase(); state.page = 1; render(false); }, 150);
    });

    $('#prev').addEventListener('click', () => { state.page--; render(false); $('#tableWrap').scrollTop = 0; });
    $('#next').addEventListener('click', () => { state.page++; render(false); $('#tableWrap').scrollTop = 0; });
    $('#btnDownload').addEventListener('click', downloadXlsx);
    $('#btnFill').addEventListener('click', openFill);
    $('#fillCancel').addEventListener('click', closeFill);
    $('#fillSubmit').addEventListener('click', submitFill);
    $('#fillModal').addEventListener('keydown', e => {
      if (e.key === 'Escape') closeFill();
      if (e.key === 'Enter' && e.target.matches('.fill-in')) {
        e.preventDefault();
        const ins = Array.from(document.querySelectorAll('#fillBody .fill-in'));
        const next = ins[ins.indexOf(e.target) + 1];
        if (next) next.focus(); else $('#fillSubmit').focus();
      }
    });

    // kolom ringkas: sembunyikan kolom top up yang jarang dipakai (hanya tampilan, ekspor tetap lengkap)
    function applyCompact(on) {
      $('#grid').classList.toggle('compact', on);
      $('#btnCompact').setAttribute('aria-pressed', on);
      $('#btnCompact').classList.toggle('on', on);
      $('#gTop').colSpan = on ? 6 : 16;
    }
    $('#btnCompact').addEventListener('click', () => applyCompact($('#btnCompact').getAttribute('aria-pressed') !== 'true'));
    applyCompact(true);

    // edit Note3 / Notes langsung di tabel -> hitung ulang baris itu
    const tbody = $('#tbody');
    tbody.addEventListener('change', e => {
      const input = e.target.closest('input[data-f]');
      if (!input) return;
      const tr = input.closest('tr');
      const r = state.rows[Number(tr.dataset.i)];
      r[input.dataset.f] = input.value;
      r.res = Core.computeRow(r, state.index, { used: Core.usedBy(state.rows, r) });
      paintRow(tr, r);
      updateReadout();
      updateFillBtn();
    });
    tbody.addEventListener('keydown', e => {
      if (e.key === 'Enter' && e.target.matches('input[data-f]')) e.target.blur();
    });

    render(false);
  })();
}