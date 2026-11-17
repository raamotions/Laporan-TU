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
  const LINE_RE = /^'?(\d{2})\/(\d{2})\/(\d{4}),(.*),'?\d*,([\d.]+),(CR|DB),([\d.]+)\s*$/;

  function parseRekening(text) {
    const lines = String(text).split(/\r?\n/);
    const credits = [];
    let debit = 0, skipped = 0;
    for (const line of lines) {
      if (!/^'?\d{2}\/\d{2}\/\d{4},/.test(line)) continue;
      const m = line.match(LINE_RE);
      if (!m) { skipped++; continue; }
      if (m[6] === 'DB') { debit++; continue; }
      credits.push({
        idx: credits.length,
        date: mk(+m[3], +m[2], +m[1]),
        ket: m[4].trim(),
        amt: parseFloat(m[5])
      });
    }
    const nama = (String(text).match(/^Nama,=,(.*)$/m) || [])[1] || '';
    const rek = (String(text).match(/^No\. Rekening,=,'?(.*)$/m) || [])[1] || '';
    return { credits, debit, skipped, nama: nama.trim(), rek: rek.trim() };
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

  /* ---- pencocokan ---- */
  const squash = s => String(s).replace(/\s+/g, '').toLowerCase();

  function buildIndex(credits) {
    const sorted = credits.map(c => ({
      idx: c.idx, amt: c.amt, date: c.date, ket: c.ket,
      low: c.ket.toLowerCase(), sq: squash(c.ket)
    })).sort((a, b) => a.amt - b.amt || a.idx - b.idx);
    return { sorted };
  }

  function lowerBound(arr, v) {
    let lo = 0, hi = arr.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (arr[mid].amt < v) lo = mid + 1; else hi = mid;
    }
    return lo;
  }

  // Cari baris rekening koran: nominal dalam [target, target+999],
  // lalu nama = Note3 utuh, kalau gagal kata pertama, kalau gagal kata terakhir.
  function findMatch(index, note3, notes, coins) {
    let target;
    if (notes) {
      const m = String(notes).match(/[\d.]+/);
      const digits = m ? m[0].replace(/\./g, '') : '';
      target = digits ? parseFloat(digits) : 0;
    } else {
      target = coins * 100;
    }
    const arr = index.sorted;
    const cands = [];
    for (let i = lowerBound(arr, target); i < arr.length && arr[i].amt <= target + 999; i++) cands.push(arr[i]);
    if (!cands.length) return null;
    cands.sort((a, b) => a.idx - b.idx); // urutan asli rekening koran

    const full = squash(note3);
    if (full) {
      const hit = cands.find(c => c.sq.includes(full));
      if (hit) return hit;
    }
    const words = String(note3).replace(/\n/g, ' ').trim().split(/\s+/).filter(Boolean);
    for (const w of [words[0], words[words.length - 1]]) {
      if (!w) continue;
      const lw = w.toLowerCase();
      const hit = cands.find(c => c.low.includes(lw));
      if (hit) return hit;
    }
    return null;
  }

  // Hasil satu baris: status, tanggal rekening, keterangan, nominal, coin, selisih
  function computeRow(row, index) {
    const E = Number(row.src['Coins']) || 0;
    const note3 = String(row.note3 == null ? '' : row.note3);
    const key = note3.trim().toLowerCase();
    const notes = String(row.notes == null ? '' : row.notes).trim();

    if (!key) return { status: 'empty', R: null, S: '', T: '', V: '', W: '' };
    if (key === 'komisi' || key === 'ro reseller') {
      return { status: 'special', R: row.t, Rfull: true, S: note3, T: 0, V: 0, W: E };
    }
    const hit = findMatch(index, note3, notes, E);
    if (!hit) {
      return { status: 'miss', R: row.t, Rfull: true, S: note3, T: 0, V: 0, W: E };
    }
    let T = E * 100;
    if (notes) {
      const q = notes.match(/\(([\d.]+)\)/);
      if (q) { const n = parseFloat(q[1].replace(/\./g, '')); if (!isNaN(n)) T = n; }
    }
    return { status: 'match', R: hit.date, Rfull: false, S: hit.ket, T, V: T / 100, W: E - T / 100 };
  }

  const api = {
    HEADERS, pad, mk, parseTime, fmtDateTime, fmtDate, fmtDateFull,
    parseRekening, parseTopupAOA, buildIndex, findMatch, computeRow
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

      const text = new TextDecoder('utf-8').decode(buf);
      if (/^'?\d{2}\/\d{2}\/\d{4},/m.test(text) || /No\. Rekening/i.test(text)) {
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
      if (slot === 'rekening') {
        const p = data.parsed;
        stateEl.textContent =
          `${data.name} (${bytes(data.size)})\n` +
          `${p.credits.length} transaksi kredit, ${p.debit} debit diabaikan` +
          (p.nama ? `\nRekening ${p.nama}` : '');
      } else {
        const first = data.rows[0] && data.rows[0].t, last = data.rows[data.rows.length - 1] && data.rows[data.rows.length - 1].t;
        stateEl.textContent =
          `${data.name} (${bytes(data.size)})\n` +
          `${data.rows.length} baris top up` +
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
          state.slots[res.kind] = { name: file.name, size: file.size, ...(res.kind === 'topup' ? { rows: res.rows } : { parsed: res.parsed }) };
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
      state.index = Core.buildIndex(rekening.parsed.credits);
      state.rows = topup.rows.map(r => ({ src: r.src, t: r.t, note3: r.note3, notes: '', res: null }));
      state.rows.forEach(r => { r.res = Core.computeRow(r, state.index); });
      state.page = 1;
      render(true);
      const s = summary();
      toast(`${s.total} baris diproses: ${s.match} cocok, ${s.special} komisi/RO, ${s.check} perlu dicek.`, 'ok');
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
        r.res = Core.computeRow(r, state.index);      // cari lagi di rekening koran
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
      const s = { total: state.rows.length, match: 0, special: 0, miss: 0, empty: 0, check: 0, diff: 0 };
      for (const r of state.rows) {
        const st = r.res.status;
        s[st]++;
        if (st !== 'empty') s.diff += Number(r.res.W) || 0;
      }
      s.check = s.miss + s.empty;
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

    function rowHTML(r, i) {
      const s = r.src;
      const { rTxt, w, wCls } = resCells(r);
      const x = r.res;
      return `<tr data-i="${i}" class="is-${x.status}">` +
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
        `<td data-c="S" title="${esc(x.S)}">${esc(x.S)}</td>` +
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
      tr.className = 'is-' + x.status;
      const set = (c, txt) => { const td = tr.querySelector(`td[data-c="${c}"]`); if (td) { td.textContent = txt; return td; } };
      set('R', rTxt);
      const sTd = set('S', x.S); if (sTd) sTd.title = x.S;
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

      const first = state.rows[0] && state.rows[0].t;
      const stamp = first ? `${first.y}-${Core.pad(first.mo)}-${Core.pad(first.d)}` : 'hasil';
      const buf = await wb.xlsx.writeBuffer();
      const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `Hasil_Reren_${stamp}.xlsx`;
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
      r.res = Core.computeRow(r, state.index);
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