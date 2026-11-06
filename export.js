// Excel export in the center's own forms (штатное расписание + штатная расстановка), built from current data (with edits applied)
(function () {
  const DEC = 'Требует решения';
  const TNR = 'Times New Roman';
  const thin = { style: 'thin', color: { argb: 'FF000000' } };
  const BOX = { top: thin, left: thin, bottom: thin, right: thin };
  const FILL = c => ({ type: 'pattern', pattern: 'solid', fgColor: { argb: c } });
  const Y = FILL('FFFFF2CC'), R = FILL('FFFCE4D6'), G = FILL('FFEDEDED');
  const norm = t => String(t || '').toLowerCase().replace(/ё/g, 'е').replace(/[^а-яa-z]/g, '');
  const r2 = x => Math.round((x || 0) * 100) / 100;
  const disp = p => p.includes('·') ? p.split('·')[1].trim() + ' сектора контакт-центра' : p;
  const dlabel = d => d.replace(' · ', ' — ').replace('Кабинет разведения и подготовки ЛС', 'Кабинет разведения и подготовки лекарственных средств');
  function inferCat(p) {
    const l = p.toLowerCase();
    if (/^(директор|заместитель|руководитель|заведующ)/.test(l)) return 'Руководитель';
    if (l.startsWith('врач') || l.includes('ординатор') || l.includes('профессор') || l.includes('фармаколог')) return 'Врачебный персонал';
    if (/^(медицинская сестра|помощник врача|старшая медицинская|лаборант|старший лаборант)/.test(l)) return 'Средний медицинский персонал';
    if (/^(санитарка|сестра-хозяйка|швея|раздатчица)/.test(l)) return 'Младший медицинский персонал';
    if (/^(провизор|фармацевт)/.test(l)) return 'Фармацевтический персонал';
    return 'Специалист';
  }
  function sections(D, uid) {
    const u = D.units.find(x => x.id === uid), top = [], out = [];
    const isTop = (g, d, p) => d.name === 'Руководство' || (uid === 'ЦА' && p === d.name && /директор/i.test(p)) || (uid === 'ЦА' && p === 'Генеральный директор');
    u.groups.forEach(g => g.depts.forEach(d => d.pos.forEach(p => { if (isTop(g, d, p)) top.push([g.name, d.name, p]); })));
    out.push([uid === 'ЦА' ? 'ТОП менеджмент' : 'Административно-управленческий персонал', top]);
    u.groups.forEach(g => g.depts.forEach(d => {
      const rest = d.pos.filter(p => !isTop(g, d, p)).map(p => [g.name, d.name, p]);
      if (rest.length) out.push([dlabel(d.name), rest]);
    }));
    return out;
  }
  const src = (D, u, k) => D.shrsrc.filter(x => x.u === u && x.g === k[0] && x.d === k[1] && x.p === k[2]);
  const cnt = (D, u, k) => r2(src(D, u, k).reduce((a, x) => a + x.cnt, 0));
  const ppl = (D, u, k) => D.rows.filter(r => r.u === u && r.g === k[0] && r.d === k[1] && r.p === k[2] && r.note !== 'вакансия');
  function put(ws, r, c, v, o = {}) {
    const cell = ws.getCell(r, c); cell.value = v === '' ? null : v;
    cell.font = { name: TNR, size: o.size || 10, bold: !!o.bold, italic: !!o.italic };
    cell.alignment = { wrapText: true, vertical: 'middle', horizontal: o.center ? 'center' : undefined };
    if (o.border !== false) cell.border = BOX;
    if (o.fill) cell.fill = o.fill;
  }
  function textRows(ws, rows, r0, ncol, opts = {}) {
    rows.forEach((row, i) => row.forEach((v, j) => {
      if (v === '' || v == null || j >= ncol) return;
      const c = ws.getCell(r0 + i, j + 1); c.value = v; c.font = { name: TNR, size: opts.size || 10, bold: !!opts.bold };
      c.alignment = { wrapText: false, vertical: 'middle' };
    }));
  }

  // ---------------- ШТАТНОЕ РАСПИСАНИЕ ----------------
  const SHR = [['ЦА', 'ЦГ', 'Итого по ТОО "Центр гематологии" (центральный аппарат)'], ['Астана', 'Ф', 'Итого по медицинскому подразделению по г.Астана'],
    ['Лаборатория', 'ФЛ', 'Итого по лабораторному подразделению и СООЛЛ'], ['КРГ', 'ФК', 'Итого по филиалу "Центр гематологии города Караганды" ТОО "Центр гематологии"'],
    ['УК', 'ФУ', 'Итого по филиалу «Гематологическое отделение Восточно-Казахстанской области» ТОО «Центр гематологии»']];
  function shrSheet(wb, D, uid, pre, totalLabel) {
    const T = D.templates.shr[uid];
    const ws = wb.addWorksheet(T.sheet, { pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 } });
    [31, 11, 60, 22, 46].forEach((w, i) => ws.getColumn(i + 1).width = w);
    const head = T.head.map(r => r.slice());
    head[3][3] = '№ ______ от «___» __________ 2026 г.';
    textRows(ws, head.slice(0, 4), 1, 4);
    head.slice(5, 8).forEach((row, i) => { if (row[0]) { ws.mergeCells(6 + i, 1, 6 + i, 5); const c = ws.getCell(6 + i, 1); c.value = String(row[0]).trim(); c.font = { name: TNR, size: 11, bold: i === 0 }; c.alignment = { horizontal: 'center' }; } });
    const H = ['Категория персонала', 'Код должности', 'Должность (специальность, профессия), разряд, класс (категория) квалификации', 'Штатная численность (ед.)', 'Примечание'];
    H.forEach((h, i) => { put(ws, 9, i + 1, h, { bold: true, center: true }); put(ws, 10, i + 1, null); ws.mergeCells(9, i + 1, 10, i + 1); put(ws, 11, i + 1, i + 1, { center: true }); });
    let r = 12, grand = 0, si = 0;
    sections(D, uid).forEach(([title, keys]) => {
      si++;
      for (let c = 1; c <= 5; c++) put(ws, r, c, c === 1 ? title : null, { bold: true, fill: G });
      ws.mergeCells(r, 1, r, 5); r++;
      let sub = 0, pi = 0;
      keys.forEach(k => {
        const s = src(D, uid, k), n = cnt(D, uid, k), pp = ppl(D, uid, k); pi++;
        const cat = (s[0] && s[0].cat) || inferCat(k[2]);
        const olds = [...new Set(s.filter(x => norm(x.old) !== norm(k[2])).map(x => x.old))];
        let note = olds.length ? 'Было: ' + olds.join('; ') : '', fill;
        if (s.length && s.every(x => x.f === 'new')) note = 'Новая должность — вакансия';
        if (!s.length) { note = 'Новая должность — определить штатную численность' + (pp.length ? ` (по расстановке занято ${r2(pp.reduce((a, x) => a + x.r, 0))})` : ''); fill = Y; }
        put(ws, r, 1, cat, { fill }); put(ws, r, 2, `${pre}${String(si).padStart(2, '0')}-${String(pi).padStart(2, '0')}`, { center: true, fill });
        put(ws, r, 3, disp(k[2]), { fill }); put(ws, r, 4, n, { center: true, fill }); put(ws, r, 5, note, { size: 9, italic: !!note, fill });
        sub += n; r++;
      });
      [1, 2, 5].forEach(c => put(ws, r, c, null)); put(ws, r, 3, `Итого по разделу «${title}»`, { bold: true }); put(ws, r, 4, r2(sub), { bold: true, center: true });
      r++; grand += sub;
    });
    const dec = D.shrsrc.filter(x => x.u === uid && x.d === DEC);
    if (dec.length) {
      for (let c = 1; c <= 5; c++) put(ws, r, c, c === 1 ? 'Должности вне новой оргструктуры — требует решения' : null, { bold: true, fill: R });
      ws.mergeCells(r, 1, r, 5); r++;
      let sub = 0;
      dec.forEach(x => { put(ws, r, 1, x.cat || inferCat(x.old), { fill: R }); put(ws, r, 2, x.code, { center: true, fill: R }); put(ws, r, 3, x.old, { fill: R }); put(ws, r, 4, x.cnt, { center: true, fill: R }); put(ws, r, 5, x.c || 'Нет в новой оргструктуре', { size: 9, italic: true, fill: R }); sub += x.cnt; r++; });
      [1, 2, 5].forEach(c => put(ws, r, c, null)); put(ws, r, 3, 'Итого требует решения', { bold: true }); put(ws, r, 4, r2(sub), { bold: true, center: true }); r++; grand += sub;
    }
    [1, 2, 5].forEach(c => put(ws, r, c, null)); put(ws, r, 3, totalLabel, { bold: true }); put(ws, r, 4, r2(grand), { bold: true, center: true });
    r += 2;
    textRows(ws, T.sig, r, 4);
    return r2(grand);
  }

  // ---------------- ШТАТНАЯ РАССТАНОВКА ----------------
  function catgroup(cat, t) {
    const c = (cat || '').toLowerCase(); t = t.toLowerCase();
    if (c.includes('врач')) return 'Врачи';
    if (c.includes('средн') || c.includes('cредн')) return 'СМР';
    if (c.includes('младш')) return 'Младший мед.персонал';
    if (c.startsWith('руковод')) { if (/врач|ординатор|заведующ|клиническ|медицинск|лаборатор/.test(t)) return 'Врачи'; if (t.includes('сестрин')) return 'СМР'; }
    return 'Прочий персонал';
  }
  function rasBlock(ws, D, uid, r, C, nc) {
    const totals = {}, tadd = (k, d, e) => { totals[k] = totals[k] || [0, 0]; totals[k][0] += d; totals[k][1] += e; };
    sections(D, uid).forEach(([title, keys]) => {
      for (let c = 1; c <= nc; c++) put(ws, r, c, c === C.pos ? title : null, { bold: true, size: 11, fill: G });
      ws.mergeCells(r, C.pos, r, nc); r++;
      keys.forEach(k => {
        const s = src(D, uid, k), n = cnt(D, uid, k), pp = ppl(D, uid, k), occ = r2(pp.reduce((a, x) => a + x.r, 0));
        const list = pp.length ? pp : [null], r0 = r;
        const cat = (pp[0] && pp[0].cat) || (s[0] && s[0].cat) || inferCat(k[2]);
        list.forEach((p, i) => {
          let fill, notes = [];
          if (p) {
            if (p.o && norm(p.o) !== norm(p.p)) notes.push('Прежняя должность: ' + p.o);
            if (p.f === 'q') { notes.push('Уточнить: ' + (p.c || 'проверить назначение')); fill = Y; }
            if (p.mv) notes.push('Изменено в приложении');
            if (p.note) notes.push(p.note);
          } else { notes.push(s.length ? 'Вакансия' : 'Новая должность, в текущем ШР нет'); if (!s.length) fill = Y; }
          for (let c = 1; c <= nc; c++) put(ws, r, c, null, { size: 11, fill });
          if (i === 0) { put(ws, r, C.cat, cat, { size: 11, fill }); put(ws, r, C.pos, disp(k[2]), { size: 11, fill }); put(ws, r, C.sht, n, { size: 11, center: true, fill }); put(ws, r, C.vac, r2(n - occ), { size: 11, center: true, fill }); }
          put(ws, r, C.name, p ? (p.n || '') : '', { size: 11, fill }); put(ws, r, C.zan, p ? p.r : 0, { size: 11, center: true, fill });
          put(ws, r, C.note, notes.join('; '), { size: 10, italic: true, fill });
          r++;
        });
        if (r - r0 > 1) ['cat', 'pos', 'sht', 'vac'].forEach(kk => ws.mergeCells(r0, C[kk], r - 1, C[kk]));
        tadd(catgroup(cat, k[2]), n, occ);
      });
    });
    const dp = D.rows.filter(x => x.u === uid && x.d === DEC && x.note !== 'вакансия'), ds = D.shrsrc.filter(x => x.u === uid && x.d === DEC);
    if (dp.length || ds.length) {
      for (let c = 1; c <= nc; c++) put(ws, r, c, c === C.pos ? 'Должности вне новой оргструктуры — требует решения' : null, { bold: true, size: 11, fill: R });
      ws.mergeCells(r, C.pos, r, nc); r++;
      const by = {}; dp.forEach(x => (by[x.o || '—'] = by[x.o || '—'] || []).push(x));
      Object.entries(by).forEach(([o, ps]) => {
        const r0 = r, cat = ps[0].cat || inferCat(o);
        ps.forEach((p, i) => {
          for (let c = 1; c <= nc; c++) put(ws, r, c, null, { size: 11, fill: R });
          if (i === 0) { put(ws, r, C.cat, cat, { size: 11, fill: R }); put(ws, r, C.pos, o, { size: 11, fill: R }); }
          put(ws, r, C.name, p.n || '', { size: 11, fill: R }); put(ws, r, C.zan, p.r, { size: 11, center: true, fill: R });
          put(ws, r, C.note, [p.c, p.note].filter(Boolean).join('; '), { size: 10, italic: true, fill: R }); r++;
        });
        if (r - r0 > 1) ['cat', 'pos'].forEach(kk => ws.mergeCells(r0, C[kk], r - 1, C[kk]));
        tadd(catgroup(cat, o), 0, ps.reduce((a, x) => a + x.r, 0));
      });
      ds.forEach(x => tadd(catgroup(x.cat, x.old), x.cnt, 0));
    }
    return [r, totals];
  }
  function rasTotals(ws, r, C, totals, label) {
    r++;
    ['Врачи', 'СМР', 'Младший мед.персонал', 'Прочий персонал'].forEach((k, i) => {
      const [d, e] = totals[k] || [0, 0];
      put(ws, r, C.pos, i === 0 ? 'Всего' : null, { bold: true, size: 11 }); put(ws, r, C.name, k, { size: 11 });
      put(ws, r, C.sht, r2(d), { size: 11, center: true }); put(ws, r, C.zan, r2(e), { size: 11, center: true }); put(ws, r, C.vac, r2(d - e), { size: 11, center: true }); r++;
    });
    const D_ = r2(Object.values(totals).reduce((a, v) => a + v[0], 0)), E_ = r2(Object.values(totals).reduce((a, v) => a + v[1], 0));
    put(ws, r, C.pos, null); put(ws, r, C.name, label, { bold: true, size: 11 });
    put(ws, r, C.sht, D_, { bold: true, size: 11, center: true }); put(ws, r, C.zan, E_, { bold: true, size: 11, center: true }); put(ws, r, C.vac, r2(D_ - E_), { bold: true, size: 11, center: true });
    return r + 1;
  }
  function rasHead(ws, rows, r0, ncol) {
    rows.forEach((row, i) => row.forEach((v, j) => {
      if (j >= ncol) return;
      const isHdr = i === rows.length - 1;
      if (isHdr) put(ws, r0 + i, j + 1, v || null, { bold: true, size: 11, center: true });
      else if (v) { const c = ws.getCell(r0 + i, j + 1); c.value = v; c.font = { name: TNR, size: 12, bold: true }; }
    }));
  }
  const C7 = { cat: 1, pos: 2, name: 3, sht: 4, zan: 5, vac: 6, note: 7 }, C8 = { unit: 1, cat: 2, pos: 3, name: 4, sht: 5, zan: 6, vac: 7, note: 8 };
  const TITLE = 'Штатная расстановка по новой оргструктуре (проект), ноябрь 2026';

  window.buildShr = async function (D) {
    const wb = new ExcelJS.Workbook();
    SHR.forEach(([u, pre, lbl]) => shrSheet(wb, D, u, pre, lbl));
    return wb.xlsx.writeBuffer();
  };
  window.buildRas = async function (D) {
    const wb = new ExcelJS.Workbook();
    // Астана sheet: ЦА block + Астана block
    let ws = wb.addWorksheet('Астана', { pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 } });
    [27, 20, 37, 38, 14, 14, 12, 44].forEach((w, i) => ws.getColumn(i + 1).width = w);
    const h1 = D.templates.ras['Астана'].head.map(r => r.slice()); h1[0][0] = TITLE;
    rasHead(ws, h1, 1, 8);
    let [r, t] = rasBlock(ws, D, 'ЦА', 5, C8, 8);
    r = rasTotals(ws, r, C8, t, 'ИТОГО по АУП (центральный аппарат)');
    r += 3;
    const h2 = D.templates.ras['Астана2'].head.map(x => x.slice()); h2[0][0] = TITLE;
    rasHead(ws, h2, r, 8); r += 3;
    [r, t] = rasBlock(ws, D, 'Астана', r, C8, 8);
    rasTotals(ws, r, C8, t, 'ИТОГО по подразделению г. Астана');
    [['УК', 'УК', 'ИТОГО по филиалу г. Усть-Каменогорск'], ['КРГ', 'КРГ', 'ИТОГО по филиалу г. Караганды'], ['Лаб и ООЛЛ', 'Лаборатория', 'ИТОГО по Лаборатории и службе ООЛЛ']].forEach(([sh, uid, lbl]) => {
      const w = wb.addWorksheet(sh, { pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 } });
      [22, 36, 40, 12, 12, 12, 44].forEach((x, i) => w.getColumn(i + 1).width = x);
      const h = D.templates.ras[sh].head.map(x => x.slice()); h[0][0] = TITLE;
      rasHead(w, h, 1, 7);
      const [rr, tt] = rasBlock(w, D, uid, 5, C7, 7);
      rasTotals(w, rr, C7, tt, lbl);
    });
    return wb.xlsx.writeBuffer();
  };
})();
