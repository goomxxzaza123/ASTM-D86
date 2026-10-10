// โหลด/ส่งออกข้อมูลฟอร์มเป็นไฟล์ Excel ใช้ร่วมกันทั้ง index.html และ AD86LDO.html
// ไฟล์ที่ส่งออกจัดหน้าเหมือนรายงาน และโหลดกลับเข้าฟอร์มได้

const EXCEL_HEADER_KEYS = [
  { id: 'testDate',   re: /วันที่|date/i },
  { id: 'testTime',   re: /เวลา|time/i },
  { id: 'sampleType', re: /ประเภท|sample/i },
  { id: 'density',    re: /density|ถพ|ความหนาแน่น/i },
  { id: 'totalVol',   re: /total|ปริมาตรที่ใช้/i },
  { id: 'moisture',   re: /moisture|ความชื้น/i },
];
const EXCEL_NUM_KEYS = ['20','40','60','80','100','120','140','160','180'];
const EXCEL_ROW_KEYS = ['ibp', ...EXCEL_NUM_KEYS, 'dsl', 'fol', 'kero'];

function excelRowKey(cell) {
  let s = String(cell ?? '').trim().toLowerCase();
  if (s === 'kerosene') s = 'kero';
  if (/^\d+(\.0+)?$/.test(s)) s = String(parseFloat(s));
  return EXCEL_ROW_KEYS.includes(s) ? s : null;
}

function excelSetInput(id, val) {
  const el = document.getElementById(id);
  if (!el || el.tagName !== 'INPUT' || val === undefined || val === null || String(val).trim() === '') return false;
  el.value = String(val).replace(/^>\s*/, '').trim();
  el.dispatchEvent(new Event('input', { bubbles: true }));
  return true;
}

function importExcelFile(file) {
  if (typeof XLSX === 'undefined') { showToast('โหลดไลบรารี Excel ไม่สำเร็จ (ต้องต่อเน็ต)'); return; }
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb = XLSX.read(e.target.result, { type: 'array' });
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: '', raw: true });
      const isLabel = v => EXCEL_HEADER_KEYS.some(h => h.re.test(String(v)));
      const headers = [], dist = [], seen = new Set();
      rows.forEach(r => {
        // แถวข้อมูลกลั่น: ชื่อจุดอยู่คอลัมน์ A หรือ (แถว 20–180 แบบรายงาน) อยู่คอลัมน์ C
        let key = excelRowKey(r[0]);
        if (!key && EXCEL_NUM_KEYS.includes(excelRowKey(r[2]))) key = excelRowKey(r[2]);
        if (key) { if (!dist.some(d => d.key === key)) dist.push({ key, r }); return; }
        // แถวส่วนหัว: ค่าคือช่องถัดไปที่ไม่ว่าง
        r.forEach((c, i) => {
          const h = EXCEL_HEADER_KEYS.find(h => h.re.test(String(c)));
          if (!h || seen.has(h.id)) return;
          const v = r.slice(i + 1).find(x => String(x).trim() !== '');
          if (v === undefined || isLabel(v)) return;
          seen.add(h.id);
          headers.push({ id: h.id, val: v });
        });
      });
      // ตั้งค่า totalVol ก่อน เพื่อให้ % คำนวณถูกต้อง
      headers.sort((a, b) => (b.id === 'totalVol') - (a.id === 'totalVol'));
      let count = 0;
      headers.forEach(h => { if (excelSetInput(h.id, h.val)) count++; });
      dist.forEach(({ key, r }) => {
        if (excelSetInput('b_' + key, r[1])) count++;
        if (excelSetInput('o_' + key, r[2])) count++;
        if (excelSetInput('v_' + key, r[3])) count++;
      });
      if (typeof calcFolVol === 'function') calcFolVol();
      if (typeof calcKeroVol === 'function') calcKeroVol();
      showToast(count ? '✓ โหลดไฟล์ Excel แล้ว (' + count + ' ช่อง)' : 'ไม่พบข้อมูลที่ตรงรูปแบบ');
    } catch (err) {
      console.error(err);
      showToast('อ่านไฟล์ Excel ไม่ได้');
    }
  };
  reader.readAsArrayBuffer(file);
}

// ---------- ส่งออก ----------
const XS_BORDER = { style: 'thin', color: { rgb: '9AB8D8' } };
const XS_ALL = { top: XS_BORDER, bottom: XS_BORDER, left: XS_BORDER, right: XS_BORDER };
function xsStyle(o = {}) {
  return {
    font: { name: 'Tahoma', sz: o.sz || 11, bold: !!o.bold, color: { rgb: o.color || '111111' } },
    alignment: { horizontal: o.align || 'center', vertical: 'center', wrapText: true },
    fill: o.fill ? { patternType: 'solid', fgColor: { rgb: o.fill } } : undefined,
    border: o.border === false ? undefined : (o.border || XS_ALL),
  };
}

function excelFormVal(id) {
  const el = document.getElementById(id);
  if (!el) return '';
  const s = (el.tagName === 'INPUT' ? el.value : el.textContent).trim();
  if (s === '' || s === '—') return '';
  return /^-?\d+(\.\d+)?$/.test(s) ? parseFloat(s) : s;
}

function downloadExcelTemplate() {
  if (typeof XLSX === 'undefined') { showToast('โหลดไลบรารี Excel ไม่สำเร็จ (ต้องต่อเน็ต)'); return; }
  const ws = {}, merges = [];
  let maxRow = 0;
  const put = (r, c, v, s, z) => {
    const cell = { v: v ?? '', t: typeof v === 'number' ? 'n' : 's', s };
    if (z) cell.z = z;
    ws[XLSX.utils.encode_cell({ r, c })] = cell;
    maxRow = Math.max(maxRow, r);
  };
  const merge = (r1, c1, r2, c2, s) => {
    merges.push({ s: { r: r1, c: c1 }, e: { r: r2, c: c2 } });
    for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++)
      if (!ws[XLSX.utils.encode_cell({ r, c })]) put(r, c, '', s);
  };
  const plain = (o = {}) => xsStyle({ border: false, ...o });
  const underline = (o = {}) => xsStyle({ bold: true, border: { bottom: { style: 'thin', color: { rgb: '333333' } } }, ...o });

  // หัวรายงาน
  const title = (document.querySelector('.report-title') || {}).textContent || 'รายงาน ASTM D86';
  put(0, 0, title, plain({ bold: true, sz: 15, color: '1A2A4A' })); merge(0, 0, 0, 4, plain());
  put(1, 0, 'วันที่ตรวจสอบ :', plain({ align: 'right' }));
  put(1, 1, excelFormVal('testDate'), underline());
  put(1, 2, 'เวลา :', plain({ align: 'right' }));
  put(1, 3, excelFormVal('testTime'), underline());
  put(2, 0, 'ประเภทวัตถุดิบ :', plain({ align: 'right' }));
  put(2, 1, excelFormVal('sampleType'), underline({ align: 'left' })); merge(2, 1, 2, 4, underline());

  // ตาราง meta
  const lbl = xsStyle({ bold: true, align: 'left', fill: 'EEF4FB', color: '1A2A4A' });
  [['ค่า ถพ. (Density)', 'density', 'kg/L.', '0.000'],
   ['ปริมาตรที่ใช้ทดสอบ (Total Volume)', 'totalVol', 'ml.', '0'],
   ['% ความชื้น (% moisture)', 'moisture', '%', '0.00']].forEach(([t, id, unit, z], i) => {
    const r = 4 + i;
    put(r, 0, t, lbl); merge(r, 0, r, 2, lbl);
    put(r, 3, excelFormVal(id), xsStyle({ bold: true, sz: 12 }), z);
    put(r, 4, unit, xsStyle({ sz: 10, color: '555555' }));
  });

  // ตารางกลั่น
  put(8, 0, 'ข้อมูลการกลั่น (Distillation Data)', underline({ align: 'left', color: '1A2A4A', border: { bottom: { style: 'medium', color: { rgb: '2A4A80' } } } }));
  merge(8, 0, 8, 4, underline({ border: { bottom: { style: 'medium', color: { rgb: '2A4A80' } } } }));
  const th = xsStyle({ bold: true, fill: 'DDEAF7', color: '1A2A4A', sz: 10 });
  put(9, 0, '', th); merge(9, 0, 10, 0, th);
  put(9, 1, 'อุณหภูมิก้นหอ\n(Bottom temp) °C', th); merge(9, 1, 10, 1, th);
  put(9, 2, 'อุณหภูมิปลายหอ\n(OVHD temp) °C', th); merge(9, 2, 10, 2, th);
  put(9, 3, 'ปริมาตร', th); merge(9, 3, 9, 4, th);
  put(10, 3, 'ml.', th); put(10, 4, '%', th);

  const fills = { ibp: 'C8DFF5', dsl: 'D8EEC8', fol: 'F5E8C8', kero: 'F5E8C8' };
  const names = { ibp: 'IBP', dsl: 'DSL', fol: 'FOL', kero: 'Kerosene' };
  const groupLabel = ((document.querySelector('#distBody td.col-label[rowspan]') || {}).textContent || 'LDO').trim();
  const val = xsStyle({ bold: true });
  const green = xsStyle({ bold: true, color: '2E7D52' });
  const blue = xsStyle({ bold: true, color: '1A6BB5' });
  const ovhd = xsStyle({ color: '1A2A4A', fill: 'F5F8FC' });
  let r = 11;

  const special = k => {
    const gt = k === 'fol' || k === 'kero';
    const vol = excelFormVal('v_' + k) !== '' ? excelFormVal('v_' + k) : excelFormVal('v_' + k + '_display');
    const tone = k === 'ibp' ? blue : green;
    put(r, 0, names[k], xsStyle({ bold: true, fill: fills[k], color: '1A2A4A' }));
    put(r, 1, excelFormVal('b_' + k), val, gt ? '"> "General' : undefined);
    put(r, 2, excelFormVal('o_' + k), val, gt ? '"> "General' : undefined);
    put(r, 3, vol, tone, gt ? '0.0' : undefined);
    put(r, 4, excelFormVal('p_' + k), tone, k === 'ibp' ? undefined : '0.0');
    r++;
  };

  special('ibp');
  const g0 = r;
  EXCEL_NUM_KEYS.forEach(k => {
    put(r, 1, excelFormVal('b_' + k), val);
    put(r, 2, Number(k), ovhd);
    put(r, 3, excelFormVal('v_' + k), green);
    put(r, 4, excelFormVal('p_' + k), green, '0.0');
    r++;
  });
  const gl = xsStyle({ bold: true, fill: 'EEF4FB', color: '1A2A4A' });
  put(g0, 0, groupLabel, gl); merge(g0, 0, r - 1, 0, gl);
  ['dsl', 'fol', 'kero'].forEach(k => { if (document.getElementById('b_' + k)) special(k); });

  // สรุปผล (เฉพาะหน้า LDO)
  if (document.getElementById('text_lf')) {
    r++;
    put(r, 0, 'ผล Light fraction และ Kerosene', underline({ align: 'left', color: '1A2A4A' })); merge(r, 0, r, 4, underline());
    r++;
    [['Light fraction', 'text_lf', 'pct_lf'], ['Kerosene', 'text_kero', 'pct_kero']].forEach(([t, tid, pid]) => {
      put(r, 0, t, lbl);
      put(r, 1, excelFormVal(tid), xsStyle({ align: 'left' })); merge(r, 1, r, 3, xsStyle());
      put(r, 4, excelFormVal(pid), green);
      r++;
    });
  }

  ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: maxRow, c: 4 } });
  ws['!merges'] = merges;
  ws['!cols'] = [{ wch: 16 }, { wch: 20 }, { wch: 20 }, { wch: 12 }, { wch: 10 }];
  ws['!rows'] = Array.from({ length: maxRow + 1 }, (_, i) => ({ hpt: i === 0 ? 28 : (i === 9 || i === 10 ? 22 : 20) }));

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'D86');
  const name = excelFormVal('sampleType') || 'report';
  XLSX.writeFile(wb, 'ASTM-D86-' + name + '.xlsx');
  showToast('✓ ส่งออก Excel แล้ว');
}

document.addEventListener('DOMContentLoaded', () => {
  const inp = document.getElementById('excelFile');
  if (inp) inp.addEventListener('change', () => {
    if (inp.files[0]) importExcelFile(inp.files[0]);
    inp.value = '';
  });
});
