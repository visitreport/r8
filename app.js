/* =========================================================================
   รายงานเยี่ยมร้านค้าเทียบแผน — app.js
   ========================================================================= */

const sb = supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

let planFile = null;
let actualFile = null;
let lastComparison = null;
let lastBatchMeta = null;

/* ---------------------- helpers: parsing values ---------------------- */

function toNum(v) {
  if (v === null || v === undefined || v === '') return null;
  const n = typeof v === 'number' ? v : parseFloat(String(v).replace(/,/g, ''));
  return isNaN(n) ? null : n;
}

function toInt(v) {
  const n = toNum(v);
  return n === null ? null : Math.round(n);
}

function toStr(v) {
  if (v === null || v === undefined) return null;
  const s = String(v).trim();
  return s === '' ? null : s;
}

// รับค่าเป็น string "dd/mm/yyyy", Date object, หรือ excel serial number -> คืน "yyyy-mm-dd"
function excelValueToISODate(v) {
  if (v === null || v === undefined || v === '') return null;
  if (v instanceof Date) {
    return v.toISOString().slice(0, 10);
  }
  if (typeof v === 'number') {
    const d = XLSX.SSF.parse_date_code(v);
    if (!d) return null;
    return `${d.y}-${String(d.m).padStart(2, '0')}-${String(d.d).padStart(2, '0')}`;
  }
  const s = String(v).trim();
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) {
    const [, d, mo, y] = m;
    return `${y}-${mo.padStart(2, '0')}-${d.padStart(2, '0')}`;
  }
  return null;
}

// รับค่าเป็น string "dd/mm/yyyy hh:mm:ss", Date object, หรือ excel serial number -> คืน ISO datetime + เขตเวลา +07:00
function excelValueToISODateTime(v) {
  if (v === null || v === undefined || v === '') return null;
  if (v instanceof Date) return v.toISOString();
  if (typeof v === 'number') {
    const d = XLSX.SSF.parse_date_code(v);
    if (!d) return null;
    const pad = (n) => String(Math.floor(n)).padStart(2, '0');
    return `${d.y}-${pad(d.m)}-${pad(d.d)}T${pad(d.H || 0)}:${pad(d.M || 0)}:${pad(d.S || 0)}+07:00`;
  }
  const s = String(v).trim();
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2})(?::(\d{2}))?/);
  if (m) {
    const [, d, mo, y, hh, mm, ss] = m;
    return `${y}-${mo.padStart(2, '0')}-${d.padStart(2, '0')}T${hh.padStart(2, '0')}:${mm}:${(ss || '00').padStart(2, '0')}+07:00`;
  }
  return null;
}

function fmtDt(iso) {
  if (!iso) return '';
  return String(iso).replace('T', ' ').replace(/\+07:00$/, '').slice(0, 19);
}

function escapeHtml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/* ---------------------- helpers: sheet header detection ---------------------- */

function findHeaderRowIndex(matrix, mustInclude) {
  for (let i = 0; i < Math.min(matrix.length, 20); i++) {
    const row = matrix[i] || [];
    const rowSet = new Set(row.map((c) => (c === null || c === undefined ? '' : String(c).trim())));
    if (mustInclude.every((h) => rowSet.has(h))) return i;
  }
  return 0;
}

// ตรวจสอบว่าไฟล์ที่อัปโหลดมาตรงกับชนิดไฟล์ที่คาดหวังไหม (กันกรณีอัปโหลดไฟล์สลับช่องกัน)
// requiredCols: คอลัมน์ที่ต้องมีถ้าเป็นไฟล์ชนิดนี้จริง
// conflictCols: คอลัมน์เฉพาะของไฟล์ "อีกชนิด" ที่ไม่ควรเจอในไฟล์นี้
function detectHeaderAndValidate(matrix, requiredCols, conflictCols, fileLabel, otherLabel) {
  for (let i = 0; i < Math.min(matrix.length, 20); i++) {
    const rowSet = new Set((matrix[i] || []).map((c) => (c === null || c === undefined ? '' : String(c).trim())));
    if (requiredCols.every((h) => rowSet.has(h))) return i;
  }
  for (let i = 0; i < Math.min(matrix.length, 20); i++) {
    const rowSet = new Set((matrix[i] || []).map((c) => (c === null || c === undefined ? '' : String(c).trim())));
    if (conflictCols.every((h) => rowSet.has(h))) {
      throw new Error(`ไฟล์นี้ดูเหมือนจะเป็นไฟล์ "${otherLabel}" ไม่ใช่ไฟล์ "${fileLabel}" — กรุณาตรวจสอบว่าอัปโหลดไฟล์สลับช่องกันหรือไม่`);
    }
  }
  throw new Error(`ไม่พบคอลัมน์ที่จำเป็นสำหรับไฟล์ "${fileLabel}" (ต้องมีคอลัมน์: ${requiredCols.join(', ')}) กรุณาตรวจสอบว่าเลือกไฟล์ถูกต้องหรือไม่`);
}

/* ---------------------- parse: ไฟล์แผนเยี่ยม ---------------------- */

function parsePlanWorkbook(wb) {
  const sheetName = wb.SheetNames.includes('CallPlanData') ? 'CallPlanData' : wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const matrix = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
  const headerIdx = detectHeaderAndValidate(
    matrix,
    ['รหัสร้านค้า', 'วันที่เข้าเยี่ยม', 'ลำดับการเข้าเยี่ยม'],
    ['ในแผน/นอกแผน', 'เวลาเข้า'],
    'แผนเยี่ยม',
    'เยี่ยมร้านค้าจริง'
  );
  const rows = XLSX.utils.sheet_to_json(ws, { range: headerIdx, defval: null, raw: true });

  return rows
    .map((r) => ({
      company: toStr(r['บริษัท']),
      channel: toStr(r['ช่องทาง']),
      visit_order: toInt(r['ลำดับการเข้าเยี่ยม']),
      box_name: toStr(r['ชื่อกล่อง']),
      store_code: toStr(r['รหัสร้านค้า']),
      store_name: toStr(r['ชื่อร้านค้า']),
      region: toStr(r['ภาค']),
      sales_unit: toStr(r['หน่วยขาย']),
      store_type: toStr(r['ประเภทร้านค้า']),
      employee_code: toStr(r['รหัสพนักงาน']),
      employee_firstname: toStr(r['ชื่อพนักงาน']),
      employee_lastname: toStr(r['นามสกุลพนักงาน']),
      plan_visit_date: excelValueToISODate(r['วันที่เข้าเยี่ยม']),
      latitude: toNum(r['Latitude']),
      longitude: toNum(r['Longitude']),
    }))
    .filter((r) => r.store_code && r.plan_visit_date);
}

/* ---------------------- parse: ไฟล์เยี่ยมร้านค้าจริง ---------------------- */

function parseActualWorkbook(wb) {
  const sheetName = wb.SheetNames.includes('Sheet1') ? 'Sheet1' : wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const matrix = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
  const headerIdx = detectHeaderAndValidate(
    matrix,
    ['รหัสร้านค้า', 'วันที่เข้าเยี่ยม', 'ในแผน/นอกแผน'],
    ['ลำดับการเข้าเยี่ยม', 'หน่วยขาย'],
    'เยี่ยมร้านค้าจริง',
    'แผนเยี่ยม'
  );
  const rows = XLSX.utils.sheet_to_json(ws, { range: headerIdx, defval: null, raw: true });

  return rows
    .map((r) => ({
      seq_no: toInt(r['ลำดับที่']),
      employee_code: toStr(r['รหัสพนักงาน']),
      employee_name: toStr(r['ชื่อพนักงาน']),
      box_name: toStr(r['ชื่อกล่อง']),
      store_code: toStr(r['รหัสร้านค้า']),
      store_name: toStr(r['ชื่อร้านค้า']),
      lat_in: toNum(r['ละติจูดเข้า']),
      lng_in: toNum(r['ลองจิจูดเข้า']),
      distance_in_m: toNum(r['ระยะห่างเข้า (เมตร)']),
      time_in: excelValueToISODateTime(r['เวลาเข้า']),
      reason_in: toStr(r['เหตุผลการเข้านอกพิกัด']),
      lat_out: toNum(r['ละติจูดออก']),
      lng_out: toNum(r['ลองจิจูดออก']),
      distance_out_m: toNum(r['ระยะห่างออก (เมตร)']),
      time_out: excelValueToISODateTime(r['เวลาออก']),
      reason_out: toStr(r['เหตุผลการออกนอกพิกัด']),
      duration_minutes: toInt(r['ใช้เวลา']),
      note: toStr(r['Note']),
      source_plan_status: toStr(r['ในแผน/นอกแผน']),
      visit_date: excelValueToISODate(r['วันที่เข้าเยี่ยม']),
      photo_url: toStr(r['รูปถ่าย Check-In']),
      closed_status: toStr(r['สถานะร้านปิด/เลิกกิจการ']),
      closed_note: toStr(r['หมายเหตุสถานะร้านปิด/เลิกกิจการ']),
      closed_photo_url: toStr(r['รูปสถานะร้านปิด/เลิกกิจการ']),
      lat_record: toNum(r['ละติจูด(บันทึกพิกัด)']),
      lng_record: toNum(r['ลองจิจูด(บันทึกพิกัด)']),
      distance_record_m: toNum(r['ระยะห่างบันทึกพิกัด(เมตร)']),
      qty_whiskey_white: toNum(r['สุราขาว']),
      qty_whiskey_color: toNum(r['สุราสี']),
      qty_rtd: toNum(r['RTD']),
      qty_beer: toNum(r['เบียร์']),
      qty_oishi: toNum(r['โออิชิ']),
      qty_est: toNum(r['เอส']),
      qty_soda: toNum(r['โซดา']),
      qty_water: toNum(r['น้ำ']),
      qty_other: toNum(r['อื่นๆ']),
      qty_total: toNum(r['total']),
    }))
    .filter((r) => r.store_code && r.visit_date);
}

/* ---------------------- เปรียบเทียบแผน vs เยี่ยมจริง ---------------------- */
// จับคู่ด้วย รหัสร้านค้า + เดือนเดียวกัน (ไม่ต้องตรงวันที่เป๊ะ)
// จัดกลุ่มเป็นระดับ "ร้านค้า x เดือน" ก่อน แล้วดูว่ากลุ่มนั้นมีการเยี่ยมจริงหรือไม่

function monthOf(isoDate) {
  return isoDate ? isoDate.slice(0, 7) : null; // 'YYYY-MM'
}

function uniqueNames(list) {
  const set = new Set(list.filter(Boolean).map((s) => s.trim()).filter(Boolean));
  return Array.from(set).join(', ');
}

function computeComparison(planRows, actualRows) {
  const planGroups = new Map();
  planRows.forEach((p) => {
    const month = monthOf(p.plan_visit_date);
    const key = p.store_code + '||' + month;
    if (!planGroups.has(key)) {
      planGroups.set(key, {
        store_code: p.store_code,
        store_name: p.store_name,
        region: p.region,
        sales_unit: p.sales_unit,
        store_type: p.store_type,
        month,
        plans: [],
      });
    }
    planGroups.get(key).plans.push(p);
  });

  const actualGroups = new Map();
  actualRows.forEach((a) => {
    const month = monthOf(a.visit_date);
    const key = a.store_code + '||' + month;
    if (!actualGroups.has(key)) {
      actualGroups.set(key, {
        store_code: a.store_code,
        store_name: a.store_name,
        month,
        visits: [],
      });
    }
    actualGroups.get(key).visits.push(a);
  });

  const visited = [];
  const notVisited = [];
  planGroups.forEach((g, key) => {
    const ag = actualGroups.get(key);
    if (ag && ag.visits.length) {
      visited.push({ planGroup: g, actualGroup: ag });
    } else {
      notVisited.push(g);
    }
  });

  const offPlan = [];
  actualGroups.forEach((ag, key) => {
    if (!planGroups.has(key)) offPlan.push(ag);
  });

  return { visited, notVisited, offPlan };
}

/* ---------------------- UI: log ---------------------- */

function logReset() {
  const el = document.getElementById('log');
  el.textContent = '';
  el.classList.remove('error');
  el.style.display = '';
}
function logAppend(msg) {
  const el = document.getElementById('log');
  el.textContent += (el.textContent ? '\n' : '') + msg;
}
function logError(msg) {
  const el = document.getElementById('log');
  el.classList.add('error');
  el.textContent += (el.textContent ? '\n' : '') + '❌ ' + msg;
}

/* ---------------------- Supabase I/O ---------------------- */

async function checkConnection() {
  const el = document.getElementById('connStatus');
  if (!el) return;
  const textEl = el.querySelector('span');
  try {
    const { error } = await sb.from('upload_batches').select('id', { count: 'exact', head: true });
    if (error) throw error;
    textEl.textContent = 'เชื่อมต่อ Supabase สำเร็จ';
    el.className = 'connection ok';
  } catch (e) {
    textEl.textContent = 'เชื่อมต่อไม่สำเร็จ - ตรวจสอบ config.js (' + (e.message || '') + ')';
    el.className = 'connection warn';
  }
}

async function insertChunked(table, rows, batchId, chunkSize = 500, concurrency = 4) {
  const withBatch = rows.map((r) => ({ ...r, batch_id: batchId }));
  const chunks = [];
  for (let i = 0; i < withBatch.length; i += chunkSize) {
    chunks.push(withBatch.slice(i, i + chunkSize));
  }
  let nextIdx = 0;

  async function worker() {
    while (nextIdx < chunks.length) {
      const myIdx = nextIdx++;
      const { error } = await sb.from(table).insert(chunks[myIdx]);
      if (error) throw error;
    }
  }

  const workerCount = Math.min(concurrency, chunks.length) || 1;
  const workers = Array.from({ length: workerCount }, () => worker());
  await Promise.all(workers);
}

/* ---------------------- file inputs ---------------------- */

document.getElementById('file-plan').addEventListener('change', (e) => {
  planFile = e.target.files[0] || null;
  document.getElementById('fn-plan').textContent = planFile ? planFile.name : '';
  document.getElementById('dz-plan').classList.toggle('filled', !!planFile);
  updateProcessBtn();
});

document.getElementById('file-actual').addEventListener('change', (e) => {
  actualFile = e.target.files[0] || null;
  document.getElementById('fn-actual').textContent = actualFile ? actualFile.name : '';
  document.getElementById('dz-actual').classList.toggle('filled', !!actualFile);
  updateProcessBtn();
});

function updateProcessBtn() {
  document.getElementById('btn-process').disabled = !(planFile && actualFile);
}

/* ---------------------- process button ---------------------- */

document.getElementById('btn-process').addEventListener('click', async () => {
  const btn = document.getElementById('btn-process');
  const btnOriginalText = btn.textContent;
  btn.disabled = true;
  btn.textContent = 'กำลังประมวลผล...';
  logReset();
  try {
    logAppend('กำลังอ่านไฟล์แผนเยี่ยม...');
    let planWb;
    try {
      planWb = XLSX.read(await planFile.arrayBuffer(), { type: 'array', cellDates: false });
    } catch (e) {
      throw new Error('ไม่สามารถเปิดไฟล์แผนเยี่ยมได้ ตรวจสอบว่าเป็นไฟล์ Excel (.xlsx/.xls) ที่ไม่เสียหาย');
    }
    const planRows = parsePlanWorkbook(planWb);
    logAppend(`อ่านแผนเยี่ยมสำเร็จ: ${planRows.length} แถว`);

    logAppend('กำลังอ่านไฟล์เยี่ยมร้านค้าจริง...');
    let actualWb;
    try {
      actualWb = XLSX.read(await actualFile.arrayBuffer(), { type: 'array', cellDates: false });
    } catch (e) {
      throw new Error('ไม่สามารถเปิดไฟล์เยี่ยมร้านค้าจริงได้ ตรวจสอบว่าเป็นไฟล์ Excel (.xlsx/.xls) ที่ไม่เสียหาย');
    }
    const actualRows = parseActualWorkbook(actualWb);
    logAppend(`อ่านเยี่ยมจริงสำเร็จ: ${actualRows.length} แถว`);

    if (!planRows.length) throw new Error('ไม่พบข้อมูลในไฟล์แผนเยี่ยม (ตรวจสอบคอลัมน์ รหัสร้านค้า / วันที่เข้าเยี่ยม)');
    if (!actualRows.length) throw new Error('ไม่พบข้อมูลในไฟล์เยี่ยมร้านค้าจริง (ตรวจสอบคอลัมน์ รหัสร้านค้า / วันที่เข้าเยี่ยม)');

    // เช็คว่าร้านค้าในไฟล์ทั้ง 2 มีตรงกันบ้างไหม — ถ้าไม่มีเลย มักแปลว่าเป็นไฟล์คนละบริษัท/สาขา/ช่วงเวลากัน
    const planStoreSet = new Set(planRows.map((r) => r.store_code));
    const overlapCount = actualRows.reduce((n, r) => n + (planStoreSet.has(r.store_code) ? 1 : 0), 0);
    if (overlapCount === 0) {
      throw new Error(
        'ไม่พบรหัสร้านค้าที่ตรงกันเลยระหว่างไฟล์แผนเยี่ยมกับไฟล์เยี่ยมร้านค้าจริง — ไฟล์ทั้ง 2 อาจไม่ใช่คู่เดียวกัน (คนละบริษัท/สาขา/ช่วงเวลา) กรุณาตรวจสอบไฟล์ที่เลือกอีกครั้ง'
      );
    }

    const allDates = [...planRows.map((r) => r.plan_visit_date), ...actualRows.map((r) => r.visit_date)]
      .filter(Boolean)
      .sort();

    const batchMetaInsert = {
      company: planRows[0]?.company || null,
      channel: planRows[0]?.channel || null,
      region: planRows[0]?.region || null,
      period_start: allDates[0] || null,
      period_end: allDates[allDates.length - 1] || null,
      plan_filename: planFile.name,
      actual_filename: actualFile.name,
      plan_row_count: planRows.length,
      actual_row_count: actualRows.length,
    };

    logAppend('กำลังสร้างชุดข้อมูล (batch) ใน Supabase...');
    const { data: batchData, error: batchErr } = await sb.from('upload_batches').insert(batchMetaInsert).select().single();
    if (batchErr) throw batchErr;
    const batchId = batchData.id;

    logAppend('กำลังบันทึกแผนเยี่ยมและเยี่ยมจริงลง Supabase (พร้อมกัน)...');
    await Promise.all([
      insertChunked('visit_plans', planRows, batchId),
      insertChunked('actual_visits', actualRows, batchId),
    ]);

    logAppend('กำลังคำนวณเปรียบเทียบ...');
    const cmp = computeComparison(planRows, actualRows);
    lastComparison = cmp;
    lastBatchMeta = batchData;
    renderResults(batchData, cmp);

    logAppend('เสร็จสิ้น ✅ บันทึกและสรุปผลเรียบร้อย');
    document.getElementById('log').style.display = 'none';
  } catch (err) {
    console.error(err);
    logError(err.message || String(err));
  } finally {
    btn.disabled = !(planFile && actualFile);
    btn.textContent = btnOriginalText;
  }
});

/* ---------------------- results rendering ---------------------- */

let currentTabKind = 'visited';

function renderResults(batchMeta, cmp) {
  document.getElementById('results-section').style.display = '';
  document.getElementById('table-card').style.display = '';
  document.getElementById('stat-total').textContent = cmp.visited.length + cmp.notVisited.length;
  document.getElementById('stat-visited').textContent = cmp.visited.length;
  document.getElementById('stat-notvisited').textContent = cmp.notVisited.length;
  document.getElementById('stat-offplan').textContent = cmp.offPlan.length;
  document.getElementById('period-hint').textContent = `${batchMeta.period_start || ''} ถึง ${batchMeta.period_end || ''}`;
  document.querySelectorAll('.filter-pill').forEach((x) => x.classList.remove('active'));
  document.querySelector('.filter-pill[data-tab="visited"]').classList.add('active');
  document.getElementById('search-input').value = '';
  currentTabKind = 'visited';
  renderTable('visited');
}

function renderTable(kind) {
  currentTabKind = kind;
  const container = document.getElementById('table-container');
  if (!lastComparison) {
    container.innerHTML = '<div class="empty-note">ยังไม่มีข้อมูล</div>';
    return;
  }
  let rows, headers, mapper;
  if (kind === 'visited') {
    rows = lastComparison.visited;
    headers = ['เดือน', 'รหัสร้านค้า', 'ชื่อร้านค้า', 'ภาค/หน่วยขาย', 'พนักงานตามแผน', 'จำนวนครั้งตามแผน', 'จำนวนครั้งที่ไปจริง', 'วันที่ไปครั้งแรก', 'วันที่ไปล่าสุด'];
    mapper = ({ planGroup: g, actualGroup: ag }) => [
      g.month,
      g.store_code,
      g.store_name,
      `${g.region || ''}/${g.sales_unit || ''}`,
      uniqueNames(g.plans.map((p) => `${p.employee_firstname || ''} ${p.employee_lastname || ''}`)),
      g.plans.length,
      ag.visits.length,
      ag.visits.map((v) => v.visit_date).sort()[0] || '',
      ag.visits.map((v) => v.visit_date).sort().slice(-1)[0] || '',
    ];
  } else if (kind === 'notvisited') {
    rows = lastComparison.notVisited;
    headers = ['เดือน', 'รหัสร้านค้า', 'ชื่อร้านค้า', 'ภาค/หน่วยขาย', 'พนักงานตามแผน', 'จำนวนครั้งตามแผน'];
    mapper = (g) => [
      g.month,
      g.store_code,
      g.store_name,
      `${g.region || ''}/${g.sales_unit || ''}`,
      uniqueNames(g.plans.map((p) => `${p.employee_firstname || ''} ${p.employee_lastname || ''}`)),
      g.plans.length,
    ];
  } else {
    rows = lastComparison.offPlan;
    headers = ['เดือน', 'รหัสร้านค้า', 'ชื่อร้านค้า', 'พนักงานที่เยี่ยม', 'จำนวนครั้งที่ไปจริง', 'วันที่ไปครั้งแรก', 'วันที่ไปล่าสุด'];
    mapper = (ag) => [
      ag.month,
      ag.store_code,
      ag.store_name,
      uniqueNames(ag.visits.map((v) => v.employee_name)),
      ag.visits.length,
      ag.visits.map((v) => v.visit_date).sort()[0] || '',
      ag.visits.map((v) => v.visit_date).sort().slice(-1)[0] || '',
    ];
  }

  if (!rows.length) {
    container.innerHTML = '<div class="empty-note">ไม่มีข้อมูลในหมวดนี้</div>';
    return;
  }

  const query = (document.getElementById('search-input').value || '').trim().toLowerCase();
  const mappedAll = rows.map((r) => ({ r, cells: mapper(r) }));
  const filtered = query ? mappedAll.filter(({ cells }) => cells.some((c) => String(c ?? '').toLowerCase().includes(query))) : mappedAll;

  if (!filtered.length) {
    container.innerHTML = '<div class="empty-note">ไม่พบข้อมูลที่ตรงกับคำค้นหา</div>';
    return;
  }

  const maxRows = 500;
  const shown = filtered.slice(0, maxRows);
  let html = '<table class="data"><thead><tr>' + headers.map((h) => `<th>${h}</th>`).join('') + '</tr></thead><tbody>';
  shown.forEach(({ cells }) => {
    html += '<tr>' + cells.map((c) => `<td>${escapeHtml(c ?? '')}</td>`).join('') + '</tr>';
  });
  html += '</tbody></table>';
  if (filtered.length > maxRows) {
    html += `<div class="empty-note">แสดง ${maxRows} จาก ${filtered.length} แถว — ดาวน์โหลด Excel เพื่อดูข้อมูลทั้งหมด</div>`;
  }
  container.innerHTML = html;
}

document.querySelectorAll('.filter-pill').forEach((t) => {
  t.addEventListener('click', () => {
    document.querySelectorAll('.filter-pill').forEach((x) => x.classList.remove('active'));
    t.classList.add('active');
    renderTable(t.dataset.tab);
  });
});

document.getElementById('search-input').addEventListener('input', () => {
  renderTable(currentTabKind);
});


/* ---------------------- download report ---------------------- */

document.getElementById('btn-download').addEventListener('click', () => {
  if (!lastComparison || !lastBatchMeta) return;
  downloadReport(lastBatchMeta, lastComparison);
});

function downloadReport(batchMeta, cmp) {
  const wb = XLSX.utils.book_new();

  const summaryRows = [
    { A: 'บริษัท', B: batchMeta.company || '' },
    { A: 'ช่องทาง', B: batchMeta.channel || '' },
    { A: 'ช่วงเวลา', B: `${batchMeta.period_start || ''} ถึง ${batchMeta.period_end || ''}` },
    { A: 'ไฟล์แผนเยี่ยม', B: batchMeta.plan_filename || '' },
    { A: 'ไฟล์เยี่ยมจริง', B: batchMeta.actual_filename || '' },
    { A: '', B: '' },
    { A: 'จำนวนร้านในแผนทั้งหมด', B: cmp.visited.length + cmp.notVisited.length },
    { A: 'ร้านที่เยี่ยมแล้ว', B: cmp.visited.length },
    { A: 'ร้านที่ยังไม่เยี่ยม', B: cmp.notVisited.length },
    { A: 'ร้านที่เยี่ยมนอกแผน', B: cmp.offPlan.length },
  ];
  const wsSummary = XLSX.utils.json_to_sheet(summaryRows, { skipHeader: true });
  XLSX.utils.book_append_sheet(wb, wsSummary, 'สรุป');

  const visitedRows = cmp.visited.map(({ planGroup: g, actualGroup: ag }) => {
    const dates = ag.visits.map((v) => v.visit_date).sort();
    return {
      เดือน: g.month,
      รหัสร้านค้า: g.store_code,
      ชื่อร้านค้า: g.store_name,
      ภาค: g.region,
      หน่วยขาย: g.sales_unit,
      ประเภทร้านค้า: g.store_type,
      พนักงานตามแผน: uniqueNames(g.plans.map((p) => `${p.employee_firstname || ''} ${p.employee_lastname || ''}`)),
      จำนวนครั้งตามแผน: g.plans.length,
      พนักงานที่เยี่ยมจริง: uniqueNames(ag.visits.map((v) => v.employee_name)),
      จำนวนครั้งที่ไปจริง: ag.visits.length,
      วันที่ไปครั้งแรก: dates[0] || '',
      วันที่ไปล่าสุด: dates[dates.length - 1] || '',
    };
  });
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(visitedRows), 'ร้านที่เยี่ยมแล้ว');

  const notVisitedRows = cmp.notVisited.map((g) => ({
    เดือน: g.month,
    รหัสร้านค้า: g.store_code,
    ชื่อร้านค้า: g.store_name,
    ภาค: g.region,
    หน่วยขาย: g.sales_unit,
    ประเภทร้านค้า: g.store_type,
    พนักงานตามแผน: uniqueNames(g.plans.map((p) => `${p.employee_firstname || ''} ${p.employee_lastname || ''}`)),
    จำนวนครั้งตามแผน: g.plans.length,
  }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(notVisitedRows), 'ร้านที่ยังไม่เยี่ยม');

  const offPlanRows = cmp.offPlan.map((ag) => {
    const dates = ag.visits.map((v) => v.visit_date).sort();
    return {
      เดือน: ag.month,
      รหัสร้านค้า: ag.store_code,
      ชื่อร้านค้า: ag.store_name,
      พนักงานที่เยี่ยม: uniqueNames(ag.visits.map((v) => v.employee_name)),
      จำนวนครั้งที่ไปจริง: ag.visits.length,
      วันที่ไปครั้งแรก: dates[0] || '',
      วันที่ไปล่าสุด: dates[dates.length - 1] || '',
    };
  });
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(offPlanRows), 'ร้านที่เยี่ยมนอกแผน');

  const fname = `รายงานเยี่ยมร้านค้าเทียบแผน_${batchMeta.period_start || ''}_${batchMeta.period_end || ''}.xlsx`.replace(/\s+/g, '');
  XLSX.writeFile(wb, fname);
}

/* ---------------------- init ---------------------- */

checkConnection();
