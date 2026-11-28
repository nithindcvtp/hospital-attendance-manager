/* Shared logic used by the browser (instant UI) and the server (validation + Excel).
   Pure functions only: no I/O. */
(function (root) {
  'use strict';

  const SHIFTS = { D4: { start: '10:00', end: '18:30' }, M2: { start: '08:00', end: '16:30' } };
  const S = {
    NORMAL: 'Normal Duty', LEAVE: 'Leave', LEAVE_FH: 'Half Day Leave (First Half)', LEAVE_SH: 'Half Day Leave (Second Half)',
    WO: 'Weekly Off', HOL: 'Holiday', ADJ: 'Shift Adjustment', COVER: 'Covered Duty',
  };
  const STATUSES = [S.NORMAL, S.LEAVE, S.LEAVE_FH, S.LEAVE_SH, S.WO, S.HOL, S.ADJ, S.COVER];
  const LEAVES = [S.LEAVE, S.LEAVE_FH, S.LEAVE_SH];
  const WORKING = [S.NORMAL, S.ADJ, S.COVER];
  const NO_TIME = [S.LEAVE, S.WO, S.HOL];
  // Codes shown in the duty list (the hospital file only shows D4 / M2 / WO; leave codes are configurable here).
  const CODES = { WO: 'WO', HOL: 'H', LEAVE: 'L', LEAVE_FH: 'L-FH', LEAVE_SH: 'L-SH' };
  const COVER_TYPES = ['Full Duty Coverage', 'Shift Adjustment'];
  const AUTO = 'Auto: ';
  const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  const DOW = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  const SETTING_KEYS = ['HospitalName', 'Department', 'SupervisorName', 'TimeZone', 'AutoFrom'];

  const pad = (n) => String(n).padStart(2, '0');
  const daysIn = (y, m) => new Date(Date.UTC(y, m, 0)).getUTCDate();
  const isDate = (s) => typeof s === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(s) && new Date(s + 'T00:00:00Z').toISOString().slice(0, 10) === s;
  const isTime = (s) => typeof s === 'string' && /^([01]\d|2[0-3]):[0-5]\d$/.test(s);
  const clean = (s, max = 500) => String(s == null ? '' : s).trim().slice(0, max);
  const uid = (p) => p + Date.now().toString(36).toUpperCase() + Math.random().toString(36).slice(2, 6).toUpperCase();
  const fmt12 = (t) => { if (!t) return ''; const [h, m] = t.split(':').map(Number); return `${pad(h % 12 === 0 ? 12 : h % 12)}:${pad(m)} ${h >= 12 ? 'PM' : 'AM'}`; };
  const range12 = (a, b) => (a && b ? `${fmt12(a)} – ${fmt12(b)}` : '');
  const codeOf = (s, e) => Object.keys(SHIFTS).find((k) => SHIFTS[k].start === s && SHIFTS[k].end === e) || '';
  const toMin = (t) => Number(t.slice(0, 2)) * 60 + Number(t.slice(3));
  const toTime = (m) => `${pad(Math.floor(m / 60))}:${pad(m % 60)}`;
  const fail = (m) => { throw new Error(m); };
  const emptyOps = () => ({ att: [], delAtt: [], cov: [], delCov: [], shifts: [], employees: [], settings: {} });

  /* ---------- lookups ---------- */
  const findAtt = (db, date, id) => db.attendance.find((a) => a.Date === date && a.EmployeeID === id);
  function effectiveShift(db, id, y, m) {
    const r = db.shiftRows.find((x) => Number(x.Year) === y && Number(x.Month) === m && x.EmployeeID === id);
    if (r) return { start: r.StartTime, end: r.EndTime };
    const e = db.employees.find((x) => x.id === id);
    return e && e.fixed && SHIFTS[e.fixed] ? { start: SHIFTS[e.fixed].start, end: SHIFTS[e.fixed].end } : null;
  }
  function shiftMap(db, y, m) {
    const o = {};
    db.employees.forEach((e) => { const s = effectiveShift(db, e.id, y, m); if (s) o[e.id] = s; });
    return o;
  }
  const schedOn = (db, date, id) => {
    const r = findAtt(db, date, id);
    return r && r.ScheduledStart ? { start: r.ScheduledStart, end: r.ScheduledEnd } : effectiveShift(db, id, Number(date.slice(0, 4)), Number(date.slice(5, 7)));
  };

  /* ---------- ops (changes to apply locally and sync to the sheet) ---------- */
  function mergeOps(a, b) {
    const o = emptyOps();
    const up = (list, key, items, dels) => {
      const m = new Map(list.map((x) => [key(x), x]));
      items.forEach((x) => m.set(key(x), x));
      dels.forEach((id) => m.delete(id));
      return [...m.values()];
    };
    o.att = up(up(a.att, (x) => x.RecordID, [], b.delAtt), (x) => x.RecordID, b.att, []);
    o.delAtt = [...new Set([...a.delAtt.filter((id) => !b.att.some((x) => x.RecordID === id)), ...b.delAtt])];
    o.cov = up(up(a.cov, (x) => x.CoverageID, [], b.delCov), (x) => x.CoverageID, b.cov, []);
    o.delCov = [...new Set([...a.delCov.filter((id) => !b.cov.some((x) => x.CoverageID === id)), ...b.delCov])];
    const sk = (x) => `${x.Year}-${x.Month}-${x.EmployeeID}`;
    o.shifts = up(a.shifts, sk, b.shifts, []);
    o.employees = up(a.employees, (x) => x.id, b.employees, []);
    o.settings = { ...a.settings, ...b.settings };
    return o;
  }
  const isEmpty = (o) => !(o.att.length || o.delAtt.length || o.cov.length || o.delCov.length || o.shifts.length || o.employees.length || Object.keys(o.settings).length);

  function applyOps(db, ops) {
    const upsert = (list, key, items, dels) => {
      const m = new Map(list.map((x) => [key(x), x]));
      dels.forEach((id) => m.delete(id));
      items.forEach((x) => m.set(key(x), x));
      return [...m.values()];
    };
    const n = { ...db };
    // Keep one record per (date, employee) even if ids differ.
    const attKey = (x) => x.RecordID;
    let att = upsert(db.attendance, attKey, ops.att, ops.delAtt);
    const seen = new Map();
    att.forEach((x) => seen.set(x.Date + '|' + x.EmployeeID, x));
    n.attendance = [...seen.values()];
    n.coverage = upsert(db.coverage, (x) => x.CoverageID, ops.cov, ops.delCov);
    n.shiftRows = upsert(db.shiftRows, (x) => `${x.Year}-${x.Month}-${x.EmployeeID}`, ops.shifts, []);
    n.employees = upsert(db.employees, (x) => x.id, ops.employees, []);
    n.settings = { ...db.settings, ...ops.settings };
    return n;
  }

  /* ---------- validation ---------- */
  function checkTimes(a, b, what) {
    if ((a && !isTime(a)) || (b && !isTime(b))) fail(`${what} times must be valid (HH:MM).`);
    if ((a && !b) || (!a && b)) fail(`Enter both ${what.toLowerCase()} start and end time, or leave both empty.`);
  }

  /* ---------- attendance planning ---------- */
  function halfWorked(sched, status) {
    const a = toMin(sched.start), b = toMin(sched.end), mid = a + Math.round((b - a) / 2);
    return status === S.LEAVE_FH ? [toTime(mid), sched.end] : [sched.start, toTime(mid)];
  }

  // items: [{date, employeeId, status, shiftCode?, actualStart?, actualEnd?, remarks?}]
  function planAttendance(db, items) {
    const ops = emptyOps();
    const empIds = new Set(db.employees.map((e) => e.id));
    const byKey = new Map();
    items.forEach((it) => {
      if (!isDate(it.date)) fail('Invalid date.');
      if (!empIds.has(it.employeeId)) fail('Unknown employee.');
      if (!STATUSES.includes(it.status)) fail('Invalid attendance status.');
      if (it.shiftCode && !SHIFTS[it.shiftCode]) fail('Invalid shift.');
      byKey.set(it.date + '|' + it.employeeId, it);
    });
    byKey.forEach((it) => {
      const y = Number(it.date.slice(0, 4)), m = Number(it.date.slice(5, 7));
      const ex = findAtt(db, it.date, it.employeeId);
      let sched = it.shiftCode ? SHIFTS[it.shiftCode] : ex && ex.ScheduledStart ? { start: ex.ScheduledStart, end: ex.ScheduledEnd } : effectiveShift(db, it.employeeId, y, m);
      if (!sched) fail(`Shifts for ${MONTHS[m - 1]} ${y} are not configured yet. Configure them first.`);
      let a = clean(it.actualStart, 5), b = clean(it.actualEnd, 5);
      if (NO_TIME.includes(it.status)) { a = ''; b = ''; }
      checkTimes(a, b, 'Actual');
      if ((it.status === S.LEAVE_FH || it.status === S.LEAVE_SH) && !a) [a, b] = halfWorked(sched, it.status);
      ops.att.push({
        RecordID: ex ? ex.RecordID : uid('A'), Date: it.date, EmployeeID: it.employeeId,
        ScheduledStart: sched.start, ScheduledEnd: sched.end, ActualStart: a, ActualEnd: b, Status: it.status, Remarks: clean(it.remarks),
      });
    });
    return ops;
  }

  /* ---------- automatic replacement ----------
     When the 8:00-4:30 (M2) person is on leave (full day or first half), the 10:00-6:30 (D4) colleague
     takes over the 8:00 duty. If the D4 person is on leave nothing changes. */
  function revertCover(db, c) {
    const ops = emptyOps();
    ops.delCov.push(c.CoverageID);
    const r = findAtt(db, c.Date, c.CoveringEmployeeID);
    if (r && r.Status === S.COVER) ops.att.push({ ...r, Status: S.NORMAL, ActualStart: r.ScheduledStart, ActualEnd: r.ScheduledEnd });
    return ops;
  }

  function reconcile(db, date, absentId) {
    let ops = emptyOps();
    const notes = [];
    const y = Number(date.slice(0, 4)), m = Number(date.slice(5, 7));
    const rec = findAtt(db, date, absentId);
    const sched = schedOn(db, date, absentId);
    const trigger = !!(rec && (rec.Status === S.LEAVE || rec.Status === S.LEAVE_FH) && sched && codeOf(sched.start, sched.end) === 'M2');
    const autos = db.coverage.filter((c) => c.Date === date && c.AbsentEmployeeID === absentId && c.Remarks.startsWith(AUTO));
    let chosen = null;
    if (trigger) {
      const cands = db.employees.filter((e) => e.active && !e.fixed && e.id !== absentId).filter((e) => {
        const s = schedOn(db, date, e.id);
        const r = findAtt(db, date, e.id);
        return s && codeOf(s.start, s.end) === 'D4' && (!r || ![...LEAVES, S.WO, S.HOL].includes(r.Status));
      });
      chosen = cands.find((e) => autos.some((a) => a.CoveringEmployeeID === e.id)) || cands[0] || null;
    }
    autos.forEach((a) => { if (!chosen || a.CoveringEmployeeID !== chosen.id) ops = mergeOps(ops, revertCover(db, a)); });
    if (chosen) {
      const old = autos.find((a) => a.CoveringEmployeeID === chosen.id);
      const M2 = SHIFTS.M2;
      ops.cov.push({
        CoverageID: old ? old.CoverageID : uid('C'), Date: date, AbsentEmployeeID: absentId, CoveringEmployeeID: chosen.id, CoverageType: 'Shift Adjustment',
        ActualStart: M2.start, ActualEnd: M2.end, Remarks: AUTO + 'takes over the 8:00 AM duty',
      });
      const r = findAtt(db, date, chosen.id);
      const cs = (r && r.ScheduledStart) ? { start: r.ScheduledStart, end: r.ScheduledEnd } : effectiveShift(db, chosen.id, y, m);
      ops.att.push({
        RecordID: r ? r.RecordID : uid('A'), Date: date, EmployeeID: chosen.id, ScheduledStart: cs.start, ScheduledEnd: cs.end,
        ActualStart: M2.start, ActualEnd: M2.end, Status: S.COVER, Remarks: r ? r.Remarks : '',
      });
      notes.push(`${chosen.name} now covers the 8:00 AM duty on ${date.split('-').reverse().join('/')}`);
    }
    return { ops, notes };
  }

  // Run planned ops, then re-check replacements for the affected employee/dates.
  function finish(db, ops, pairs) {
    let total = ops, cur = applyOps(db, ops);
    const notes = [];
    const done = new Set();
    pairs.forEach(([date, id]) => {
      const k = date + '|' + id;
      if (done.has(k)) return; done.add(k);
      const r = reconcile(cur, date, id);
      total = mergeOps(total, r.ops); cur = applyOps(cur, r.ops); notes.push(...r.notes);
    });
    return { ops: total, notes };
  }

  function planDay(db, item) { return finish(db, planAttendance(db, [item]), [[item.date, item.employeeId]]); }

  function dateList(from, to) {
    if (!isDate(from)) fail('Invalid start date.');
    to = to || from;
    if (!isDate(to)) fail('Invalid end date.');
    if (to < from) fail('End date cannot be before start date.');
    const out = [];
    for (let d = new Date(from + 'T00:00:00Z'); d.toISOString().slice(0, 10) <= to; d.setUTCDate(d.getUTCDate() + 1)) {
      out.push(d.toISOString().slice(0, 10));
      if (out.length > 62) fail('Date range is too long (maximum 62 days).');
    }
    return out;
  }

  // kind: 'full' | 'first' | 'second'
  function planLeave(db, p) {
    const status = p.kind === 'first' ? S.LEAVE_FH : p.kind === 'second' ? S.LEAVE_SH : S.LEAVE;
    const dates = dateList(p.from, p.to);
    const ops = planAttendance(db, dates.map((date) => ({ date, employeeId: p.employeeId, status, remarks: p.remarks })));
    return finish(db, ops, dates.map((d) => [d, p.employeeId]));
  }

  function planDeleteAttendance(db, id) {
    const r = db.attendance.find((a) => a.RecordID === id);
    if (!r) fail('Record not found.');
    const ops = emptyOps(); ops.delAtt.push(id);
    return finish(db, ops, [[r.Date, r.EmployeeID]]);
  }

  function planWeeklyOffs(db, { year, month, weekday }) {
    const items = [];
    for (let d = 1; d <= daysIn(year, month); d++) {
      const date = `${year}-${pad(month)}-${pad(d)}`;
      if (new Date(date + 'T00:00:00Z').getUTCDay() !== weekday) continue;
      db.employees.filter((e) => e.active).forEach((e) => { if (!findAtt(db, date, e.id)) items.push({ date, employeeId: e.id, status: S.WO }); });
    }
    return { ops: items.length ? planAttendance(db, items) : emptyOps(), count: items.length };
  }

  /* ---------- automatic daily attendance ----------
     Every day, employees flagged "auto" are marked present on their monthly shift (Weekly Off on Sundays),
     unless a record already exists. Changes (leave etc.) are made by hand afterwards. */
  function planAutoPresent(db, today) {
    const ops = emptyOps();
    if (!isDate(today)) return ops;
    const hasFrom = isDate(db.settings.AutoFrom);
    const from = hasFrom ? db.settings.AutoFrom : today;
    if (!hasFrom) ops.settings.AutoFrom = today;
    const autos = db.employees.filter((e) => e.active && e.auto);
    if (!autos.length || from > today) return ops;
    const items = [];
    let n = 0;
    for (let d = new Date(from + 'T00:00:00Z'); d.toISOString().slice(0, 10) <= today && n < 400; d.setUTCDate(d.getUTCDate() + 1), n++) {
      const date = d.toISOString().slice(0, 10);
      autos.forEach((e) => {
        if (findAtt(db, date, e.id)) return;
        const s = effectiveShift(db, e.id, Number(date.slice(0, 4)), Number(date.slice(5, 7)));
        if (!s) return;                                   // month not configured yet: nothing to mark
        if (d.getUTCDay() === 0) items.push({ date, employeeId: e.id, status: S.WO });
        else items.push({ date, employeeId: e.id, status: S.NORMAL, actualStart: s.start, actualEnd: s.end });
      });
    }
    if (items.length) ops.att = planAttendance(db, items).att;
    return ops;
  }

  /* ---------- manual coverage ---------- */
  function planCoverage(db, p) {
    const emps = db.employees;
    const name = (id) => (emps.find((e) => e.id === id) || {}).name;
    if (!name(p.absentId)) fail('Select the employee on leave.');
    if (!name(p.coveringId)) fail('Select the covering employee.');
    if (p.absentId === p.coveringId) fail('The covering employee must be different from the employee on leave.');
    if (!COVER_TYPES.includes(p.type)) fail('Invalid coverage type.');
    const a = clean(p.actualStart, 5), b = clean(p.actualEnd, 5);
    checkTimes(a, b, 'Actual duty');
    const editing = p.id ? db.coverage.find((c) => c.CoverageID === p.id) : null;
    if (p.id && !editing) fail('Coverage record not found.');
    const dates = editing ? [p.from] : dateList(p.from, p.to);
    if (editing && !isDate(p.from)) fail('Invalid date.');
    const ops = emptyOps(), items = [];
    dates.forEach((date) => {
      const clash = db.coverage.find((c) => c.Date === date && c.AbsentEmployeeID === p.absentId && c.CoverageID !== (editing && editing.CoverageID));
      const target = editing || clash;
      ops.cov.push({ CoverageID: target ? target.CoverageID : uid('C'), Date: date, AbsentEmployeeID: p.absentId, CoveringEmployeeID: p.coveringId, CoverageType: p.type, ActualStart: a, ActualEnd: b, Remarks: clean(p.remarks) });
      if (p.updateAttendance !== false) {
        const exA = findAtt(db, date, p.absentId), exC = findAtt(db, date, p.coveringId);
        if (exC && LEAVES.includes(exC.Status) && exC.Status === S.LEAVE) fail(`${name(p.coveringId)} is on leave on ${date}, so cannot cover.`);
        if (!exA || !LEAVES.includes(exA.Status)) items.push({ date, employeeId: p.absentId, status: S.LEAVE, remarks: exA ? exA.Remarks : '' });
        items.push({ date, employeeId: p.coveringId, status: S.COVER, actualStart: a, actualEnd: b, remarks: exC ? exC.Remarks : '' });
      }
    });
    const attOps = items.length ? planAttendance(db, items) : emptyOps();
    // Manual coverage replaces any automatic one for the same day.
    return mergeOps(attOps, ops);
  }
  const planDeleteCoverage = (db, id) => { const o = emptyOps(); if (!db.coverage.some((c) => c.CoverageID === id)) fail('Record not found.'); o.delCov.push(id); return o; };

  /* ---------- shifts / employees ---------- */
  function planShifts(db, { year, month, codes }) {
    const ops = emptyOps();
    db.employees.filter((e) => e.active).forEach((e) => {
      const code = e.fixed || codes[e.id];
      if (!SHIFTS[code]) fail(`Please choose a shift for ${e.name}.`);
      ops.shifts.push({ Month: month, Year: year, EmployeeID: e.id, StartTime: SHIFTS[code].start, EndTime: SHIFTS[code].end });
    });
    return ops;
  }

  function planEmployee(db, p) {
    const name = clean(p.name, 60);
    if (!name) fail('Employee name is required.');
    if (db.employees.some((e) => e.name.toLowerCase() === name.toLowerCase() && e.id !== p.id)) fail('Another employee already has this name.');
    let id = p.id;
    if (!id) id = 'E' + String(db.employees.reduce((m, e) => Math.max(m, Number(e.id.replace(/\D/g, '')) || 0), 0) + 1).padStart(3, '0');
    const ops = emptyOps();
    const ex = db.employees.find((e) => e.id === id) || {};
    const keep = (v, cur, dflt) => (v === undefined ? (cur === undefined ? dflt : cur) : v);
    const fixed = keep(p.fixedShift, ex.fixed, '');
    ops.employees.push({ id, name, active: p.active !== false, code: clean(keep(p.code, ex.code, ''), 20), fixed: SHIFTS[fixed] ? fixed : '', auto: !!keep(p.auto, ex.auto, false) });
    return ops;
  }

  /* ---------- server-side validation of incoming ops ---------- */
  function validateOps(db, p) {
    const ops = emptyOps(), errors = [];
    const empIds = new Set(db.employees.map((e) => e.id));
    const tryIt = (fn) => { try { fn(); } catch (e) { errors.push(e.message); } };
    (p.att || []).forEach((r) => tryIt(() => {
      if (!clean(r.RecordID, 40)) fail('Missing record id.');
      if (!isDate(r.Date)) fail('Invalid date.');
      if (!empIds.has(r.EmployeeID)) fail('Unknown employee.');
      if (!STATUSES.includes(r.Status)) fail('Invalid attendance status.');
      if (!isTime(r.ScheduledStart) || !isTime(r.ScheduledEnd)) fail('Invalid scheduled time.');
      checkTimes(clean(r.ActualStart, 5), clean(r.ActualEnd, 5), 'Actual');
      ops.att.push({ RecordID: clean(r.RecordID, 40), Date: r.Date, EmployeeID: r.EmployeeID, ScheduledStart: r.ScheduledStart, ScheduledEnd: r.ScheduledEnd, ActualStart: clean(r.ActualStart, 5), ActualEnd: clean(r.ActualEnd, 5), Status: r.Status, Remarks: clean(r.Remarks) });
    }));
    (p.cov || []).forEach((r) => tryIt(() => {
      if (!clean(r.CoverageID, 40)) fail('Missing coverage id.');
      if (!isDate(r.Date)) fail('Invalid date.');
      if (!empIds.has(r.AbsentEmployeeID) || !empIds.has(r.CoveringEmployeeID) || r.AbsentEmployeeID === r.CoveringEmployeeID) fail('Invalid coverage employees.');
      if (!COVER_TYPES.includes(r.CoverageType)) fail('Invalid coverage type.');
      checkTimes(clean(r.ActualStart, 5), clean(r.ActualEnd, 5), 'Actual duty');
      ops.cov.push({ CoverageID: clean(r.CoverageID, 40), Date: r.Date, AbsentEmployeeID: r.AbsentEmployeeID, CoveringEmployeeID: r.CoveringEmployeeID, CoverageType: r.CoverageType, ActualStart: clean(r.ActualStart, 5), ActualEnd: clean(r.ActualEnd, 5), Remarks: clean(r.Remarks) });
    }));
    (p.shifts || []).forEach((r) => tryIt(() => {
      const y = Number(r.Year), m = Number(r.Month);
      if (!Number.isInteger(y) || y < 2000 || y > 2100 || !Number.isInteger(m) || m < 1 || m > 12) fail('Invalid shift month.');
      if (!empIds.has(r.EmployeeID) || !codeOf(r.StartTime, r.EndTime)) fail('Invalid shift.');
      ops.shifts.push({ Month: m, Year: y, EmployeeID: r.EmployeeID, StartTime: r.StartTime, EndTime: r.EndTime });
    }));
    (p.employees || []).forEach((r) => tryIt(() => {
      if (!/^E\d{3,}$/.test(r.id || '')) fail('Invalid employee id.');
      const name = clean(r.name, 60);
      if (!name) fail('Employee name is required.');
      if (db.employees.some((e) => e.name.toLowerCase() === name.toLowerCase() && e.id !== r.id)) fail('Another employee already has this name.');
      ops.employees.push({ id: r.id, name, active: r.active !== false, code: clean(r.code, 20), fixed: SHIFTS[r.fixed] ? r.fixed : '', auto: !!r.auto });
    }));
    tryIt(() => {
      const s = p.settings || {};
      SETTING_KEYS.forEach((k) => { if (k in s) ops.settings[k] = clean(s[k], 120); });
      if (ops.settings.AutoFrom && !isDate(ops.settings.AutoFrom)) delete ops.settings.AutoFrom;
      if (ops.settings.TimeZone) { try { new Intl.DateTimeFormat('en', { timeZone: ops.settings.TimeZone }); } catch (e) { delete ops.settings.TimeZone; fail('Invalid time zone.'); } }
    });
    ops.delAtt = (p.delAtt || []).map((x) => clean(x, 40)).filter(Boolean);
    ops.delCov = (p.delCov || []).map((x) => clean(x, 40)).filter(Boolean);
    return { ops, errors };
  }

  /* ---------- report ---------- */
  function cellInfo(rec, defCode) {
    if (!rec) return { text: '', changed: false };
    switch (rec.Status) {
      case S.WO: return { text: CODES.WO, changed: false };
      case S.HOL: return { text: CODES.HOL, changed: false };
      case S.LEAVE: return { text: CODES.LEAVE, changed: true };
      case S.LEAVE_FH: return { text: CODES.LEAVE_FH, changed: true };
      case S.LEAVE_SH: return { text: CODES.LEAVE_SH, changed: true };
      default: break;
    }
    const sched = codeOf(rec.ScheduledStart, rec.ScheduledEnd);
    if (rec.Status === S.NORMAL) return { text: sched || 'DUTY', changed: !!defCode && sched !== defCode };
    return { text: codeOf(rec.ActualStart, rec.ActualEnd) || (sched || 'DUTY') + '*', changed: true };
  }

  function buildReport({ db, year, month, employeeId, today }) {
    const employees = db.employees;
    const names = Object.fromEntries(employees.map((e) => [e.id, e.name]));
    const shifts = shiftMap(db, year, month);
    const ym = `${year}-${pad(month)}`;
    const att = db.attendance.filter((a) => a.Date.startsWith(ym));
    const cov = db.coverage.filter((c) => c.Date.startsWith(ym));
    const withRecords = new Set(att.map((a) => a.EmployeeID));
    let emps = employees.filter((e) => e.active || withRecords.has(e.id));
    if (employeeId && employeeId !== 'ALL') emps = emps.filter((e) => e.id === employeeId);
    const attMap = new Map(att.map((a) => [a.Date + '|' + a.EmployeeID, a]));
    const total = daysIn(year, month);

    const rows = [];
    for (let d = 1; d <= total; d++) {
      const date = `${ym}-${pad(d)}`;
      for (const e of emps) {
        const rec = attMap.get(date + '|' + e.id);
        const sh = shifts[e.id];
        const status = rec ? rec.Status : 'Not Recorded';
        const ss = rec ? rec.ScheduledStart : sh ? sh.start : '';
        const se = rec ? rec.ScheduledEnd : sh ? sh.end : '';
        let covered = '';
        if (status === S.COVER) {
          covered = cov.filter((c) => c.Date === date && c.CoveringEmployeeID === e.id).map((c) => names[c.AbsentEmployeeID] || c.AbsentEmployeeID).join(', ');
        } else if (LEAVES.includes(status)) {
          const by = cov.filter((c) => c.Date === date && c.AbsentEmployeeID === e.id).map((c) => names[c.CoveringEmployeeID] || c.CoveringEmployeeID);
          if (by.length) covered = `Covered by ${by.join(', ')}`;
        }
        const incomplete = status === 'Not Recorded' ? date <= today : WORKING.includes(status) && !(rec.ActualStart && rec.ActualEnd);
        rows.push({
          date, employeeId: e.id, employee: e.name, scheduledStart: ss, scheduledEnd: se, scheduled: range12(ss, se),
          actualStart: rec ? rec.ActualStart : '', actualEnd: rec ? rec.ActualEnd : '', actual: rec ? range12(rec.ActualStart, rec.ActualEnd) : '',
          status, covered, remarks: rec ? rec.Remarks : '', incomplete, recordId: rec ? rec.RecordID : '',
        });
      }
    }

    const roster = emps.map((e) => {
      const def = shifts[e.id] ? codeOf(shifts[e.id].start, shifts[e.id].end) : '';
      return {
        employeeId: e.id, code: e.code || '', name: e.name.toUpperCase(),
        cells: Array.from({ length: total }, (_, i) => {
          const d = i + 1;
          return { day: d, dow: DOW[new Date(Date.UTC(year, month - 1, d)).getUTCDay()], ...cellInfo(attMap.get(`${ym}-${pad(d)}|${e.id}`), def) };
        }),
      };
    }).sort((a, b) => (Number(b.code) || 0) - (Number(a.code) || 0));

    const summaries = emps.map((e) => {
      const mine = rows.filter((r) => r.employeeId === e.id);
      const count = (s) => mine.filter((r) => r.status === s).length;
      const half = count(S.LEAVE_FH) + count(S.LEAVE_SH);
      const weeklyOffs = count(S.WO), holidays = count(S.HOL);
      return {
        employeeId: e.id, employee: e.name, code: e.code || '', scheduledDays: total - weeklyOffs - holidays,
        attendanceDays: count(S.NORMAL) + count(S.ADJ) + count(S.COVER) + half * 0.5, leaveDays: count(S.LEAVE) + half * 0.5, weeklyOffs, holidays,
        coverageDays: cov.filter((c) => c.CoveringEmployeeID === e.id).length, shiftAdjustments: count(S.ADJ),
        notRecorded: mine.filter((r) => r.status === 'Not Recorded' && r.incomplete).length,
        incompleteRecords: mine.filter((r) => r.incomplete && r.status !== 'Not Recorded').length,
      };
    });

    const coverageRows = cov
      .filter((c) => !employeeId || employeeId === 'ALL' || c.AbsentEmployeeID === employeeId || c.CoveringEmployeeID === employeeId)
      .sort((a, b) => a.Date.localeCompare(b.Date))
      .map((c) => ({ id: c.CoverageID, date: c.Date, absent: names[c.AbsentEmployeeID] || c.AbsentEmployeeID, covering: names[c.CoveringEmployeeID] || c.CoveringEmployeeID, type: c.CoverageType, timing: range12(c.ActualStart, c.ActualEnd), remarks: c.Remarks }));

    const schedule = emps.map((e) => ({
      employeeId: e.id, employee: e.name, code: e.code || '', start: shifts[e.id] ? shifts[e.id].start : '', end: shifts[e.id] ? shifts[e.id].end : '',
      shift: shifts[e.id] ? codeOf(shifts[e.id].start, shifts[e.id].end) : '',
    }));

    let workingDays = 0;
    for (let d = 1; d <= total; d++) {
      const date = `${ym}-${pad(d)}`;
      const day = rows.filter((r) => r.date === date);
      if (!day.length || !day.every((r) => r.status === S.WO || r.status === S.HOL)) workingDays++;
    }
    const stats = {
      label: `${MONTHS[month - 1]} ${year}`, workingDays, coverageDays: cov.length,
      adjustmentDays: rows.filter((r) => r.status === S.ADJ).length, incomplete: rows.filter((r) => r.incomplete).length,
    };
    return { year, month, label: stats.label, rows, roster, summaries, coverage: coverageRows, schedule, stats };
  }

  const api = {
    SHIFTS, S, STATUSES, LEAVES, WORKING, CODES, COVER_TYPES, AUTO, MONTHS, DOW, SETTING_KEYS,
    pad, daysIn, isDate, isTime, clean, uid, fmt12, range12, codeOf, emptyOps, isEmpty, mergeOps, applyOps,
    findAtt, effectiveShift, shiftMap, planAttendance, planDay, planLeave, planDeleteAttendance, planWeeklyOffs,
    planCoverage, planDeleteCoverage, planShifts, planEmployee, planAutoPresent, validateOps, buildReport,
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.Core = api;
})(typeof window !== 'undefined' ? window : this);
