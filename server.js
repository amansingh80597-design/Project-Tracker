// Retail Engineering Project Tracker - backend (zero framework, JSON file store)
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { fileURLToPath } from 'node:url';
import nodemailer from 'nodemailer';
import * as XLSX from 'xlsx';
import { STAGES } from './lib.js';
import { seedData } from './seed.js';
import { parseWorkbook, matchOutlet } from './importer.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PORT = process.env.PORT || 3000;
const DB_FILE = process.env.DB_FILE || path.join(__dirname, 'data', 'db.json');
const PUBLIC = path.join(__dirname, 'public');

/* ---------- store ---------- */
let db;
function load() {
  fs.mkdirSync(path.dirname(DB_FILE), { recursive: true });
  if (fs.existsSync(DB_FILE)) db = JSON.parse(fs.readFileSync(DB_FILE, 'utf8'));
  else { db = seedData(hashPw); save(); }
  db.assets ||= []; db.scrap ||= []; db.settings.salesAreas ||= {};
  ensureDefaults();
  for (const u of db.users) if (u.mustChange === undefined) u.mustChange = ['admin123', 'vendor123'].some((p) => checkPw(p, u));
  db.recovered ||= []; db.autoPipe ||= [];
  if (autoScrap(db.assets, 'system', 'replaced earlier') + backfillDuParts()) save();
}
function save() {
  const tmp = DB_FILE + '.tmp';
  fs.writeFileSync(tmp, JSON.stringify(db, null, 1));
  fs.renameSync(tmp, DB_FILE);
}
const uid = (p) => p + '_' + crypto.randomBytes(4).toString('hex');

/* ---------- auth ---------- */
function hashPw(pw, salt = crypto.randomBytes(8).toString('hex')) {
  return { salt, hash: crypto.scryptSync(pw, salt, 32).toString('hex') };
}
function checkPw(pw, u) {
  const h = crypto.scryptSync(pw, u.salt, 32);
  return crypto.timingSafeEqual(h, Buffer.from(u.hash, 'hex'));
}
function sign(payload) {
  const body = Buffer.from(JSON.stringify(payload)).toString('base64url');
  const sig = crypto.createHmac('sha256', db.secret).update(body).digest('base64url');
  return body + '.' + sig;
}
function verify(token) {
  if (!token) return null;
  const [body, sig] = token.split('.');
  const good = crypto.createHmac('sha256', db.secret).update(body || '').digest('base64url');
  if (!sig || sig.length !== good.length || !crypto.timingSafeEqual(Buffer.from(sig), Buffer.from(good))) return null;
  const p = JSON.parse(Buffer.from(body, 'base64url').toString());
  return p.exp > Date.now() ? p : null;
}

/* ---------- date helpers ---------- */
const iso = (d) => { const z = new Date(d.getTime() - d.getTimezoneOffset() * 6e4); return z.toISOString().slice(0, 10); };
const parse = (s) => new Date(s + 'T00:00:00');
const addDays = (s, n) => { const d = parse(s); d.setDate(d.getDate() + n); return iso(d); };
const isWork = (s) => !(db?.settings?.weekOff || [0]).includes(parse(s).getDay());
const nextWork = (s) => { while (!isWork(s)) s = addDays(s, 1); return s; };
const addWork = (s, n) => { s = nextWork(s); while (n > 0) { s = nextWork(addDays(s, 1)); n--; } return s; };
const today = () => iso(new Date());

/* ---------- workflow helpers ---------- */
const stageList = (job) => STAGES[job.type].filter((s) => s.key !== 'cabling' || job.cabling === 'Required');
const stageIdx = (job, key = job.stage) => stageList(job).findIndex((s) => s.key === key);
const vendorOf = (job) => job.vendorId || db.teams.find((t) => t.id === job.teamId)?.vendorId || null;
const soOf = (o) => { const a = db.settings.salesAreas?.[o.salesArea] || {}; return { name: o.soName || a.name || '', email: o.soEmail || a.email || '', phone: o.soPhone || a.phone || '' }; };
const outletOf = (job) => db.outlets.find((o) => o.code === job.outletCode);
const teamOf = (job) => db.teams.find((t) => t.id === job.teamId);
const needs = (job) => {
  const n = [];
  if (job.newQty > 0 && job.newModel) n.push({ kind: 'DU', model: job.newModel, qty: job.newQty });
  if (job.stpQty > 0 && job.stpModel) n.push({ kind: 'STP', model: job.stpModel, qty: job.stpQty });
  return n;
};
const usable = (r, kind, model, job) => r.kind === kind && r.model === model && (job && job.source === 'Reuse' ? r.cond === 'Reusable' : (r.cond === 'New' || r.cond === 'Refurbished'));
const avail = (r) => r.qty - r.reserved;
function log(job, text, by) { job.history.push({ at: new Date().toISOString(), by, text }); }
function move(type, row, qty, ref, by) {
  db.movements.push({ id: uid('mv'), at: new Date().toISOString(), type, kind: row.kind, model: row.model, cond: row.cond, store: row.store, qty, ref, by });
}
function duration(job) {
  if (job.daysOverride > 0) return job.daysOverride;
  const o = outletOf(job) || {}; const dr = o.driveway === 'RCC' ? 'RCC' : 'Paver';
  const d = db.settings.durations[job.type]?.[dr] || 2;
  const qty = job.type === 'STP_REPLACEMENT' ? job.stpQty : job.type === 'RATIONALIZATION' ? job.oldQty : job.newQty;
  return Math.max(1, Math.ceil(d * Math.max(1, qty || 1) / 2)) + (job.cabling === 'Required' ? db.settings.cablingDays || 0 : 0);
}

function unitsMove(r, from, to, n, job) {
  const out = []; if (!r || !r.units?.length) return out;
  for (const u of r.units) {
    if (out.length >= n) break; if (u.status !== from || (from === 'Reserved' && u.job !== job.ref)) continue;
    u.status = to; if (to === 'In stock') delete u.job; else u.job = job.ref;
    if (to === 'Installed') { u.outletCode = job.outletCode; u.installedOn = today(); }
    out.push(u);
  }
  return out;
}
function registerInstalled(r, units, job) {
  for (const u of units) {
    if (db.assets.some((a) => a.kind === r.kind && a.serial === u.serial && a.status !== 'Removed')) continue;
    db.assets.push({ id: uid('a'), kind: r.kind, serial: u.serial, sap: job.outletCode, make: r.make || '', model: r.model, category: '', pumpType: r.kind === 'DU' ? 'Pressure' : '', nozzles: r.nozzles || null, hp: r.hp || '', year: new Date().getFullYear(), contract: '', source: 'TRACKER', status: 'Active', installedBy: job.ref, installedOn: today() });
  }
}
function reserve(job, by) {
  job.reserved = []; job.shortage = [];
  for (const n of needs(job)) {
    let left = n.qty;
    for (const r of db.stock.filter((r) => usable(r, n.kind, n.model, job))) {
      const take = Math.min(left, Math.max(0, avail(r)));
      if (take > 0) { r.reserved += take; left -= take; job.reserved.push({ stockId: r.id, qty: take }); unitsMove(r, 'In stock', 'Reserved', take, job); move('RESERVE', r, take, job.ref, by); }
    }
    if (left > 0) job.shortage.push({ kind: n.kind, model: n.model, qty: left });
  }
}
function release(job, by) {
  for (const x of job.reserved || []) {
    const r = db.stock.find((r) => r.id === x.stockId);
    if (r) { r.reserved = Math.max(0, r.reserved - x.qty); unitsMove(r, 'Reserved', 'In stock', x.qty, job); move('RELEASE', r, x.qty, job.ref, by); }
  }
  job.reserved = []; job.shortage = [];
}
function consume(job, by) {
  for (const x of job.reserved || []) {
    const r = db.stock.find((r) => r.id === x.stockId);
    if (r) { r.qty -= x.qty; r.reserved = Math.max(0, r.reserved - x.qty); registerInstalled(r, unitsMove(r, 'Reserved', 'Installed', x.qty, job), job); move('INSTALL', r, x.qty, job.ref, by); }
  }
  for (const s of job.shortage || []) { // try to cover shortfall from stock received late
    let left = s.qty;
    for (const r of db.stock.filter((r) => usable(r, s.kind, s.model, job))) {
      const take = Math.min(left, Math.max(0, avail(r)));
      if (take > 0) { r.qty -= take; left -= take; registerInstalled(r, unitsMove(r, 'In stock', 'Installed', take, job), job); move('INSTALL', r, take, job.ref, by); }
    }
    if (left > 0) throw new Error(`Cannot mark installed: ${left} x ${s.model} not in stock. Receive stock first.`);
  }
  job.reserved = []; job.shortage = [];
}
function addReturned(job, qty, cond, by) {
  if (job.type === 'RATIONALIZATION' && cond === 'Serviceable') cond = 'Reusable';
  const kind = job.type === 'STP_REPLACEMENT' ? 'STP' : 'DU';
  const model = job.oldModel || (kind === 'DU' ? 'Suction DU (removed)' : 'STP (removed)');
  const store = db.settings.stores[0] || 'Main Store';
  let r = db.stock.find((r) => r.kind === kind && r.model === model && r.cond === cond && r.store === store);
  if (!r) { r = { id: uid('st'), kind, model, cond, store, qty: 0, reserved: 0 }; db.stock.push(r); }
  r.qty += qty; move('RETURN', r, qty, job.ref, by);
}

/* ---------- email ---------- */
const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const fmtDate = (s) => (s ? parse(s).toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' }) : '-');
const TYPE_LABEL = { DU_REPLACEMENT: 'DU Replacement (Suction to Pressure)', RATIONALIZATION: 'DU Rationalization', STP_REPLACEMENT: 'STP Replacement' };

function buildMail(job, event, extraCc = []) {
  const o = outletOf(job) || {}; const t = teamOf(job) || {}; const v = db.vendors.find((v) => v.id === t.vendorId) || {};
  const s = db.settings;
  const so = soOf(o); const to = [so.email].filter(Boolean);
  const cc = [...new Set([o.rmEmail || s.rm?.email, s.engHead?.email, ...(s.defaultCc || []), t.email, v.email, ...extraCc, ...(job.extraCc || [])].filter(Boolean))].filter((e) => !to.includes(e));
  const rows = [
    ['Reference', job.ref], ['Work', TYPE_LABEL[job.type]], ['Retail Outlet', `${o.name || job.outletCode} (${job.outletCode})`],
    ['Location', [o.location, o.district].filter(Boolean).join(', ')], ['Dealer', [o.dealerName, o.dealerPhone].filter(Boolean).join(' / ') || '-'], ['Region / Sales Area', `${o.region || '-'} / ${o.salesArea || '-'}`],
    ['Team deployed', `${t.name || '-'} (${v.name || '-'})`], ['Team members', (t.members || []).filter((m) => m.name || m.phone).map((m) => `${m.name || '-'}${m.phone ? ' (' + m.phone + ')' : ''}`).join(', ') || '-'],
    ['Scheduled', `${fmtDate(job.plannedStart)} to ${fmtDate(job.plannedEnd)}`],
  ];
  const scope = needs(job).map((n) => `${n.qty} x ${n.model}`).join('; ') || (job.oldQty ? `Removal of ${job.oldQty} DU` : '-');
  rows.push(['Scope', scope]);
  if (o.driveway) rows.push(['Driveway', o.driveway]);
  if (job.type === 'DU_REPLACEMENT') rows.push(['Automation cabling', job.cabling === 'Required' ? 'Required' : 'Not required']);
  const lay = (job.files || []).filter((f) => f.kind === 'layout'), pt = job.permit || {};
  rows.push(['Work permit', pt.status === 'Issued' ? ['Issued', pt.type === 'Other' ? pt.typeText : pt.type, pt.ref].filter(Boolean).join(' - ') : 'Not yet issued']);
  if (job.type === 'DU_REPLACEMENT') rows.push(['Approved layout', lay.length ? 'Attached to this mail: ' + lay.map((f) => f.name).join(', ') : 'Not yet available']);
  if (job.remarks) rows.push(['Remarks', job.remarks]);
  let subject, intro;
  const mt = s.mail || {}, fill = (tpl) => String(tpl).replace(/\{(\w+)\}/g, (_, k) => ({ ref: job.ref, work: TYPE_LABEL[job.type], outlet: o.name || job.outletCode, code: job.outletCode, date: fmtDate(job.plannedStart), team: t.name || '' }[k] ?? ''));
  if (event === 'deployed') {
    subject = fill(mt.deploySubject || '[{ref}] Team deployed - {work} at {outlet} ({date})');
    intro = mt.deployIntro || 'This is to inform you that a team has been deployed for the following work at your retail outlet.';
  } else {
    subject = fill(mt.completeSubject || '[{ref}] Completed - {work} at {outlet}');
    intro = mt.completeIntro || 'The following work has been reported as commissioned/completed at your retail outlet.';
    rows.push(['Completed on', fmtDate(job.stageDates?.commissioned?.date || job.stageDates?.dismantled?.date)]);
    if (job.commissioning?.nozzleSale) rows.push(['Nozzle sale started', job.commissioning.nozzleSale === 'Yes' ? 'Yes, from ' + fmtDate(job.commissioning.saleStart) : 'No - ' + job.commissioning.saleReason]);
  }
  const html = `<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#1b2a41"><p>Dear ${esc(so.name || 'Sir/Madam')},</p><p>${esc(intro)}</p>
<table style="border-collapse:collapse;min-width:420px">${rows.map(([k, v]) => `<tr><td style="padding:6px 12px;border:1px solid #d5dbe5;background:#f3f6fb;font-weight:600">${esc(k)}</td><td style="padding:6px 12px;border:1px solid #d5dbe5">${esc(v)}</td></tr>`).join('')}</table>
<p style="margin-top:16px">Regards,<br>${esc(s.senderName || 'Retail Engineering')}<br>${esc(s.orgLine || '')}</p>
<p style="color:#7a8699;font-size:12px">System generated mail from the Retail Engineering Project Tracker.</p></div>`;
  const text = `Dear ${so.name || 'Sir/Madam'},\n\n${intro}\n\n${rows.map(([k, v]) => `${k}: ${v}`).join('\n')}\n\nRegards,\n${s.senderName || 'Retail Engineering'}\n${s.orgLine || ''}`;
  return { to, cc, subject, html, text, attachments: lay.map((f) => ({ filename: f.name, path: path.join(UPLOADS, job.id, f.id + f.ext) })) };
}
async function sendMail(job, event, extraCc, by) {
  const m = buildMail(job, event, extraCc), atts = m.attachments; delete m.attachments;
  const rec = { id: uid('ml'), at: new Date().toISOString(), event, jobId: job.id, ref: job.ref, by, ...m, attachNames: atts.map((a) => a.filename), status: 'preview', error: null };
  const smtp = db.settings.smtp || {};
  if (!m.to.length) { rec.status = 'failed'; rec.error = 'No Sales Officer email (set it under Outlets > Sales Officers by sales area)'; }
  else if (smtp.host) {
    try {
      const tr = nodemailer.createTransport({ host: smtp.host, port: Number(smtp.port) || 587, secure: !!smtp.secure, auth: smtp.user ? { user: smtp.user, pass: smtp.pass } : undefined });
      await tr.sendMail({ from: smtp.from || smtp.user, to: m.to, cc: m.cc, subject: m.subject, html: m.html, text: m.text, attachments: atts });
      rec.status = 'sent';
    } catch (e) { rec.status = 'failed'; rec.error = e.message; }
  } else rec.error = 'SMTP not configured - saved as preview only';
  db.outbox.unshift(rec); db.outbox = db.outbox.slice(0, 500);
  return rec;
}

/* ---------- actions ---------- */
function markRemoved(job) {
  for (const id of job.assetIds || []) { const a = db.assets.find((a) => a.id === id); if (a) { a.status = 'Removed'; a.removedOn = today(); a.removedBy = job.ref; } }
  if (job.type === 'DU_REPLACEMENT') autoScrap((job.assetIds || []).map((id) => db.assets.find((a) => a.id === id)), 'system', job.ref);
}
function restoreAssets(job) {
  undoAutoScrap(job); for (const id of job.assetIds || []) { const a = db.assets.find((a) => a.id === id); if (a && a.removedBy === job.ref) { a.status = 'Active'; delete a.removedOn; delete a.removedBy; } } }
const PHOTO_STAGES = ['survey', 'dismantled', 'installed', 'cabling', 'calibrated', 'commissioned', 'old_returned'];
function newRef(type) {
  const p = { DU_REPLACEMENT: 'DU', RATIONALIZATION: 'RZ', STP_REPLACEMENT: 'STP' }[type];
  db.seq[p] = (db.seq[p] || 0) + 1;
  return `${p}-${String(db.seq[p]).padStart(4, '0')}`;
}
function cablingFor(b) {
  const c = cablingBase(b); c.cablingStatus = c.cabling === 'Required' ? 'Incomplete' : c.cabling === 'Not required' ? 'Not required' : 'N/A'; return c;
}
function cablingBase(b) {
  if (b.type !== 'DU_REPLACEMENT') return { cabling: 'N/A', cablingReason: '' };
  if (b.cabling === 'Required' || b.cabling === 'Not required') return { cabling: b.cabling, cablingReason: b.cablingReason || 'Set manually' };
  const o = db.outlets.find((o) => o.code === b.outletCode) || {};
  const none = (db.settings.noCablingPhases || []).map(String).includes(String(o.phase ?? '').trim());
  return none ? { cabling: 'Not required', cablingReason: `Phase ${o.phase} site - automation cabling already available` } : { cabling: 'Required', cablingReason: o.phase ? `Phase ${o.phase} site - automation cabling needed` : 'Automation phase unknown - cabling assumed required' };
}
function makeJob(b, by) {
  if (!STAGES[b.type]) throw new Error('Unknown work type');
  if (!db.outlets.find((o) => o.code === b.outletCode)) throw new Error(`Outlet ${b.outletCode} not found - add it under Outlets first`);
  const j = {
    id: uid('job'), ref: newRef(b.type), type: b.type, outletCode: b.outletCode, due: b.due || null, priority: b.priority || 'Medium',
    newModel: b.newModel || '', newQty: +b.newQty || 0, stpModel: b.stpModel || '', stpQty: +b.stpQty || 0, oldModel: b.oldModel || '', oldQty: +b.oldQty || 0,
    remarks: b.remarks || '', reason: b.reason || '', extraCc: b.extraCc || [], source: b.source === 'Reuse' ? 'Reuse' : 'New', permit: { status: 'Not issued' }, automation: { status: b.type === 'DU_REPLACEMENT' ? 'Pending' : 'N/A' }, vendorId: b.vendorId || null, planDate: b.planDate || null, seq: b.seq || null, avgSales: b.avgSales ?? null, assetIds: b.assetIds || [], planSource: b.planSource || '', daysOverride: +b.daysOverride || 0, ...cablingFor(b), files: [], teamId: null, plannedStart: null, plannedEnd: null,
    stage: STAGES[b.type][0].key, stageDates: {}, estimate: { amount: null, status: 'Not submitted' }, reserved: [], shortage: [], history: [],
  };
  j.stageDates[j.stage] = { date: today(), by };
  log(j, 'Job created', by);
  db.jobs.push(j); return j;
}
async function deployJob(job, { teamId, start, end, extraCc = [], mail = true }, by) {
  const team = db.teams.find((t) => t.id === teamId);
  if (!team) throw new Error('Team not found');
  if (db.settings.permitRequired === 'block' && job.permit?.status !== 'Issued') throw new Error('Work permit has not been issued for this outlet yet (change this rule under Settings > Work rules)');
  if (stageIdx(job) >= stageIdx(job, 'deployed')) throw new Error('Job already deployed');
  job.teamId = teamId; job.vendorId = team.vendorId; job.plannedStart = start; job.plannedEnd = end || addWork(start, duration(job) - 1);
  job.stage = 'deployed'; job.stageDates.deployed = { date: today(), by };
  reserve(job, by);
  log(job, `Team ${team.name} deployed ${fmtDate(job.plannedStart)} to ${fmtDate(job.plannedEnd)}` + (job.shortage.length ? ` (STOCK SHORT: ${job.shortage.map((s) => s.qty + ' ' + s.model).join(', ')})` : ''), by);
  let rec = null;
  if (mail) { rec = await sendMail(job, 'deployed', extraCc, by); log(job, `Deployment mail ${rec.status} to ${rec.to.join(', ') || 'nobody'} (cc ${rec.cc.length})`, 'system'); }
  return rec;
}
async function setStage(job, key, p, user) {
  const list = stageList(job); const to = list.findIndex((s) => s.key === key); const cur = stageIdx(job);
  if (to < 0) throw new Error('Invalid stage');
  const admin = user.role === 'admin';
  const st = list[to];
  if (to === cur) throw new Error('Already at this stage');
  if (to < cur) {
    if (!admin) throw new Error('Only admin can move a stage backwards');
    if (cur >= stageIdx(job, 'installed') && job.type !== 'RATIONALIZATION') throw new Error('Cannot step back after installation (inventory already updated)');
    if (job.type === 'RATIONALIZATION' && cur >= stageIdx(job, 'old_returned')) throw new Error('Cannot step back after units returned to store');
    if (cur >= stageIdx(job, 'deployed') && to < stageIdx(job, 'deployed')) { release(job, user.name); job.teamId = null; job.plannedStart = job.plannedEnd = null; }
    if (cur >= stageIdx(job, 'dismantled') && to < stageIdx(job, 'dismantled')) restoreAssets(job);
    for (const s of list.slice(to + 1)) delete job.stageDates[s.key];
    job.stage = key; log(job, `Stage moved back to "${st.label}"${p.note ? ': ' + p.note : ''}`, user.name); return;
  }
  if (!admin) {
    if (st.by !== 'V') throw new Error(`"${st.label}" is updated by HPCL (admin)`);
    if (to !== cur + 1) throw new Error('Update stages in sequence');
    if (list[cur + 1].key !== key) throw new Error('Update stages in sequence');
  }
  if (key === 'deployed') throw new Error('Use Deploy team (this sends the mail to Sales Officer)');
  if (key === 'cabling' && (p.cablingStatus || 'Completed') !== 'Completed') {
    if (!['In progress', 'Incomplete'].includes(p.cablingStatus)) throw new Error('Invalid cabling status');
    job.cablingStatus = p.cablingStatus; log(job, `Automation cabling: ${p.cablingStatus}${p.note ? ' - ' + p.note : ''}`, user.name); return;
  }
  if (!admin && !vendorOf(job)) throw new Error('Job not allotted to a vendor');
  if (to > stageIdx(job, 'deployed') && !job.teamId) throw new Error('Deploy a team first');
  const photos = (job.files || []).filter((f) => f.kind === 'photo' && f.stage === key).length;
  const need = db.settings.photoOverrides?.[key] ?? db.settings.photosPerStage ?? 4;
  if (!admin && (db.settings.photoStages || PHOTO_STAGES).includes(key) && photos < need) throw new Error(`Upload ${need} photos for "${st.label}" first (you have ${photos})`);
  if (db.settings.ackRequired !== false && ['old_returned', 'closed'].includes(key) && !(job.files || []).some((f) => f.kind === 'ack')) throw new Error("Upload the dealer's acknowledgement letter first");
  if (key === 'calibrated') {
    if (!admin && !p.calibrationDone) throw new Error('Confirm that calibration & sealing is done');
    job.commissioning = { ...(job.commissioning || {}), calibrationDone: !!p.calibrationDone, sealNo: p.sealNo || '' };
  }
  if (key === 'commissioned') {
    if (!['Yes', 'No'].includes(p.nozzleSale)) throw new Error('Tell us whether nozzle sale has started (Yes/No)');
    if (p.nozzleSale === 'No' && !String(p.saleReason || '').trim()) throw new Error('Give the reason nozzle sale has not started');
    job.commissioning = { ...(job.commissioning || {}), commissioned: true, nozzleSale: p.nozzleSale, saleStart: p.nozzleSale === 'Yes' ? (p.saleStart || p.date || today()) : '', saleReason: p.saleReason || '' };
  }
  if (key === 'est_submitted') {
    const amt = Number(p.amount); if (!(amt > 0)) throw new Error('Enter the estimate amount');
    job.estimate = { amount: amt, status: 'Submitted', ref: p.estimateRef || '', submittedOn: p.date || today() };
  }
  if (key === 'est_approved') { job.estimate.status = 'Approved'; if (p.approvedAmount) job.estimate.approvedAmount = Number(p.approvedAmount); }
  if (key === 'installed') consume(job, user.name);
  if (key === 'old_returned') {
    const q = Number(p.returnedQty ?? job.oldQty ?? job.newQty);
    if (q > 0) addReturned(job, q, p.condition || 'Serviceable', user.name);
    job.returned = { qty: q, cond: p.condition || 'Serviceable' };
  }
  if (key === 'dismantled') markRemoved(job);
  if (key === 'cabling') job.cablingStatus = 'Completed';
  job.stage = key; job.stageDates[key] = { date: p.date || today(), by: user.name, note: p.note || '' };
  log(job, `${st.label}${p.note ? ': ' + p.note : ''}${key === 'est_submitted' ? ` (Rs ${job.estimate.amount})` : ''}`, user.name);
  if (key === 'commissioned' || (job.type === 'RATIONALIZATION' && key === 'dismantled')) {
    if (p.notify !== false && db.settings.mailOnComplete) { const rec = await sendMail(job, 'completed', [], user.name); log(job, `Completion mail ${rec.status}`, 'system'); }
  }
}

const lc = (x) => String(x || '').toLowerCase();
const covers = (t, o) => !t.regions?.length || t.regions.some((a) => [o.region, o.zone, o.district, o.salesArea].some((v) => lc(v) === lc(a)));
function autoPlan() {
  const t0 = nextWork(addDays(today(), 1));
  const pend = db.jobs.filter((j) => !j.teamId && stageIdx(j) < stageIdx(j, 'deployed'));
  const rank = { High: 0, Medium: 1, Low: 2 };
  pend.sort((a, b) => (rank[a.priority] - rank[b.priority]) || ((a.seq || 999) - (b.seq || 999)) || String(a.due || '9999').localeCompare(String(b.due || '9999')));
  const free = {};
  for (const t of db.teams.filter((t) => t.active !== false)) {
    const busy = db.jobs.filter((j) => j.teamId === t.id && j.plannedEnd && stageIdx(j) < stageIdx(j, 'old_returned')).map((j) => j.plannedEnd);
    const last = busy.sort().pop();
    free[t.id] = last && last >= t0 ? nextWork(addDays(last, 1)) : t0;
  }
  const have = {};
  for (const r of db.stock) { const pool = r.cond === 'Reusable' ? 'Reuse' : (r.cond === 'New' || r.cond === 'Refurbished') ? 'New' : null; if (pool) have[pool + '|' + r.kind + '|' + r.model] = (have[pool + '|' + r.kind + '|' + r.model] || 0) + avail(r); }
  const out = [];
  for (const j of pend) {
    const o = outletOf(j) || {};
    const cands = db.teams.filter((t) => t.active !== false && covers(t, o) && (!vendorOf(j) || t.vendorId === vendorOf(j)));
    if (!cands.length) { out.push({ jobId: j.id, ref: j.ref, error: `No active team ${vendorOf(j) ? 'of the allotted vendor ' : ''}covers ${o.salesArea || o.district || 'this area'}` }); continue; }
    cands.sort((a, b) => free[a.id].localeCompare(free[b.id]));
    const t = cands[0]; const start = j.planDate && j.planDate > free[t.id] ? nextWork(j.planDate) : free[t.id]; const end = addWork(start, duration(j) - 1);
    const short = [];
    for (const n of needs(j)) { const k = (j.source === 'Reuse' ? 'Reuse' : 'New') + '|' + n.kind + '|' + n.model; if ((have[k] || 0) < n.qty) short.push(`${n.qty - Math.max(0, have[k] || 0)} x ${n.model}`); have[k] = (have[k] || 0) - n.qty; }
    const flags = [];
    if (j.due && end > j.due) flags.push('Will miss due date');
    if (j.planDate && j.planDate < t0) flags.push('Planned date ' + j.planDate + ' has passed');
    if (short.length) flags.push('Stock short: ' + short.join(', '));
    if (stageIdx(j) < stageIdx(j, 'material')) flags.push('Estimate/material not cleared yet');
    if (j.permit?.status !== 'Issued') flags.push('Work permit not issued');
    out.push({ jobId: j.id, ref: j.ref, teamId: t.id, teamName: t.name, start, end, flags });
    free[t.id] = nextWork(addDays(end, 1));
  }
  return out;
}

/* ---------- view for client ---------- */
function stateFor(user) {
  const s = JSON.parse(JSON.stringify(db.settings));
  if (s.smtp?.pass) s.smtp.pass = '********';
  const base = { me: { username: user.username, name: user.name, role: user.isEditor ? 'editor' : user.role, vendorId: user.vendorId || null }, stages: STAGES, today: today() };
  if (user.role !== 'vendor') { // admin and read-only viewer see everything
    return { ...base, assetStats: assetStats(), hh: hhState(), bills: db.bills, recovered: db.recovered, reuses: db.reuses, scrapReports: db.scrapReports, recTypes: REC_TYPES, scrap: db.scrap, settings: s, vendors: db.vendors, teams: db.teams, outlets: db.outlets, jobs: db.jobs, stock: db.stock, movements: db.movements.slice(-300).reverse(),
      outbox: db.outbox.map(({ html, ...r }) => r), users: user.role !== 'admin' || user.isEditor ? [] : db.users.map((u) => ({ username: u.username, name: u.name, role: u.role, vendorId: u.vendorId, mustChange: !!u.mustChange })) };
  }
  const teams = db.teams.filter((t) => t.vendorId === user.vendorId);
  const jobs = db.jobs.filter((j) => vendorOf(j) === user.vendorId);
  const outlets = db.outlets.filter((o) => jobs.some((j) => j.outletCode === o.code) || db.scrap.some((x) => x.vendorId === user.vendorId && x.outletCode === o.code)).map(({ soEmail, soPhone, soName, rmEmail, rmName, dealerEmails, ...o }) => o);
  const myOutlets = new Set(jobs.map((j) => j.outletCode));
  return { ...base, bills: db.bills.filter((b) => b.vendorId === user.vendorId), scrap: db.scrap.filter((x) => x.vendorId === user.vendorId || myOutlets.has(x.outletCode)), settings: { durations: s.durations, mapsKey: s.mapsKey, photosPerStage: s.photosPerStage, photoStages: s.photoStages, photoOverrides: s.photoOverrides, weekOff: s.weekOff, areaColors: s.areaColors, refreshSec: s.refreshSec, brand: s.brand, theme: s.theme, logoV: s.logoV, dash: s.dash, noCablingPhases: s.noCablingPhases, cablingDays: s.cablingDays, legacyYears: s.legacyYears, dashOrder: s.dashOrder, vendorMenu: s.vendorMenu, dashTiles: s.dashTiles }, vendors: db.vendors.filter((v) => v.id === user.vendorId), teams, outlets, jobs, stock: [], movements: [], outbox: [], users: [] };
}

/* ---------- routes ---------- */
const routes = [];
const route = (method, pattern, admin, fn) => routes.push({ method, re: new RegExp('^' + pattern.replace(/:(\w+)/g, '(?<$1>[^/]+)') + '$'), admin, fn });
const bad = (m, code = 400) => Object.assign(new Error(m), { code });
const pick = (o, keys) => Object.fromEntries(keys.filter((k) => k in o).map((k) => [k, o[k]]));
const jobFor = (id) => { const j = db.jobs.find((j) => j.id === id); if (!j) throw bad('Job not found', 404); return j; };
const canSee = (user, job) => user.role !== 'vendor' || vendorOf(job) === user.vendorId;

const editorBlocked = (m, p) => p !== '/api/password' && (m === 'DELETE' || p.startsWith('/api/users') || p.startsWith('/api/admin/') || p.startsWith('/api/logo') || (p === '/api/settings' && m !== 'GET') || /\/api\/hh\/issues\/[^/]+\/cancel$/.test(p));
const FAILS = new Map(); // login throttle: ip+username -> { n, first }
route('POST', '/api/login', false, ({ body, ip }) => {
  const name = String(body.username || '').trim().toLowerCase(), key = ip + '|' + name, f = FAILS.get(key);
  if (f && Date.now() - f.first < 10 * 60e3 && f.n >= 6) throw bad('Too many wrong attempts. Please wait 10 minutes and try again.', 429);
  const u = db.users.find((u) => u.username === name);
  if (!u || !checkPw(String(body.password || ''), u)) {
    const cur = f && Date.now() - f.first < 10 * 60e3 ? f : { n: 0, first: Date.now() }; cur.n++; FAILS.set(key, cur);
    throw bad('Wrong username or password', 401);
  }
  FAILS.delete(key);
  return { token: sign({ u: u.username, exp: Date.now() + Math.min(168, Math.max(1, +db.settings.sessionHours || 12)) * 3600e3 }), mustChange: !!u.mustChange };
});
route('GET', '/api/state', false, ({ user }) => stateFor(user));
route('POST', '/api/password', false, ({ user, body }) => {
  const pw = String(body.password || '');
  if (!checkPw(String(body.current || ''), user)) throw bad('Your current password is not correct');
  if (pw.length < 8) throw bad('The new password must be at least 8 characters');
  if (['admin123', 'vendor123', 'password', '12345678'].includes(pw.toLowerCase()) || pw.toLowerCase() === user.username) throw bad('Please choose a password that is harder to guess');
  if (pw === String(body.current)) throw bad('The new password must be different from the current one');
  const rec = db.users.find((u) => u.username === user.username); Object.assign(rec, hashPw(pw)); rec.mustChange = false; save(); return { ok: true };
});

function checkAssets(ids, outletCode, exceptJobId, type) {
  const kind = type === 'STP_REPLACEMENT' ? 'STP' : 'DU', cur = new Set(db.jobs.find((j) => j.id === exceptJobId)?.assetIds || []);
  const inJob = new Set(db.jobs.filter((j) => j.stage !== 'closed' && j.id !== exceptJobId).flatMap((j) => j.assetIds || []));
  for (const id of ids) {
    const a = db.assets.find((x) => x.id === id); if (!a) throw bad('One of the selected units was not found');
    if (a.sap !== outletCode) throw bad(a.serial + ' is not installed at this outlet');
    if (a.kind !== kind) throw bad(a.serial + ' is a ' + a.kind + ', but this job is for ' + kind + 's');
    if (a.status === 'Removed' && !cur.has(id)) throw bad(a.serial + ' has already been removed');
    if (inJob.has(id)) throw bad(a.serial + ' is already part of another open job');
  }
  return [...new Set(ids)];
}
// Rule: one open DU replacement job per outlet - more suction DUs at the same outlet are added to that job.
const openDuJob = (code, except) => db.jobs.find((j) => j.type === 'DU_REPLACEMENT' && j.outletCode === code && j.stage !== 'closed' && j.id !== except);
function mergeIntoJob(j, b, by) {
  const add = (b.assetIds || []).filter((id) => !(j.assetIds || []).includes(id)), n = add.length || Math.max(0, +b.oldQty || +b.newQty || 0);
  j.assetIds = [...(j.assetIds || []), ...add]; j.oldQty = (+j.oldQty || 0) + n; j.newQty = (+j.newQty || 0) + n; if (+b.stpQty) j.stpQty = (+j.stpQty || 0) + +b.stpQty;
  const ser = add.map((id) => db.assets.find((a) => a.id === id)?.serial).filter(Boolean);
  if (!j.remarks || j.remarks.startsWith('Old DU serials:')) j.remarks = serialsNote(j.assetIds); else if (ser.length) j.remarks += '; added DU serials: ' + ser.join(', ');
  log(j, `${n} more suction DU(s) added to this job (one job per outlet)${ser.length ? ': ' + ser.join(', ') : ''}`, by);
  return n;
}
route('POST', '/api/jobs', true, ({ user, body }) => {
  if (body.assetIds?.length) body.assetIds = checkAssets(body.assetIds, body.outletCode, null, body.type);
  if (body.type === 'DU_REPLACEMENT') { const ex = openDuJob(body.outletCode); if (ex) { mergeIntoJob(ex, body, user.name); save(); return { ...ex, merged: true }; } }
  const j = makeJob(body, user.name); save(); return j;
});
route('POST', '/api/jobs/bulk', true, ({ user, body }) => {
  const made = [], errors = [];
  (body.rows || []).forEach((r, i) => { try { made.push(makeJob(r, user.name).ref); } catch (e) { errors.push(`Row ${i + 1}: ${e.message}`); } });
  save(); return { created: made.length, errors };
});
route('PATCH', '/api/jobs/:id', true, ({ user, params, body }) => {
  const j = jobFor(params.id);
  if (j.type === 'DU_REPLACEMENT' && body.outletCode && body.outletCode !== j.outletCode) { const ex = openDuJob(body.outletCode, j.id); if (ex) throw bad(`That outlet already has DU replacement job ${ex.ref} - add the DUs there (one job per outlet)`); }
  Object.assign(j, pick(body, ['due', 'priority', 'newModel', 'newQty', 'stpModel', 'stpQty', 'oldModel', 'oldQty', 'remarks', 'reason', 'extraCc', 'outletCode', 'source', 'vendorId', 'planDate']));
  ['newQty', 'stpQty', 'oldQty'].forEach((k) => (j[k] = +j[k] || 0));
  if ('assetIds' in body) j.assetIds = checkAssets(body.assetIds || [], j.outletCode, j.id, j.type);
  else if (body.outletCode && (j.assetIds || []).some((id) => db.assets.find((a) => a.id === id)?.sap !== j.outletCode)) j.assetIds = [];
  if ('daysOverride' in body) j.daysOverride = +body.daysOverride || 0;
  if (body.cabling && j.type === 'DU_REPLACEMENT' && body.cabling !== j.cabling) {
    if (stageIdx(j) >= stageIdx(j, 'installed')) throw bad('Cabling requirement cannot change after installation');
    Object.assign(j, cablingFor({ type: j.type, outletCode: j.outletCode, cabling: body.cabling, cablingReason: body.cablingReason }));
  }
  if (body.plannedStart && j.teamId) { j.plannedStart = body.plannedStart; j.plannedEnd = body.plannedEnd || addWork(body.plannedStart, duration(j) - 1); }
  log(j, 'Job details edited', user.name); save(); return j;
});
route('DELETE', '/api/jobs/:id', true, ({ params, user }) => {
  const j = jobFor(params.id); if (stageIdx(j) >= stageIdx(j, 'installed') && j.type !== 'RATIONALIZATION') throw bad('Installed jobs cannot be deleted');
  release(j, user.name); db.jobs = db.jobs.filter((x) => x.id !== j.id); save(); return { ok: true };
});
route('POST', '/api/jobs/:id/stage', false, async ({ user, params, body }) => {
  const j = jobFor(params.id); if (!canSee(user, j)) throw bad('Not your job', 403);
  await setStage(j, body.stage, body, user); save(); return j;
});
route('POST', '/api/jobs/:id/note', false, ({ user, params, body }) => {
  const j = jobFor(params.id); if (!canSee(user, j)) throw bad('Not your job', 403);
  if (!String(body.note || '').trim()) throw bad('Empty note'); log(j, 'Note: ' + body.note.trim(), user.name); save(); return j;
});
route('POST', '/api/jobs/:id/deploy', true, async ({ user, params, body }) => {
  const j = jobFor(params.id); if (!body.start) throw bad('Start date required');
  const rec = await deployJob(j, { teamId: body.teamId, start: body.start, end: body.end, extraCc: body.extraCc, mail: body.mail !== false }, user.name);
  save(); return { job: j, mail: rec && { status: rec.status, to: rec.to, cc: rec.cc, error: rec.error } };
});
route('POST', '/api/plan/auto', true, () => autoPlan());
route('POST', '/api/plan/apply', true, async ({ user, body }) => {
  const res = [];
  for (const p of body.items || []) {
    try { const j = jobFor(p.jobId); const rec = await deployJob(j, { teamId: p.teamId, start: p.start, end: p.end, mail: body.mail !== false }, user.name); res.push({ ref: j.ref, mail: rec?.status || 'not sent' }); }
    catch (e) { res.push({ ref: p.ref, error: e.message }); }
  }
  save(); return res;
});

const UPLOADS = path.join(__dirname, 'data', 'uploads');
const sLabelOf = (j, key) => stageList(j).find((s) => s.key === key)?.label || key;
route('POST', '/api/jobs/:id/files', false, ({ user, params, body }) => {
  const j = jobFor(params.id); if (!canSee(user, j)) throw bad('Not your job', 403);
  if (j.stage === 'closed') throw bad('Job is closed');
  const kind = ['ack', 'permit', 'layout'].includes(body.kind) ? body.kind : 'photo';
  if (kind === 'layout' && user.role !== 'admin') throw bad('Only HPCL can upload the approved layout', 403);
  if (kind === 'photo' && !stageList(j).some((s) => s.key === body.stage)) throw bad('Invalid stage');
  const mime = String(body.mime || '');
  if (!/^(image\/(jpeg|png|webp)|application\/pdf)$/.test(mime)) throw bad('Only JPG, PNG, WEBP or PDF allowed');
  if (kind === 'photo' && mime === 'application/pdf') throw bad('Photos must be images');
  const buf = Buffer.from(String(body.data || ''), 'base64');
  if (!buf.length || buf.length > 9e6) throw bad('File empty or too large (max 9 MB)');
  const id = uid('f'); const ext = mime === 'application/pdf' ? '.pdf' : mime === 'image/png' ? '.png' : mime === 'image/webp' ? '.webp' : '.jpg';
  fs.mkdirSync(path.join(UPLOADS, j.id), { recursive: true }); fs.writeFileSync(path.join(UPLOADS, j.id, id + ext), buf);
  const f = { id, ext, kind, stage: kind === 'photo' ? body.stage : kind, name: String(body.name || 'file').slice(0, 80), mime, at: new Date().toISOString(), by: user.name, size: buf.length };
  j.files.push(f); log(j, { ack: 'Dealer acknowledgement letter uploaded', permit: 'Work permit document uploaded', layout: 'Approved layout uploaded' }[kind] || 'Photo added (' + sLabelOf(j, body.stage) + ')', user.name); save(); return f;
});
route('DELETE', '/api/jobs/:id/files/:fid', false, ({ user, params }) => {
  const j = jobFor(params.id); if (!canSee(user, j)) throw bad('Not your job', 403);
  const f = j.files.find((f) => f.id === params.fid); if (!f) throw bad('Not found', 404);
  if (j.stage === 'closed') throw bad('Job is closed');
  if (f.kind === 'layout' && user.role !== 'admin') throw bad('Only HPCL can remove the layout', 403);
  if (user.role !== 'admin' && stageIdx(j, f.stage) >= 0 && stageIdx(j) > stageIdx(j, f.stage)) throw bad('That stage is already submitted');
  j.files = j.files.filter((x) => x !== f); fs.rmSync(path.join(UPLOADS, j.id, f.id + f.ext), { force: true }); save(); return { ok: true };
});
route('GET', '/api/files/:jid/:fid', false, ({ user, params }) => {
  const j = jobFor(params.jid); if (!canSee(user, j)) throw bad('Not your job', 403);
  const f = j.files.find((f) => f.id === params.fid); if (!f) throw bad('Not found', 404);
  return { __file: path.join(UPLOADS, j.id, f.id + f.ext), mime: f.mime, name: f.name };
});
const km = (a, b) => { const r = Math.PI / 180, dLa = (b.lat - a.lat) * r, dLo = (b.lng - a.lng) * r; const x = Math.sin(dLa / 2) ** 2 + Math.cos(a.lat * r) * Math.cos(b.lat * r) * Math.sin(dLo / 2) ** 2; return 12742 * Math.asin(Math.sqrt(x)); };
route('GET', '/api/nearest/:id', false, ({ user, params }) => {
  const j = jobFor(params.id); if (!canSee(user, j)) throw bad('Not your job', 403);
  const from = outletOf(j); if (!from || from.lat == null) return { error: 'Add latitude/longitude for this outlet to get suggestions', list: [] };
  const rank = { High: 10, Medium: 0, Low: -5 };
  const list = db.jobs.filter((x) => x.id !== j.id && x.stage !== 'closed' && stageIdx(x) < stageIdx(x, 'dismantled') && (!x.teamId || x.teamId === j.teamId))
    .map((x) => {
      const o = outletOf(x); if (!o || o.lat == null) return null;
      const air = km(from, o); const ready = stageIdx(x) >= stageIdx(x, 'material');
      return { jobId: x.id, ref: x.ref, type: x.type, name: o.name, code: o.code, district: o.district, lat: o.lat, lng: o.lng, roadKm: Math.round(air * 1.3), ready, stage: stageList(x)[stageIdx(x)].label, priority: x.priority, due: x.due, unassigned: !x.teamId, score: air * 1.3 + (ready ? 0 : 40) - (rank[x.priority] || 0) };
    }).filter(Boolean).sort((a, b) => a.score - b.score).slice(0, 5);
  return { from: { name: from.name, lat: from.lat, lng: from.lng }, list };
});
route('POST', '/api/stock', true, ({ user, body }) => {
  const kind = body.kind === 'STP' ? 'STP' : 'DU', cond = body.cond || 'New', store = body.store || db.settings.stores[0], model = String(body.model || '').trim();
  if (!model) throw bad('Model is required');
  const make = String(body.make || '').trim(), nozzles = kind === 'DU' ? +body.nozzles || null : null, hp = kind === 'STP' ? String(body.hp || '') : '';
  const serials = String(body.serials || '').split(/[,;\n]+/).map((x) => x.trim()).filter(Boolean);
  let qty = Math.floor(+body.qty) || 0;
  if (serials.length) { if (qty && qty !== serials.length) throw bad(`You entered ${serials.length} serial number(s) but quantity ${qty} - they must match`); qty = serials.length; }
  if (!(qty > 0)) throw bad('Enter the quantity (or the serial numbers)');
  if (new Set(serials).size !== serials.length) throw bad('The same serial number is entered twice');
  for (const sn of serials) {
    if (db.stock.some((r) => r.kind === kind && (r.units || []).some((u) => u.serial === sn))) throw bad(`Serial ${sn} is already in stock`);
    if (db.assets.some((a) => a.kind === kind && a.serial === sn && a.status !== 'Removed')) throw bad(`Serial ${sn} is already installed at an outlet`);
  }
  let r = db.stock.find((r) => r.kind === kind && r.model === model && r.cond === cond && r.store === store && (r.make || '') === make && (r.nozzles || null) === nozzles && (r.hp || '') === hp);
  if (!r) { r = { id: uid('st'), kind, model, make, nozzles, hp, cond, store, qty: 0, reserved: 0, units: [] }; db.stock.push(r); }
  r.units ||= [];
  for (const sn of serials) r.units.push({ serial: sn, date: body.date || today(), ref: String(body.ref || ''), status: 'In stock' });
  r.qty += qty; move('RECEIVE', r, qty, (body.ref || 'Receipt') + (serials.length ? ' · S/N ' + serials.join(', ') : ''), user.name); save(); return r;
});
route('POST', '/api/stock/:id/adjust', true, ({ user, params, body }) => {
  const r = db.stock.find((r) => r.id === params.id); if (!r) throw bad('Not found', 404);
  const d = +body.delta; if (!d) throw bad('Enter a non-zero change'); if (r.qty + d < r.reserved) throw bad('Cannot go below reserved quantity');
  r.qty += d; move('ADJUST', r, d, body.reason || 'Manual adjustment', user.name); save(); return r;
});

const crud = (name, coll, key = 'id', prefix = name[0]) => {
  route('POST', `/api/${name}`, true, ({ body }) => {
    let rec = body[key] && db[coll].find((x) => x[key] === body[key]);
    if (rec) Object.assign(rec, body); else { rec = { ...body }; if (key === 'id') rec.id = uid(prefix); db[coll].push(rec); }
    save(); return rec;
  });
  route('DELETE', `/api/${name}/:k`, true, ({ params }) => { db[coll] = db[coll].filter((x) => x[key] !== params.k); save(); return { ok: true }; });
};
const normOutlet = (r) => { for (const k of ['lat', 'lng', 'dus']) r[k] = r[k] === '' || r[k] == null || isNaN(+r[k]) ? null : +r[k]; if (r.driveway) r.driveway = /rcc/i.test(r.driveway) ? 'RCC' : 'Paver'; return r; };
crud('vendors', 'vendors', 'id', 'v'); crud('teams', 'teams', 'id', 't');
route('POST', '/api/outlets', true, ({ body }) => {
  if (!body.code) throw bad('Outlet code required'); normOutlet(body);
  const e = db.outlets.find((o) => o.code === body.code); if (e) Object.assign(e, body); else db.outlets.push(body); recomputeCabling(body.code); save(); return body;
});
function recomputeCabling(code) {
  for (const j of db.jobs) {
    if (j.outletCode !== code || j.type !== 'DU_REPLACEMENT' || j.stage === 'closed' || stageIdx(j) >= stageIdx(j, 'installed')) continue;
    if (!/^(Phase|Automation phase)/.test(j.cablingReason || '')) continue; // manually set -> keep
    const c = cablingFor({ type: j.type, outletCode: code }); if (c.cabling === j.cabling) delete c.cablingStatus; Object.assign(j, c);
  }
}
route('DELETE', '/api/outlets/:k', true, ({ params }) => { db.outlets = db.outlets.filter((o) => o.code !== params.k); save(); return { ok: true }; });
route('POST', '/api/outlets/bulk', true, ({ body }) => {
  let n = 0;
  for (const r of body.rows || []) { if (!r.code) continue; normOutlet(r); const e = db.outlets.find((o) => o.code === r.code); if (e) Object.assign(e, r); else db.outlets.push(r); recomputeCabling(r.code); n++; }
  save(); return { saved: n };
});
route('POST', '/api/users', true, ({ body }) => {
  const un = String(body.username || '').trim().toLowerCase(), role = ['admin', 'editor', 'viewer', 'vendor'].includes(body.role) ? body.role : 'vendor';
  if (!/^[a-z0-9._-]{3,30}$/.test(un)) throw bad('Username: 3-30 letters, numbers, dot, dash or underscore');
  if (String(body.password || '').length < 6) throw bad('Give a temporary password of at least 6 characters');
  if (db.users.find((u) => u.username === un)) throw bad('That username is already taken');
  if (role === 'vendor' && !db.vendors.find((v) => v.id === body.vendorId)) throw bad('Choose the vendor this login belongs to');
  db.users.push({ username: un, name: String(body.name || un).slice(0, 80), role, vendorId: role === 'vendor' ? body.vendorId : null, mustChange: true, ...hashPw(body.password) });
  save(); return { ok: true };
});
route('POST', '/api/users/:k/reset', true, ({ params, body }) => {
  const u = db.users.find((x) => x.username === params.k); if (!u) throw bad('User not found', 404);
  if (String(body.password || '').length < 6) throw bad('Give a temporary password of at least 6 characters');
  Object.assign(u, hashPw(body.password)); u.mustChange = true;
  for (const k of [...FAILS.keys()]) if (k.endsWith('|' + u.username)) FAILS.delete(k);
  save(); return { ok: true };
});
route('DELETE', '/api/users/:k', true, ({ params, user }) => {
  if (params.k === user.username) throw bad('You cannot delete yourself');
  const t = db.users.find((u) => u.username === params.k); if (t?.role === 'admin' && db.users.filter((u) => u.role === 'admin').length < 2) throw bad('At least one admin must remain');
  db.users = db.users.filter((u) => u.username !== params.k); save(); return { ok: true };
});
route('PUT', '/api/settings', true, ({ body }) => {
  const keep = db.settings.smtp?.pass; Object.assign(db.settings, body);
  if (db.settings.smtp && db.settings.smtp.pass === '********') db.settings.smtp.pass = keep; save(); return { ok: true };
});
route('POST', '/api/email/test', true, async ({ body, user }) => {
  const smtp = db.settings.smtp || {}; if (!smtp.host) throw bad('Configure SMTP first');
  const tr = nodemailer.createTransport({ host: smtp.host, port: Number(smtp.port) || 587, secure: !!smtp.secure, auth: smtp.user ? { user: smtp.user, pass: smtp.pass } : undefined });
  try { await tr.sendMail({ from: smtp.from || smtp.user, to: body.to || user.username, subject: 'Test mail - Retail Engineering Tracker', text: 'SMTP is working.' }); return { ok: true }; }
  catch (e) { throw bad('SMTP error: ' + e.message); }
});
route('GET', '/api/outbox/:id', true, ({ params }) => { const r = db.outbox.find((r) => r.id === params.id); if (!r) throw bad('Not found', 404); return r; });
route('POST', '/api/outbox/:id/resend', true, async ({ params, user }) => {
  const r = db.outbox.find((r) => r.id === params.id); if (!r) throw bad('Not found', 404);
  const rec = await sendMail(jobFor(r.jobId), r.event, [], user.name); save(); return { status: rec.status, error: rec.error };
});
route('POST', '/api/admin/clear-demo', true, () => {
  resetData(); save(); return { ok: true };
});

/* ---------- installed DU / STP register, Excel import ---------- */
function legacyStats(planned) {
  const [y0, y1] = db.settings.legacyYears || [2011, 2013];
  const all = db.assets.filter((a) => a.kind === 'DU' && a.pumpType === 'Suction' && a.year >= y0 && a.year <= y1), act = all.filter((a) => a.status !== 'Removed');
  // Programme progress: a replaced (removed) DU counts 1; a DU still installed but inside an open job counts that job's stage progress.
  const jobOf = new Map(); for (const j of db.jobs) if (j.stage !== 'closed') for (const id of j.assetIds || []) jobOf.set(id, j);
  const part = (j) => { const L = stageList(j), i = stageIdx(j); return L.length > 1 && i > 0 ? i / (L.length - 1) : 0; };
  const removed = all.length - act.length, inProg = act.reduce((s, a) => s + (jobOf.has(a.id) ? part(jobOf.get(a.id)) : 0), 0);
  return { from: y0, to: y1, total: all.length, installed: act.length, removed, inJob: act.filter((a) => planned.has(a.id)).length, outlets: new Set(act.map((a) => a.sap)).size,
    inProgress: act.filter((a) => jobOf.has(a.id) && part(jobOf.get(a.id)) > 0).length, progressPct: all.length ? Math.round(((removed + inProg) / all.length) * 1000) / 10 : 0 };
}
function assetStats() {
  const act = db.assets.filter((a) => a.status !== 'Removed'), du = act.filter((a) => a.kind === 'DU');
  const planned = new Set(db.jobs.filter((j) => j.stage !== 'closed').flatMap((j) => j.assetIds || []));
  const suction = du.filter((a) => a.pumpType === 'Suction');
  const area = new Map(db.outlets.map((o) => [o.code, o.salesArea])), byArea = {}; suction.forEach((a) => { const k = area.get(a.sap) || 'Unmapped'; byArea[k] = (byArea[k] || 0) + 1; });
  return { byArea, duTotal: du.length, suction: suction.length, pressure: du.filter((a) => a.pumpType === 'Pressure').length, unknown: du.filter((a) => !a.pumpType).length, stp: act.filter((a) => a.kind === 'STP').length,
    legacy: legacyStats(planned), suctionPlanned: suction.filter((a) => planned.has(a.id)).length, suctionOutlets: new Set(suction.map((a) => a.sap)).size, removed: db.assets.filter((a) => a.status === 'Removed').length };
}
const VENDOR_CODES = { 'friends enterprises': 'FE', 'akash enterprises': 'AE', 'cms electricals': 'CMS', 'ab multi services': 'ABM', 'shivali engineering': 'SE' };
function resetData() {
  Object.assign(db, { jobs: [], outlets: [], stock: [], movements: [], outbox: [], assets: [], scrap: [], seq: {}, hhStock: {}, hhIssues: [], hhMoves: [], bills: [], recovered: [], autoPipe: [], reuses: [], scrapReports: [] });
  db.settings.salesAreas = {};
  for (const v of db.vendors) v.code ||= VENDOR_CODES[String(v.name).toLowerCase()] || '';
  for (const t of db.teams) t.regions = [];
  fs.rmSync(path.join(__dirname, 'data', 'uploads'), { recursive: true, force: true });
}
const mostCommon = (arr) => { const c = {}; arr.forEach((x) => (c[x] = (c[x] || 0) + 1)); return +Object.entries(c).sort((a, b) => b[1] - a[1])[0]?.[0] || 0; };

function upsertAssets(list, rep) {
  const idx = new Map(db.assets.map((a) => [a.kind + '|' + a.serial, a]));
  for (const a of list) {
    const k = a.kind + '|' + a.serial, e = idx.get(k);
    if (e) { const keep = { sap: e.sap, pumpType: e.pumpType, status: e.status, id: e.id, removedOn: e.removedOn, removedBy: e.removedBy }; Object.assign(e, a, { sap: a.sap || keep.sap, pumpType: a.pumpType || keep.pumpType, status: keep.status, id: keep.id, removedOn: keep.removedOn, removedBy: keep.removedBy }); rep.assets.updated++; }
    else { const n = { id: uid('a'), status: 'Active', ...a }; db.assets.push(n); idx.set(k, n); rep.assets.added++; }
  }
}
/* mark replaced DUs complete: closes fully-covered jobs (100%), splits partly-covered jobs */
function forceClose(job, by, note) {
  for (const st of stageList(job)) job.stageDates[st.key] ||= { date: today(), by, note };
  job.stage = 'closed'; log(job, note, by);
}
function removeAssets(ids, ref) { for (const id of ids) { const a = db.assets.find((x) => x.id === id); if (a) { a.status = 'Removed'; a.removedOn = today(); a.removedBy = ref; } } autoScrap(ids.map((id) => db.assets.find((x) => x.id === id)), 'system', ref); }
const serialsNote = (ids) => { const ser = ids.map((id) => db.assets.find((a) => a.id === id)?.serial).filter(Boolean); return ser.length ? 'Old DU serials: ' + ser.join(', ') : ''; };
function completeAssets(sap, assets, by, rep, src) {
  const ids = new Set(assets.map((a) => a.id)), handled = new Set(), note = `Marked complete from ${src} (replacement already done)`;
  for (const j of db.jobs.filter((j) => j.stage !== 'closed' && (j.assetIds || []).some((id) => ids.has(id)))) {
    const mine = j.assetIds.filter((id) => ids.has(id)); mine.forEach((id) => handled.add(id));
    if (mine.length === j.assetIds.length) { removeAssets(mine, j.ref); forceClose(j, by, note); rep.completed.jobsClosed++; }
    else {
      const done = makeJob({ type: j.type, outletCode: j.outletCode, priority: j.priority, newModel: j.newModel, newQty: j.type === 'DU_REPLACEMENT' ? mine.length : 0, oldQty: mine.length, oldModel: j.oldModel, vendorId: j.vendorId, assetIds: mine, planSource: 'COMPLETED LIST', reason: j.reason, remarks: serialsNote(mine) }, by);
      removeAssets(mine, done.ref); forceClose(done, by, note); rep.completed.jobsCreated++;
      j.assetIds = j.assetIds.filter((id) => !ids.has(id)); j.oldQty = j.assetIds.length; if (j.type === 'DU_REPLACEMENT') j.newQty = j.assetIds.length; j.remarks = serialsNote(j.assetIds);
      log(j, `${mine.length} DU(s) at this outlet were already replaced - recorded as ${done.ref}; ${j.assetIds.length} remaining`, by);
    }
  }
  const rest = assets.filter((a) => !handled.has(a.id));
  if (rest.length) {
    if (!db.outlets.find((o) => o.code === sap)) { rep.notes.push(`${src}: outlet ${sap} is not in the outlet list - ${rest.length} DU(s) marked removed but no job created`); removeAssets(rest.map((a) => a.id), 'completed list'); return; }
    const done = makeJob({ type: 'DU_REPLACEMENT', outletCode: sap, newModel: `Pressure DU ${mostCommon(rest.map((a) => a.nozzles)) || 2}-nozzle`, newQty: rest.length, oldQty: rest.length, oldModel: 'Suction DU (removed)', assetIds: rest.map((a) => a.id), planSource: 'COMPLETED LIST', reason: 'Suction DU replacement (completed)', remarks: serialsNote(rest.map((a) => a.id)) }, by);
    removeAssets(done.assetIds, done.ref); forceClose(done, by, note); rep.completed.jobsCreated++;
  }
}
function applyImport(list, by) {
  const rep = { files: [], outlets: { added: 0, updated: 0 }, assets: { added: 0, updated: 0, unmapped: 0 }, jobs: { created: 0, skipped: 0 }, completed: { dus: 0, jobsClosed: 0, jobsCreated: 0, already: 0 }, notes: [] };
  const order = { dealership: 0, gvr: 1, tokheim: 2, plan: 3, completed: 4 }, coords = {};
  list.sort((a, b) => (order[a.kind] ?? 9) - (order[b.kind] ?? 9));
  const outletBy = (c) => db.outlets.find((o) => o.code === c);
  for (const p of list) {
    if (!p.kind) { rep.files.push({ name: p.name, error: p.error }); continue; }
    rep.files.push({ name: p.name, kind: p.kind });
    (p.notes || []).slice(0, 12).forEach((n) => rep.notes.push(`${p.name}: ${n}`));
    if (p.kind === 'dealership') {
      for (const o of p.outlets) {
        const { lat, lng, ...rest } = o; let e = outletBy(o.code);
        if (e) { Object.assign(e, rest); rep.outlets.updated++; } else { e = { driveway: '', phase: '', soName: '', soEmail: '', soPhone: '', rmName: '', rmEmail: '', ...rest }; db.outlets.push(e); rep.outlets.added++; }
        if (lat != null) { e.lat = lat; e.lng = lng; } else if (e.lat === undefined) { e.lat = null; e.lng = null; }
      }
    }
    if (p.kind === 'gvr') {
      Object.assign(coords, p.coords); upsertAssets(p.assets, rep);
      const unknown = [...new Set(p.assets.map((a) => a.sap))].filter((c) => !outletBy(c));
      if (unknown.length) rep.notes.push(`${p.name}: ${unknown.length} SAP codes in the DU/STP register are not in the dealership list (e.g. ${unknown.slice(0, 4).join(', ')})`);
    }
    if (p.kind === 'tokheim') {
      let miss = 0;
      for (const a of p.assets) { const o = matchOutlet(a.dealerName, a.district, db.outlets); a.sap = o ? o.code : ''; if (!o) miss++; }
      upsertAssets(p.assets, rep); rep.assets.unmapped += miss;
      if (miss) rep.notes.push(`${p.name}: ${miss} of ${p.assets.length} Tokheim DUs could not be matched to an outlet by name (shown as "Unmapped" in Installed DUs; you can assign them)`);
    }
    if (p.kind === 'completed') {
      const bySap = new Map();
      for (const a of p.assets) {
        let e = db.assets.find((x) => x.kind === 'DU' && x.serial === a.serial);
        if (!e) { e = { id: uid('a'), status: 'Active', ...a }; db.assets.push(e); rep.assets.added++; }
        if (e.status === 'Removed') { rep.completed.already++; continue; }
        e.pumpType ||= 'Suction'; rep.completed.dus++; if (!bySap.has(e.sap)) bySap.set(e.sap, []); bySap.get(e.sap).push(e);
      }
      for (const [sap, as] of bySap) completeAssets(sap, as, by, rep, p.name);
    }
    if (p.kind === 'plan') {
      for (const o of db.outlets) if (o.lat == null && coords[o.code]) Object.assign(o, coords[o.code]);
      const groups = new Map(); for (const r of p.plan) { if (!groups.has(r.sap)) groups.set(r.sap, []); groups.get(r.sap).push(r); }
      const inJob = new Set(db.jobs.filter((j) => j.stage !== 'closed').flatMap((j) => j.assetIds || []));
      for (const [sap, rows] of groups) {
        const type = rows.some((r) => r.work === 'DU REMOVAL') ? 'RATIONALIZATION' : 'DU_REPLACEMENT';
        if (db.jobs.some((j) => j.planSource && j.outletCode === sap && j.type === type)) { rep.jobs.skipped++; continue; }
        if (!outletBy(sap)) { // outlet missing from dealership list: create from plan + GVR data
          db.outlets.push({ code: sap, name: rows[0].ro, district: '', salesArea: rows[0].area, state: 'Madhya Pradesh', region: 'Jabalpur Retail Region', zone: 'Bhopal Zone', lat: coords[sap]?.lat ?? null, lng: coords[sap]?.lng ?? null, driveway: '', phase: '', soName: '', soEmail: '', soPhone: '', rmName: '', rmEmail: '', stub: true });
          rep.notes.push(`${rows[0].ro} (${sap}) is not in the dealership list - added from the plan; please complete its details`); rep.outlets.added++;
        }
        const qty = rows.length, seq = Math.min(...rows.map((r) => r.seq || 99)), avg = rows.find((r) => r.avg != null)?.avg ?? null;
        const vend = db.vendors.find((v) => v.code && v.code.toUpperCase() === (rows.find((r) => r.vendor)?.vendor || '#'));
        const assets = db.assets.filter((a) => a.sap === sap && a.kind === 'DU' && a.pumpType === 'Suction' && a.status !== 'Removed' && !inJob.has(a.id)).sort((a, b) => (a.year || 9999) - (b.year || 9999)).slice(0, qty);
        assets.forEach((a) => inJob.add(a.id));
        const noz = mostCommon(assets.map((a) => a.nozzles)) || 2, date = rows.find((r) => r.date)?.date || null;
        const isDu = type === 'DU_REPLACEMENT';
        const exJob = type === 'DU_REPLACEMENT' && openDuJob(sap);
        if (exJob) { mergeIntoJob(exJob, { assetIds: assets.map((a) => a.id), oldQty: qty }, by); rep.notes.push(`${sap}: plan DUs added to existing job ${exJob.ref} (one job per outlet)`); continue; }
        try {
          const j = makeJob({ type, outletCode: sap, due: date, priority: seq <= 2 ? 'High' : 'Medium', newModel: isDu ? `Pressure DU ${noz}-nozzle` : '', newQty: isDu ? qty : 0, oldQty: qty, oldModel: 'Suction DU (removed)',
            reason: isDu ? 'Suction DU replacement as per DU Replacement Plan' : `DU removal as per plan${avg != null ? ` (6-month average sales ${avg.toFixed(1)})` : ''}`, remarks: assets.length ? 'Old DU serials: ' + assets.map((a) => a.serial).join(', ') : '',
            vendorId: vend?.id || null, planDate: date, seq: seq === 99 ? null : seq, avgSales: avg, assetIds: assets.map((a) => a.id), planSource: 'DU REPLACEMENT PLAN' }, by);
          rep.jobs.created++; if (!assets.length) rep.notes.push(`${j.ref} (${sap}): no suction DU found in the DU register for this outlet`);
        } catch (e) { rep.notes.push(`${sap}: ${e.message}`); }
      }
    }
  }
  for (const o of db.outlets) { const n = db.assets.filter((a) => a.sap === o.code && a.kind === 'DU' && a.status !== 'Removed').length; if (n) o.dus = n; }
  return rep;
}
function cli() {
  const files = process.argv.slice(process.argv.indexOf('--import') + 1).filter((a) => !a.startsWith('--'));
  if (process.argv.includes('--reset')) resetData();
  const rep = applyImport(files.map((f) => parseWorkbook(fs.readFileSync(f), path.basename(f))), 'import'); save();
  console.log(JSON.stringify(rep, null, 1)); process.exit(0);
}
route('POST', '/api/import/excel', true, ({ user, body }) => {
  const parsed = (body.files || []).map((f) => { try { return parseWorkbook(Buffer.from(String(f.data || ''), 'base64'), f.name); } catch (e) { return { name: f.name, kind: null, error: 'Could not read file: ' + e.message }; } });
  const rep = applyImport(parsed, user.name); save(); return rep;
});
route('GET', '/api/assets', true, () => db.assets);
route('PATCH', '/api/assets/:id', true, ({ user, params, body }) => {
  const a = db.assets.find((a) => a.id === params.id); if (!a) throw bad('Not found', 404);
  if ('serial' in body) { const sn = String(body.serial || '').trim(); if (!sn) throw bad('Serial number is required'); if (db.assets.some((x) => x !== a && x.kind === (body.kind || a.kind) && x.serial === sn)) throw bad('Another unit already has serial ' + sn); a.serial = sn; }
  if (body.sap && !db.outlets.find((o) => o.code === body.sap)) throw bad('Outlet not found');
  Object.assign(a, pick(body, ['sap', 'pumpType', 'model', 'make', 'category', 'contract', 'hp', 'kind']));
  if ('nozzles' in body) a.nozzles = +body.nozzles || null; if ('year' in body) a.year = +body.year || null;
  if (body.status === 'Removed' && a.status !== 'Removed') { a.status = 'Removed'; a.removedOn = today(); a.removedBy = a.removedBy || 'edited by admin'; scrapUnit(a, user.name, 'marked removed'); }
  if (body.status === 'Active' && a.status === 'Removed') { a.status = 'Active'; delete a.removedOn; delete a.removedBy; unscrapUnit(a); }
  save(); return a;
});
route('DELETE', '/api/assets/:id', true, ({ user, params, query }) => {
  const a = db.assets.find((x) => x.id === params.id); if (!a) throw bad('Not found', 404);
  const scrapped = query.scrap !== '0' && scrapUnit(a, user.name, 'deleted from register');
  db.assets = db.assets.filter((x) => x !== a); save(); return { ok: true, scrapped };
});
route('POST', '/api/assets/plan-suction', true, ({ user, body }) => {
  const inJob = new Set(db.jobs.filter((j) => j.stage !== 'closed').flatMap((j) => j.assetIds || [])), bySap = {};
  const maxYear = +body.maxYear || 9999;
  for (const a of db.assets) if (a.kind === 'DU' && a.pumpType === 'Suction' && a.status !== 'Removed' && a.sap && !inJob.has(a.id) && (a.year || 0) <= maxYear) (bySap[a.sap] ||= []).push(a);
  let created = 0, skipped = 0, merged = 0;
  for (const [sap, as] of Object.entries(bySap)) {
    if (!db.outlets.find((o) => o.code === sap)) { skipped++; continue; }
    const ex = openDuJob(sap); if (ex) { mergeIntoJob(ex, { assetIds: as.map((a) => a.id) }, user.name); merged++; continue; }
    const noz = mostCommon(as.map((a) => a.nozzles)) || 2;
    makeJob({ type: 'DU_REPLACEMENT', outletCode: sap, priority: 'Low', newModel: `Pressure DU ${noz}-nozzle`, newQty: as.length, oldQty: as.length, oldModel: 'Suction DU (removed)', reason: 'Suction DU replacement (programme)', remarks: 'Old DU serials: ' + as.map((a) => a.serial).join(', '), assetIds: as.map((a) => a.id), planSource: 'SUCTION REGISTER' }, user.name); created++;
  }
  save(); return { created, skippedNoOutlet: skipped, merged };
});
route('POST', '/api/sales-areas', true, ({ body }) => {
  if (!body.area) throw bad('Sales area required'); db.settings.salesAreas[body.area] = { name: body.name || '', email: body.email || '', phone: body.phone || '' }; save(); return { ok: true };
});

/* ---------- scrap register ---------- */
const scrapFor = (id) => { const x = db.scrap.find((x) => x.id === id); if (!x) throw bad('Scrap item not found', 404); return x; };
const canScrap = (user, x) => user.role !== 'vendor' || x.vendorId === user.vendorId;
route('POST', '/api/scrap', false, ({ user, body }) => {
  if (!String(body.name || '').trim()) throw bad('Item name required');
  if (!db.outlets.find((o) => o.code === body.outletCode)) throw bad('Select the outlet where the item is');
  if (user.role !== 'admin' && !db.jobs.some((j) => j.outletCode === body.outletCode && vendorOf(j) === user.vendorId)) throw bad('You can only add scrap for outlets where you have a job', 403);
  const x = { id: uid('sc'), name: body.name.trim(), category: body.category || 'Other', qty: +body.qty || 1, serial: body.serial || '', outletCode: body.outletCode, note: body.note || '', status: 'At outlet', jobId: body.jobId || null, vendorId: user.vendorId || null, files: [], at: new Date().toISOString(), by: user.name };
  db.scrap.push(x); save(); return x;
});
route('PATCH', '/api/scrap/:id', true, ({ params, body }) => {
  const x = scrapFor(params.id), was = x.status; Object.assign(x, pick(body, ['name', 'category', 'qty', 'serial', 'note', 'outletCode', 'status']));
  if (x.status !== 'At outlet') x.statusOn = today();
  if (x.status !== was) for (const c of db.scrap.filter((y) => y.parentId === x.id && y.status === was)) { c.status = x.status; c.statusOn = x.statusOn; }
  save(); return x;
});
route('DELETE', '/api/scrap/:id', true, ({ params }) => {
  const x = scrapFor(params.id); fs.rmSync(path.join(UPLOADS, 'scrap', x.id), { recursive: true, force: true }); db.scrap = db.scrap.filter((y) => y !== x); save(); return { ok: true };
});
route('POST', '/api/scrap/:id/files', false, ({ user, params, body }) => {
  const x = scrapFor(params.id); if (!canScrap(user, x)) throw bad('Not allowed', 403);
  const mime = String(body.mime || ''); if (!/^image\/(jpeg|png|webp)$/.test(mime)) throw bad('Only JPG, PNG or WEBP photos');
  const buf = Buffer.from(String(body.data || ''), 'base64'); if (!buf.length || buf.length > 9e6) throw bad('File empty or too large (max 9 MB)');
  const id = uid('f'), ext = mime === 'image/png' ? '.png' : mime === 'image/webp' ? '.webp' : '.jpg';
  fs.mkdirSync(path.join(UPLOADS, 'scrap', x.id), { recursive: true }); fs.writeFileSync(path.join(UPLOADS, 'scrap', x.id, id + ext), buf);
  const f = { id, ext, mime, name: String(body.name || 'photo').slice(0, 80), at: new Date().toISOString(), by: user.name }; x.files.push(f); save(); return f;
});
route('DELETE', '/api/scrap/:id/files/:fid', false, ({ user, params }) => {
  const x = scrapFor(params.id); if (!canScrap(user, x)) throw bad('Not allowed', 403);
  const f = x.files.find((f) => f.id === params.fid); if (!f) throw bad('Not found', 404);
  x.files = x.files.filter((y) => y !== f); fs.rmSync(path.join(UPLOADS, 'scrap', x.id, f.id + f.ext), { force: true }); save(); return { ok: true };
});
route('GET', '/api/scrapfiles/:sid/:fid', false, ({ user, params }) => {
  const x = scrapFor(params.sid); if (!canScrap(user, x) && !db.jobs.some((j) => j.outletCode === x.outletCode && vendorOf(j) === user.vendorId)) throw bad('Not allowed', 403);
  const f = x.files.find((f) => f.id === params.fid); if (!f) throw bad('Not found', 404);
  return { __file: path.join(UPLOADS, 'scrap', x.id, f.id + f.ext), mime: f.mime, name: f.name };
});

/* ---------- defaults / migration ---------- */
const HH_DEFAULT = [['ms_hose', 'MS hosepipe'], ['hsd_hose', 'HSD hosepipe'], ['ms_swivel', 'MS swivel joint'], ['hsd_swivel', 'HSD swivel joint'], ['ms_breakaway', 'MS breakaway coupling'], ['hsd_breakaway', 'HSD breakaway coupling'], ['ms_spout', 'MS spout'], ['hsd_spout', 'HSD spout'], ['reducer', 'Reducer'], ['sample_seal', 'Sample seal']];
function ensureDefaults() {
  const s = db.settings;
  const d = {
    rm: { name: 'Regional Manager', email: 'nimishpsingh@hpcl.in' }, dashOrder: DASH_ORDER,
    makes: ['GVR', 'Tokheim', 'Tatsuno', 'Victor'], nozzleOptions: [1, 2, 4, 6, 8], hpOptions: ['0.75', '1.5'], sessionHours: 12, jobDefaults: { dueDays: 30, priority: 'Medium', type: 'DU_REPLACEMENT' }, vendorMenu: { billing: true, scrap: true, map: true }, billNeedsBoth: true,
    areaColors: { rewa: '#f2c200', sagar: '#1e6bd6', jabalpur: '#f57c00', balaghat: '#2e9e4f', chhindwara: '#8d5524', katni: '#7b3fb8', satna: '#ec4899' },
    hhItems: HH_DEFAULT.map(([key, label]) => ({ key, label, min: 0 })), photoStages: [...PHOTO_STAGES], photoOverrides: {}, permitRequired: 'warn', ackRequired: true, weekOff: [0], refreshSec: 45,
    brand: { title: 'Retail Engineering', subtitle: 'Project Tracker', org: 'HPCL' }, theme: { blue: '#0b4ea2', red: '#e4002b' }, legacyYears: [2011, 2013],
    dash: { map: true, pipeline: true, upcoming: true, legacy: true, areas: true, hardware: true },
    mail: { deploySubject: '[{ref}] Team deployed - {work} at {outlet} ({date})', completeSubject: '[{ref}] Completed - {work} at {outlet}',
      deployIntro: 'This is to inform you that a team has been deployed for the following work at your retail outlet. Kindly coordinate with the dealer and ensure the outlet is ready for the activity.',
      completeIntro: 'The following work has been reported as commissioned/completed at your retail outlet.' },
  };
  for (const k of Object.keys(d)) if (s[k] === undefined) s[k] = d[k];
  if (!s.dashMig1) { s.dash = { ...(s.dash || {}), overdue: false }; s.dashTiles = { ...(s.dashTiles || {}), active: false, done: false, overdue: false }; s.dashMig1 = true; }
  db.hhStock ||= {}; db.hhIssues ||= []; db.hhMoves ||= []; db.reuses ||= []; db.scrapReports ||= [];
  if (JSON.stringify(s.stores) === JSON.stringify(['Main Store']) && !db.stock.length) s.stores = ['IRD Jabalpur', 'Teori'];
  for (const t of db.teams) { // a team = two members (no team lead)
    t.members = (t.members || []).map((m) => (typeof m === 'string' ? { name: /^Team lead/i.test(m) ? '' : m, phone: '' } : m));
    while (t.members.length < 2) t.members.push({ name: '', phone: '' });
    t.members = t.members.slice(0, 2); delete t.lead; delete t.phone;
  }
  db.bills ||= [];
  for (const j of db.jobs) {
    j.permit ||= { status: 'Not issued' };
    j.cablingStatus ||= j.cabling === 'Required' ? 'Incomplete' : j.cabling === 'Not required' ? 'Not required' : 'N/A';
    j.automation ||= { status: j.type === 'DU_REPLACEMENT' ? 'Pending' : 'N/A' };
  }
}
const hhState = () => ({ items: db.settings.hhItems, stock: db.hhStock, issues: db.hhIssues.slice(-300).reverse(), moves: db.hhMoves.slice(-300).reverse() });

/* ---------- work permit ---------- */
route('POST', '/api/jobs/:id/permit', true, ({ user, params, body }) => {
  const j = jobFor(params.id), st = body.status === 'Issued' ? 'Issued' : 'Not issued';
  j.permit = { status: st, date: st === 'Issued' ? body.date || today() : '', ref: body.ref || '', type: ['Issued by officer', 'Issued by dealer', 'Other'].includes(body.type) ? body.type : '', typeText: body.type === 'Other' ? String(body.typeText || '').slice(0, 80) : '', by: user.name };
  log(j, `Work permit ${st.toLowerCase()}${body.ref ? ' (' + body.ref + ')' : ''}`, user.name); save(); return j;
});

/* ---------- hanging hardware ---------- */
const hhItem = (k) => db.settings.hhItems.find((i) => i.key === k);
const HH_DIR = path.join(__dirname, 'data', 'uploads', 'hh');
route('POST', '/api/hh/receive', true, ({ user, body }) => {
  const qty = Math.floor(+body.qty); if (!hhItem(body.key) || !(qty > 0)) throw bad('Choose an item and a quantity');
  db.hhStock[body.key] = (db.hhStock[body.key] || 0) + qty;
  db.hhMoves.push({ id: uid('hm'), at: new Date().toISOString(), date: body.date || today(), type: 'RECEIVE', key: body.key, qty, ref: body.ref || 'Receipt', by: user.name }); save(); return { ok: true };
});
route('POST', '/api/hh/adjust', true, ({ user, body }) => {
  const d = Math.trunc(+body.delta); if (!hhItem(body.key) || !d) throw bad('Choose an item and a non-zero change');
  if ((db.hhStock[body.key] || 0) + d < 0) throw bad('Stock cannot go below zero');
  db.hhStock[body.key] = (db.hhStock[body.key] || 0) + d;
  db.hhMoves.push({ id: uid('hm'), at: new Date().toISOString(), date: today(), type: 'ADJUST', key: body.key, qty: d, ref: body.reason || 'Manual adjustment', by: user.name }); save(); return { ok: true };
});
function saveHhFile(issueId, f, role) {
  const mime = String(f.mime || ''); if (!/^(image\/(jpeg|png|webp)|application\/pdf)$/.test(mime)) throw bad('Attachments must be JPG, PNG, WEBP or PDF');
  const buf = Buffer.from(String(f.data || ''), 'base64'); if (!buf.length || buf.length > 9e6) throw bad('Attachment empty or larger than 9 MB');
  const id = uid('f'), ext = mime === 'application/pdf' ? '.pdf' : mime === 'image/png' ? '.png' : mime === 'image/webp' ? '.webp' : '.jpg';
  fs.mkdirSync(path.join(HH_DIR, issueId), { recursive: true }); fs.writeFileSync(path.join(HH_DIR, issueId, id + ext), buf);
  return { id, ext, mime, name: String(f.name || 'file').slice(0, 80), role };
}
route('POST', '/api/hh/issue', true, ({ user, body }) => {
  const lines = (body.lines || []).map((l) => ({ outletCode: String(l.outletCode || ''), key: l.key, qty: Math.floor(+l.qty) })).filter((l) => l.qty > 0);
  if (!lines.length) throw bad('Add at least one outlet / item / quantity line');
  for (const l of lines) { if (!db.outlets.find((o) => o.code === l.outletCode)) throw bad(`Outlet ${l.outletCode} not found`); if (!hhItem(l.key)) throw bad('Unknown item'); }
  const need = {}; lines.forEach((l) => (need[l.key] = (need[l.key] || 0) + l.qty));
  for (const [k, q] of Object.entries(need)) if ((db.hhStock[k] || 0) < q) throw bad(`Only ${db.hhStock[k] || 0} x ${hhItem(k).label} in stock, ${q} requested`);
  if (!(body.fse || []).length) throw bad('Attach the Field Service Engineer report');
  if (!(body.dealer || []).length) throw bad("Attach the dealer's letter");
  db.seq.HH = (db.seq.HH || 0) + 1; const id = uid('hi'), no = 'HH-' + String(db.seq.HH).padStart(4, '0');
  let fse, dealer; try { fse = body.fse.map((f) => saveHhFile(id, f, 'fse')); dealer = body.dealer.map((f) => saveHhFile(id, f, 'dealer')); } catch (e) { fs.rmSync(path.join(HH_DIR, id), { recursive: true, force: true }); db.seq.HH--; throw e; }
  for (const l of lines) { db.hhStock[l.key] -= l.qty; db.hhMoves.push({ id: uid('hm'), at: new Date().toISOString(), date: body.date || today(), type: 'ISSUE', key: l.key, qty: -l.qty, ref: `${no} → ${l.outletCode}`, by: user.name }); }
  const issue = { id, no, date: body.date || today(), lines, fse, dealer, fseRef: body.fseRef || '', dealerRef: body.dealerRef || '', note: body.note || '', by: user.name, at: new Date().toISOString(), cancelled: false };
  db.hhIssues.push(issue); save(); return issue;
});
route('POST', '/api/hh/issues/:id/cancel', true, ({ user, params }) => {
  const i = db.hhIssues.find((x) => x.id === params.id); if (!i) throw bad('Not found', 404); if (i.cancelled) throw bad('Already cancelled');
  for (const l of i.lines) { db.hhStock[l.key] = (db.hhStock[l.key] || 0) + l.qty; db.hhMoves.push({ id: uid('hm'), at: new Date().toISOString(), date: today(), type: 'CANCEL', key: l.key, qty: l.qty, ref: `${i.no} cancelled`, by: user.name }); }
  i.cancelled = true; i.cancelledBy = user.name; save(); return i;
});
route('GET', '/api/hhfiles/:iid/:fid', true, ({ params }) => {
  const i = db.hhIssues.find((x) => x.id === params.iid); const f = i && [...i.fse, ...i.dealer].find((f) => f.id === params.fid); if (!f) throw bad('Not found', 404);
  return { __file: path.join(HH_DIR, i.id, f.id + f.ext), mime: f.mime, name: f.name };
});

/* ---------- logo ---------- */
route('POST', '/api/logo', true, ({ body }) => {
  const mime = String(body.mime || ''); if (!/^image\/(png|jpeg|webp)$/.test(mime)) throw bad('Logo must be PNG, JPG or WEBP');
  const buf = Buffer.from(String(body.data || ''), 'base64'); if (!buf.length || buf.length > 4e6) throw bad('Logo empty or larger than 4 MB');
  for (const e of ['png', 'jpg', 'webp']) fs.rmSync(path.join(__dirname, 'data', 'logo.' + e), { force: true });
  fs.writeFileSync(path.join(__dirname, 'data', 'logo.' + (mime === 'image/png' ? 'png' : mime === 'image/webp' ? 'webp' : 'jpg')), buf);
  db.settings.logoV = Date.now(); save(); return { ok: true };
});
route('POST', '/api/logo/reset', true, () => { for (const e of ['png', 'jpg', 'webp']) fs.rmSync(path.join(__dirname, 'data', 'logo.' + e), { force: true }); db.settings.logoV = Date.now(); save(); return { ok: true }; });

/* ---------- automation integration, cabling status ---------- */
const DASH_ORDER = ['map', 'legacy', 'upcoming', 'setup', 'areas', 'pipeline', 'automation', 'cabling', 'permits', 'billing', 'inventory', 'hardware', 'scrap', 'vendors', 'activity', 'overdue', 'estimates'];
const AUTOMATION_STATUS = ['Pending', 'In progress', 'Integrated', 'Issue / not working', 'Not applicable'];
route('POST', '/api/jobs/:id/automation', false, ({ user, params, body }) => {
  const j = jobFor(params.id); if (!canSee(user, j)) throw bad('Not your job', 403);
  if (j.type !== 'DU_REPLACEMENT') throw bad('Automation integration applies to DU replacement jobs');
  if (!AUTOMATION_STATUS.includes(body.status)) throw bad('Invalid status');
  j.automation = { status: body.status, date: body.date || today(), note: String(body.note || '').slice(0, 200), by: user.name };
  log(j, `Automation integration: ${body.status}${body.note ? ' - ' + body.note : ''}`, user.name); save(); return j;
});
route('POST', '/api/jobs/:id/cabling', false, ({ user, params, body }) => {
  const j = jobFor(params.id); if (!canSee(user, j)) throw bad('Not your job', 403);
  if (j.type !== 'DU_REPLACEMENT') throw bad('Cabling applies to DU replacement jobs');
  const admin = user.role === 'admin', ok = admin ? ['Completed', 'In progress', 'Incomplete', 'Not required'] : ['In progress', 'Incomplete'];
  if (!ok.includes(body.status)) throw bad('That status can only be set by HPCL or by completing the cabling step');
  if (body.status === 'Not required') { j.cabling = 'Not required'; j.cablingReason = 'Set manually'; }
  else if (j.cabling !== 'Required') throw bad('Automation cabling is not required at this outlet');
  j.cablingStatus = body.status; log(j, `Automation cabling: ${body.status}${body.note ? ' - ' + body.note : ''}`, user.name); save(); return j;
});

/* ---------- vendor bills ---------- */
const BILL_DIR = path.join(__dirname, 'data', 'uploads', 'bills');
const billFor = (id) => { const b = db.bills.find((x) => x.id === id); if (!b) throw bad('Bill not found', 404); return b; };
const canBill = (user, b) => user.role !== 'vendor' || b.vendorId === user.vendorId;
const billLocked = (b) => !!(b.reVerified || b.soVerified || b.releasedOn);
function cleanBill(b, body, user) {
  for (const k of ['callupNo', 'callupDate', 'billNo', 'billDate', 'description', 'remarks']) if (k in body) b[k] = String(body[k] ?? '').slice(0, 300);
  if ('amount' in body) { const a = Number(body.amount); if (!(a >= 0)) throw bad('Enter a valid amount'); b.amount = a; }
  if ('jobIds' in body) b.jobIds = (body.jobIds || []).filter((id) => { const j = db.jobs.find((x) => x.id === id); return j && (user.role === 'admin' || vendorOf(j) === b.vendorId); });
}
route('POST', '/api/bills', false, ({ user, body }) => {
  const vendorId = user.role === 'admin' ? body.vendorId : user.vendorId;
  if (!db.vendors.find((v) => v.id === vendorId)) throw bad('Choose the vendor');
  db.seq.BL = (db.seq.BL || 0) + 1;
  const b = { id: uid('bl'), no: 'BL-' + String(db.seq.BL).padStart(4, '0'), vendorId, jobIds: [], callupNo: '', callupDate: '', billNo: '', billDate: '', amount: 0, description: '', remarks: '', files: [], reVerified: null, soVerified: null, releasedOn: '', releasedAmount: null, createdBy: user.name, at: new Date().toISOString() };
  cleanBill(b, body, user); db.bills.push(b); save(); return b;
});
route('PATCH', '/api/bills/:id', false, ({ user, params, body }) => {
  const b = billFor(params.id); if (!canBill(user, b)) throw bad('Not your bill', 403);
  if (user.role !== 'admin' && billLocked(b)) throw bad('This bill is already under verification - ask HPCL to change it', 403);
  cleanBill(b, body, user); if (user.role === 'admin' && body.vendorId && db.vendors.find((v) => v.id === body.vendorId)) b.vendorId = body.vendorId; save(); return b;
});
route('POST', '/api/bills/:id/verify', true, ({ user, params, body }) => {
  const b = billFor(params.id), key = body.who === 'SO' ? 'soVerified' : body.who === 'RE' ? 'reVerified' : null; if (!key) throw bad('Say who verified (RE or SO)');
  if (b.releasedOn && !body.on) throw bad('Cancel the release first');
  if (key === 'reVerified' && body.on && !b.soVerified) throw bad('The Sales Officer must verify the bill before the RE team');
  if (key === 'soVerified' && !body.on && b.reVerified) throw bad('Remove the RE-team verification first');
  b[key] = body.on ? { on: body.date || today(), by: String(body.by || user.name).slice(0, 80), note: String(body.note || '').slice(0, 200) } : null; save(); return b;
});
route('POST', '/api/bills/:id/release', true, ({ params, body }) => {
  const b = billFor(params.id);
  if (!body.on) { b.releasedOn = ''; b.releasedAmount = null; save(); return b; }
  if (!body.force && db.settings.billNeedsBoth !== false && !(b.reVerified && b.soVerified)) throw bad('Both RE-team and SO verification are needed before release (tick "release anyway" to override)');
  b.releasedOn = body.date || today(); b.releasedAmount = body.amount !== '' && body.amount != null ? Number(body.amount) : b.amount; save(); return b;
});
route('DELETE', '/api/bills/:id', true, ({ params }) => { const b = billFor(params.id); fs.rmSync(path.join(BILL_DIR, b.id), { recursive: true, force: true }); db.bills = db.bills.filter((x) => x !== b); save(); return { ok: true }; });
route('POST', '/api/bills/:id/files', false, ({ user, params, body }) => {
  const b = billFor(params.id); if (!canBill(user, b)) throw bad('Not your bill', 403);
  if (user.role !== 'admin' && billLocked(b)) throw bad('This bill is already under verification', 403);
  const mime = String(body.mime || ''); if (!/^(image\/(jpeg|png|webp)|application\/pdf)$/.test(mime)) throw bad('Only JPG, PNG, WEBP or PDF');
  const buf = Buffer.from(String(body.data || ''), 'base64'); if (!buf.length || buf.length > 9e6) throw bad('File empty or larger than 9 MB');
  const id = uid('f'), ext = mime === 'application/pdf' ? '.pdf' : mime === 'image/png' ? '.png' : mime === 'image/webp' ? '.webp' : '.jpg';
  fs.mkdirSync(path.join(BILL_DIR, b.id), { recursive: true }); fs.writeFileSync(path.join(BILL_DIR, b.id, id + ext), buf);
  const f = { id, ext, mime, kind: body.kind === 'callup' ? 'callup' : 'bill', name: String(body.name || 'file').slice(0, 80), at: new Date().toISOString(), by: user.name }; b.files.push(f); save(); return f;
});
route('DELETE', '/api/bills/:id/files/:fid', false, ({ user, params }) => {
  const b = billFor(params.id); if (!canBill(user, b)) throw bad('Not your bill', 403); if (user.role !== 'admin' && billLocked(b)) throw bad('This bill is already under verification', 403);
  const f = b.files.find((x) => x.id === params.fid); if (!f) throw bad('Not found', 404);
  b.files = b.files.filter((x) => x !== f); fs.rmSync(path.join(BILL_DIR, b.id, f.id + f.ext), { force: true }); save(); return { ok: true };
});
route('GET', '/api/billfiles/:bid/:fid', false, ({ user, params }) => {
  const b = billFor(params.bid); if (!canBill(user, b)) throw bad('Not your bill', 403); const f = b.files.find((x) => x.id === params.fid); if (!f) throw bad('Not found', 404);
  return { __file: path.join(BILL_DIR, b.id, f.id + f.ext), mime: f.mime, name: f.name };
});

/* ---------- automatic scrap for replaced suction DUs (+ their pipeline) ---------- */
// Each removed suction DU becomes a "Scrap DUs" item lying at its outlet, and every outlet where a suction DU was
// replaced gets one "Pipeline & fittings" item. Flags (asset.scrapLogged, db.autoPipe) make this happen only once,
// so deleting or collecting the scrap later is never undone by a restart.
// Rule: a DU or STP taken out of the register (status Removed, or deleted) becomes scrap at its outlet.
// Suction DUs go through autoScrap so the outlet's old pipeline is added too.
function addDuParts(x, a, by) {
  if (!x || x.partsLogged || x.category !== 'Scrap DUs') return;
  const n = +(a && a.nozzles) || +((String(x.note || '').match(/(\d+)\s*nozzle/) || [])[1]) || 0, qty = n || 1;
  for (const [cat, word] of [['Nozzles', 'Nozzle'], ['Hosepipes', 'Hosepipe']]) {
    db.scrap.push({ id: uid('sc'), name: `${word}${qty > 1 ? 's' : ''} from DU ${x.serial}`, category: cat, qty, serial: x.serial, outletCode: x.outletCode,
      note: n ? `${n}-nozzle DU — one ${word.toLowerCase()} per nozzle` : 'Nozzle count not recorded for this DU — please check the quantity', status: x.status, statusOn: x.statusOn, jobId: null, vendorId: null, files: [], at: new Date().toISOString(), by, auto: 'du-part', assetId: x.assetId || null, parentId: x.id });
  }
  x.partsLogged = true;
  for (const r of db.recovered || []) if (r.outletCode === x.outletCode && REC_SCRAP_CAT[r.type]) linkRecovered(r);
}
/* Recovered (reusable) hosepipes / nozzles and the auto scrap of the same DU must never both count the same piece:
   recording a recovered part takes that quantity off the outlet's auto nozzle/hosepipe scrap (same DU first);
   deleting or reducing the recovered entry gives it back. r.linkQty = pieces recorded as recovered (reuse does not change it). */
const REC_SCRAP_CAT = { ms_hose: 'Hosepipes', hsd_hose: 'Hosepipes', nozzle: 'Nozzles' };
function linkRecovered(r) {
  const cat = REC_SCRAP_CAT[r.type]; if (!cat) return 0;
  r.fromScrap ||= []; r.linkQty ??= r.qty;
  let want = r.linkQty - r.fromScrap.reduce((a, l) => a + l.qty, 0), took = 0; if (want <= 0) return 0;
  const cands = db.scrap.filter((x) => x.auto === 'du-part' && x.category === cat && x.outletCode === r.outletCode && x.status === 'At outlet' && x.qty > 0).sort((a, b) => (b.serial === r.serial) - (a.serial === r.serial));
  for (const x of cands) {
    if (want <= 0) break; const t = Math.min(want, x.qty);
    x.qty -= t; x.toReusable = (x.toReusable || 0) + t; want -= t; took += t; r.fromScrap.push({ qty: t, item: { ...x, qty: 0 } });
    if (x.qty === 0) db.scrap = db.scrap.filter((y) => y !== x);
  }
  return took;
}
function unlinkRecovered(r, back = Infinity) {
  while (back > 0 && r.fromScrap?.length) {
    const l = r.fromScrap[r.fromScrap.length - 1], t = Math.min(back, l.qty); let x = db.scrap.find((y) => y.id === l.item.id);
    if (!x) { if (!db.scrap.some((y) => y.id === l.item.parentId)) { r.fromScrap.pop(); continue; } x = { ...l.item, qty: 0 }; db.scrap.push(x); } // DU itself no longer in scrap: nothing to give back
    x.qty += t; x.toReusable = Math.max(0, (x.toReusable || 0) - t); l.qty -= t; back -= t; if (l.qty === 0) r.fromScrap.pop();
  }
}
function backfillDuParts() { let n = 0; for (const x of db.scrap.filter((s) => s.auto === 'du' && s.category === 'Scrap DUs' && !s.partsLogged)) { addDuParts(x, db.assets.find((a) => a.id === x.assetId), 'system'); n++; } return n; }
function scrapUnit(a, by, ref) {
  if (!a || a.scrapLogged || !db.outlets.find((o) => o.code === a.sap)) return false;
  if (a.kind === 'DU' && a.pumpType === 'Suction') return autoScrap([{ ...a, status: 'Removed' }], by, ref) > 0 && (a.scrapLogged = true);
  db.scrap.push({ id: uid('sc'), name: `${a.kind === 'DU' ? (a.pumpType || '') + ' DU' : 'STP'} ${a.serial}${a.make ? ' · ' + a.make : ''}${a.model ? ' ' + a.model : ''}`.trim(), category: a.kind === 'DU' ? 'Scrap DUs' : 'Other items', qty: 1, serial: a.serial, outletCode: a.sap,
    note: `Removed from the installed register${a.nozzles ? ' · ' + a.nozzles + ' nozzle' : ''}${a.hp ? ' · ' + a.hp + ' HP' : ''}${a.year ? ' · ' + a.year : ''}${ref ? ' · ' + ref : ''}`, status: 'At outlet', jobId: null, vendorId: null, files: [], at: new Date().toISOString(), by, auto: 'du', assetId: a.id });
  addDuParts(db.scrap[db.scrap.length - 1], a, by); a.scrapLogged = true; return true;
}
function unscrapUnit(a) { // put back to "installed": take its scrap entry off again (only if it is still lying at the outlet)
  const before = db.scrap.length;
  db.scrap = db.scrap.filter((x) => !((x.auto === 'du' || x.auto === 'du-part') && x.assetId === a.id && x.status === 'At outlet'));
  if (db.scrap.length !== before) delete a.scrapLogged;
  if (!db.scrap.some((x) => x.auto === 'du' && x.outletCode === a.sap)) { db.scrap = db.scrap.filter((x) => !(x.auto === 'pipeline' && x.outletCode === a.sap && x.status === 'At outlet')); db.autoPipe = (db.autoPipe || []).filter((c) => c !== a.sap); }
}
function autoScrap(assets, by, ref) {
  let n = 0; db.autoPipe ||= [];
  const fresh = assets.filter((a) => a && a.kind === 'DU' && a.pumpType === 'Suction' && a.status === 'Removed' && !a.scrapLogged && db.outlets.find((o) => o.code === a.sap));
  for (const a of fresh) {
    db.scrap.push({ id: uid('sc'), name: `Suction DU ${a.serial}${a.make ? ' · ' + a.make : ''}${a.model ? ' ' + a.model : ''}`, category: 'Scrap DUs', qty: 1, serial: a.serial, outletCode: a.sap, note: `Replaced suction DU${a.nozzles ? ' (' + a.nozzles + ' nozzle)' : ''}${a.year ? ', ' + a.year : ''}${ref ? ' · ' + ref : ''}`, status: 'At outlet', jobId: null, vendorId: null, files: [], at: new Date().toISOString(), by, auto: 'du', assetId: a.id });
    addDuParts(db.scrap[db.scrap.length - 1], a, by); a.scrapLogged = true; n++;
  }
  for (const sap of new Set(fresh.map((a) => a.sap))) {
    if (db.autoPipe.includes(sap)) continue;
    db.scrap.push({ id: uid('sc'), name: 'Suction pipeline & fittings', category: 'Pipeline & fittings', qty: 1, serial: '', outletCode: sap, note: 'Old suction pipeline left after suction DU replacement (1 lot)', status: 'At outlet', jobId: null, vendorId: null, files: [], at: new Date().toISOString(), by, auto: 'pipeline' });
    db.autoPipe.push(sap); n++;
  }
  return n;
}
function undoAutoScrap(job) { // a job stepped back before "dismantled": take its auto scrap off the map again
  const ids = new Set(job.assetIds || []), saps = new Set();
  db.scrap = db.scrap.filter((x) => { const hit = (x.auto === 'du' || x.auto === 'du-part') && ids.has(x.assetId) && x.status === 'At outlet'; if (hit) { saps.add(x.outletCode); const a = db.assets.find((y) => y.id === x.assetId); if (a) delete a.scrapLogged; } return !hit; });
  for (const sap of saps) if (!db.scrap.some((x) => x.auto === 'du' && x.outletCode === sap)) { db.scrap = db.scrap.filter((x) => !(x.auto === 'pipeline' && x.outletCode === sap && x.status === 'At outlet')); db.autoPipe = (db.autoPipe || []).filter((c) => c !== sap); }
}

/* ---------- reusable parts recovered from old DUs ---------- */
const REC_TYPES = [['ms_hose', 'MS hosepipe', 'Hosepipes'], ['hsd_hose', 'HSD hosepipe', 'Hosepipes'], ['swivel', 'Swivel joint', 'Pipeline & fittings'], ['metering', 'Metering unit', 'Other items'], ['nozzle', 'Nozzle', 'Nozzles'], ['breakaway', 'Breakaway coupling', 'Hosepipes'], ['other', 'Other part', 'Other items']];
const recType = (k) => REC_TYPES.find((t) => t[0] === k);
const recFor = (id) => { const r = db.recovered.find((x) => x.id === id); if (!r) throw bad('Item not found', 404); return r; };
function recToScrap(r, by) {
  if (r.scrapId && db.scrap.find((x) => x.id === r.scrapId)) return;
  const t = recType(r.type), x = { id: uid('sc'), name: `${r.label || t[1]} (from old DU${r.serial ? ' ' + r.serial : ''})`, category: t[2], qty: r.qty, serial: r.serial || '', outletCode: r.outletCode, note: 'Recovered DU part marked as scrap' + (r.note ? ' · ' + r.note : ''), status: 'At outlet', jobId: null, vendorId: null, files: [], at: new Date().toISOString(), by, auto: 'recovered', recoveredId: r.id };
  db.scrap.push(x); r.scrapId = x.id;
}
route('POST', '/api/recovered', true, ({ user, body }) => {
  if (!db.outlets.find((o) => o.code === body.outletCode)) throw bad('Choose the outlet the parts came from');
  const items = (body.items || []).map((i) => ({ type: i.type, qty: Math.floor(+i.qty) })).filter((i) => recType(i.type) && i.qty > 0);
  if (!items.length) throw bad('Enter a quantity for at least one part');
  const cond = body.cond === 'Scrap' ? 'Scrap' : 'Reusable', made = [];
  for (const i of items) {
    const r = { id: uid('rc'), type: i.type, label: i.type === 'other' && body.otherLabel ? String(body.otherLabel).slice(0, 60) : recType(i.type)[1], qty: i.qty, outletCode: body.outletCode, serial: String(body.serial || '').slice(0, 40), assetId: body.assetId || null, cond, store: String(body.store || '').slice(0, 60), note: String(body.note || '').slice(0, 200), at: new Date().toISOString(), by: user.name, scrapId: null };
    r.linkQty = r.qty; linkRecovered(r);
    if (cond === 'Scrap') recToScrap(r, user.name);
    db.recovered.push(r); made.push(r);
  }
  save(); return made;
});
route('PATCH', '/api/recovered/:id', true, ({ user, params, body }) => {
  const r = recFor(params.id);
  if ('qty' in body) {
    const q = Math.floor(+body.qty); if (!(q > 0)) throw bad('Quantity must be at least 1');
    const d = q - r.qty; r.qty = q; r.linkQty = Math.max(0, (r.linkQty ?? q - d) + d); const x = db.scrap.find((s) => s.id === r.scrapId); if (x) x.qty = q;
    if (d > 0) linkRecovered(r); else { const taken = (r.fromScrap || []).reduce((a, l) => a + l.qty, 0); if (taken > r.linkQty) unlinkRecovered(r, taken - r.linkQty); }
  }
  for (const k of ['store', 'note']) if (k in body) r[k] = String(body[k] || '').slice(0, k === 'note' ? 200 : 60);
  if (body.cond === 'Scrap' && r.cond !== 'Scrap') { r.cond = 'Scrap'; recToScrap(r, user.name); }
  if (body.cond === 'Reusable' && r.cond !== 'Reusable') { r.cond = 'Reusable'; db.scrap = db.scrap.filter((s) => s.id !== r.scrapId); r.scrapId = null; }
  save(); return r;
});
route('DELETE', '/api/recovered/:id', true, ({ params }) => {
  const r = recFor(params.id); unlinkRecovered(r); db.scrap = db.scrap.filter((s) => s.id !== r.scrapId); db.recovered = db.recovered.filter((x) => x !== r); save(); return { ok: true };
});

/* ---------- remove all jobs (admin; backup is written first) ---------- */
function clearJobs(by, type) {
  const dir = path.join(__dirname, 'data', 'backups'); fs.mkdirSync(dir, { recursive: true });
  const file = path.join(dir, `db-${new Date().toISOString().replace(/[:.]/g, '-')}-before-removing-jobs.json`);
  fs.writeFileSync(file, JSON.stringify(db));
  const gone = db.jobs.filter((j) => !type || j.type === type), ids = new Set(gone.map((j) => j.id)), n = gone.length;
  for (const j of gone) { release(j, by); fs.rmSync(path.join(__dirname, 'data', 'uploads', j.id), { recursive: true, force: true }); }
  for (const b of db.bills) b.jobIds = (b.jobIds || []).filter((id) => !ids.has(id));
  for (const x of db.scrap) if (ids.has(x.jobId)) x.jobId = null;
  db.jobs = db.jobs.filter((j) => !ids.has(j.id));
  if (!db.outbox.some((m) => ids.has(m.jobId))) { const p = { DU_REPLACEMENT: 'DU', RATIONALIZATION: 'RZ', STP_REPLACEMENT: 'STP' }; for (const [t, k] of Object.entries(p)) if ((!type || type === t) && !db.jobs.some((j) => j.type === t)) delete db.seq[k]; } // nothing refers to old job numbers, so start again at 0001
  return { removed: n, backup: path.basename(file) };
}
route('POST', '/api/admin/clear-jobs', true, ({ user, body }) => { const r = clearJobs(user.name, STAGES[body.type] ? body.type : null); save(); return r; });

/* ---------- reused parts: moved from one outlet to another ---------- */
route('POST', '/api/reuse', true, ({ user, body }) => {
  const qty = Math.floor(+body.qty); if (!(qty > 0)) throw bad('Enter the quantity reused');
  const to = String(body.toOutlet || ''); if (!db.outlets.find((o) => o.code === to)) throw bad('Choose the outlet where the part was fitted');
  let r = null, from = String(body.fromOutlet || ''), type = body.type;
  if (body.recoveredId) {
    r = recFor(body.recoveredId); if (r.cond !== 'Reusable') throw bad('Only items marked reusable can be reused');
    if (qty > r.qty) throw bad(`Only ${r.qty} available in this entry`); from = r.outletCode; type = r.type;
  }
  if (!recType(type)) throw bad('Choose the item'); if (!db.outlets.find((o) => o.code === from)) throw bad('Choose the outlet the part came from');
  if (from === to) throw bad('The "from" and "to" outlets must be different');
  const x = { id: uid('ru'), type, label: r ? r.label : recType(type)[1], qty, fromOutlet: from, toOutlet: to, recoveredId: r?.id || null, serial: r?.serial || String(body.serial || '').slice(0, 40), date: body.date || today(), jobRef: String(body.jobRef || '').slice(0, 20), note: String(body.note || '').slice(0, 200), by: user.name, at: new Date().toISOString() };
  if (r) { r.qty -= qty; if (r.qty === 0) r.cond = 'Reused'; }
  db.reuses.push(x); save(); return x;
});
route('DELETE', '/api/reuse/:id', true, ({ params }) => {
  const x = db.reuses.find((y) => y.id === params.id); if (!x) throw bad('Not found', 404);
  const r = x.recoveredId && db.recovered.find((y) => y.id === x.recoveredId); if (r) { r.qty += x.qty; if (r.cond === 'Reused') r.cond = 'Reusable'; }
  db.reuses = db.reuses.filter((y) => y !== x); save(); return { ok: true };
});

/* ---------- installed DU / STP register: add a unit ---------- */
route('POST', '/api/assets', true, ({ body }) => {
  const kind = body.kind === 'STP' ? 'STP' : 'DU', serial = String(body.serial || '').trim(); if (!serial) throw bad('Serial number is required');
  if (db.assets.some((a) => a.kind === kind && a.serial === serial)) throw bad(`A ${kind} with serial ${serial} already exists`);
  if (body.sap && !db.outlets.find((o) => o.code === body.sap)) throw bad('Outlet not found');
  const a = { id: uid('a'), kind, serial, sap: body.sap || '', make: String(body.make || ''), model: String(body.model || ''), category: String(body.category || ''), pumpType: kind === 'DU' ? (['Suction', 'Pressure'].includes(body.pumpType) ? body.pumpType : '') : '', nozzles: +body.nozzles || null, hp: kind === 'STP' ? String(body.hp || '') : '', year: +body.year || null, contract: String(body.contract || ''), source: 'MANUAL', status: 'Active' };
  db.assets.push(a); save(); return a;
});

/* ---------- scrap reports from Sales Officers ---------- */
const SREP_DIR = path.join(__dirname, 'data', 'uploads', 'scrap-reports');
const SREP_MIME = /^(application\/pdf|image\/(jpeg|png|webp)|text\/csv|application\/vnd\.ms-excel|application\/vnd\.openxmlformats-officedocument\.spreadsheetml\.sheet)$/;
// read a spreadsheet report and hand the rows back for review (nothing is registered here)
route('POST', '/api/scrap-reports/parse', true, ({ body }) => {
  const buf = Buffer.from(String(body.data || ''), 'base64'); if (!buf.length || buf.length > 9e6) throw bad('File empty or larger than 9 MB');
  let wb; try { wb = XLSX.read(buf, { type: 'buffer', cellDates: true }); } catch { throw bad('Could not read this spreadsheet'); }
  const sheets = wb.SheetNames.map((n) => ({ name: n, rows: XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, defval: '', raw: false }).slice(0, 2000) }));
  return { sheets };
});
route('POST', '/api/scrap-reports', true, ({ user, body }) => {
  const mime = String(body.mime || ''); if (!SREP_MIME.test(mime)) throw bad('Report must be a PDF, photo, Excel or CSV file');
  const buf = Buffer.from(String(body.data || ''), 'base64'); if (!buf.length || buf.length > 9e6) throw bad('File empty or larger than 9 MB');
  const id = uid('sr'), ext = (String(body.name || '').match(/\.[a-z0-9]{2,5}$/i) || ['.bin'])[0].toLowerCase();
  fs.mkdirSync(SREP_DIR, { recursive: true }); fs.writeFileSync(path.join(SREP_DIR, id + ext), buf);
  const r = { id, ext, mime, name: String(body.name || 'report').slice(0, 100), soName: String(body.soName || '').slice(0, 80), date: body.date || today(), note: String(body.note || '').slice(0, 200), itemIds: [], by: user.name, at: new Date().toISOString() };
  db.scrapReports.push(r); save(); return r;
});
route('GET', '/api/scrapreports/:id', true, ({ params }) => {
  const r = db.scrapReports.find((x) => x.id === params.id); if (!r) throw bad('Not found', 404);
  return { __file: path.join(SREP_DIR, r.id + r.ext), mime: r.mime, name: r.name };
});
route('DELETE', '/api/scrap-reports/:id', true, ({ params, query }) => {
  const r = db.scrapReports.find((x) => x.id === params.id); if (!r) throw bad('Not found', 404);
  if (query.items === '1') db.scrap = db.scrap.filter((x) => x.reportId !== r.id);
  fs.rmSync(path.join(SREP_DIR, r.id + r.ext), { force: true }); db.scrapReports = db.scrapReports.filter((x) => x !== r); save(); return { ok: true };
});
// register many scrap items at once (from a reviewed report, or a bulk entry)
route('POST', '/api/scrap/bulk', true, ({ user, body }) => {
  const rep = body.reportId ? db.scrapReports.find((x) => x.id === body.reportId) : null;
  const items = (body.items || []).filter((i) => String(i.name || '').trim());
  if (!items.length) throw bad('No items to register');
  items.forEach((i, n) => { if (!db.outlets.find((o) => o.code === i.outletCode)) throw bad(`Row ${n + 1} (${i.name}): choose a valid outlet`); if (!(Math.floor(+i.qty) > 0)) throw bad(`Row ${n + 1} (${i.name}): quantity must be at least 1`); });
  const made = items.map((i) => ({ id: uid('sc'), name: String(i.name).trim().slice(0, 120), category: String(i.category || 'Other items'), qty: Math.floor(+i.qty), serial: String(i.serial || '').slice(0, 60), outletCode: i.outletCode, note: String(i.note || (rep ? `From SO scrap report${rep.soName ? ' (' + rep.soName + ')' : ''} dated ${rep.date}` : '')).slice(0, 200), status: 'At outlet', jobId: null, vendorId: null, files: [], at: new Date().toISOString(), by: user.name, reportId: rep?.id || null }));
  db.scrap.push(...made); if (rep) rep.itemIds.push(...made.map((x) => x.id)); save(); return { created: made.length };
});
// admin: delete a stock line (only when nothing is reserved from it)
route('DELETE', '/api/stock/:id', true, ({ user, params }) => {
  const r = db.stock.find((x) => x.id === params.id); if (!r) throw bad('Not found', 404);
  if (r.reserved > 0 || (r.units || []).some((u) => u.status === 'Reserved')) throw bad('Units of this line are reserved for a job - release them first');
  move('DELETE', r, -r.qty, 'Stock line deleted', user.name); db.stock = db.stock.filter((x) => x !== r); save(); return { ok: true };
});

/* ---------- http ---------- */
const MIME = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript', '.css': 'text/css', '.svg': 'image/svg+xml', '.webmanifest': 'application/manifest+json', '.json': 'application/json', '.png': 'image/png', '.jpg': 'image/jpeg', '.webp': 'image/webp', '.ico': 'image/x-icon' };
const readBody = (req) => new Promise((res, rej) => { let d = ''; req.on('data', (c) => { d += c; if (d.length > 14e6) req.destroy(); }); req.on('end', () => { try { res(d ? JSON.parse(d) : {}); } catch { rej(bad('Bad JSON')); } }); });

const server = http.createServer(async (req, res) => {
  const url = new URL(req.url, 'http://x');
  res.setHeader('X-Content-Type-Options', 'nosniff'); res.setHeader('X-Frame-Options', 'SAMEORIGIN'); res.setHeader('Referrer-Policy', 'same-origin');
  if (req.headers['x-forwarded-proto'] === 'https') res.setHeader('Strict-Transport-Security', 'max-age=31536000');
  const send = (code, obj) => { res.writeHead(code, { 'Content-Type': 'application/json' }); res.end(JSON.stringify(obj)); };
  if (url.pathname.startsWith('/api/')) {
    // The single-file RE-Tracker.html (opened from a phone/PC as a local file) sends Origin "null".
    // Sign-in still needs a valid token, so this only lets that file talk to the server.
    if (req.headers.origin === 'null') {
      res.setHeader('Access-Control-Allow-Origin', 'null'); res.setHeader('Vary', 'Origin');
      res.setHeader('Access-Control-Allow-Headers', 'Authorization, Content-Type'); res.setHeader('Access-Control-Allow-Methods', 'GET, POST, PUT, PATCH, DELETE');
      res.setHeader('Access-Control-Allow-Private-Network', 'true'); res.setHeader('Access-Control-Max-Age', '600');
      if (req.method === 'OPTIONS') { res.writeHead(204); return res.end(); }
    }
    try {
      const r = routes.map((r) => ({ r, m: r.method === req.method && url.pathname.match(r.re) })).find((x) => x.m);
      if (!r) return send(404, { error: 'Not found' });
      let user = null;
      if (r.r.fn !== routes[0].fn) {
        const p = verify((req.headers.authorization || '').replace('Bearer ', '') || url.searchParams.get('t'));
        user = p && db.users.find((u) => u.username === p.u);
        if (!user) return send(401, { error: 'Please sign in' });
        if (user.role === 'editor') {
          if (editorBlocked(req.method, url.pathname)) return send(403, { error: 'Editors cannot delete records, cancel issues, manage users or change Settings. Please ask an admin.' });
          user = { ...user, role: 'admin', isEditor: true };
        }
        if (r.r.admin && user.role !== 'admin' && !(user.role === 'viewer' && req.method === 'GET')) return send(403, { error: 'Admin only' });
        if (user.role === 'viewer' && req.method !== 'GET' && url.pathname !== '/api/password') return send(403, { error: 'This is a view-only account - you can look but not change anything' });
        if (user.mustChange && url.pathname !== '/api/password') return send(403, { error: 'Please set a new password first', mustChange: true });
      }
      const body = ['POST', 'PUT', 'PATCH'].includes(req.method) ? await readBody(req) : {};
      const out = await r.r.fn({ user, body, params: r.m.groups || {}, query: Object.fromEntries(url.searchParams), ip: String(req.headers['cf-connecting-ip'] || req.socket.remoteAddress || '') });
      if (out && out.__file) { res.writeHead(200, { 'Content-Type': out.mime, 'Content-Disposition': 'inline; filename="' + out.name.replace(/[^\w.\- ]/g, '_') + '"', 'Cache-Control': 'private, max-age=3600' }); return fs.createReadStream(out.__file).pipe(res); }
      send(200, out);
    } catch (e) { send(e.code || 500, { error: e.message }); if (!e.code) console.error(e); }
    return;
  }
  if (url.pathname === '/brand/logo') {
    const custom = ['png', 'jpg', 'webp'].map((e) => path.join(__dirname, 'data', 'logo.' + e)).find((p) => fs.existsSync(p)), lf = custom || path.join(PUBLIC, 'logo.jpg');
    res.writeHead(200, { 'Content-Type': MIME[path.extname(lf)] || 'image/jpeg', 'Cache-Control': 'public, max-age=300' }); return fs.createReadStream(lf).pipe(res);
  }
  let f = path.normalize(path.join(PUBLIC, url.pathname === '/' ? 'index.html' : url.pathname));
  if (!f.startsWith(PUBLIC) || !fs.existsSync(f) || fs.statSync(f).isDirectory()) f = path.join(PUBLIC, 'index.html');
  res.writeHead(200, { 'Content-Type': MIME[path.extname(f)] || 'application/octet-stream', 'Cache-Control': 'no-cache' });
  fs.createReadStream(f).pipe(res);
});

load();
if (process.argv.includes('--clear-jobs')) { const t = (process.argv.find((a) => a.startsWith('--type=')) || '').slice(7) || null; const r = clearJobs('admin (command line)', t); save(); console.log(JSON.stringify(r)); process.exit(0); }
if (process.argv.includes('--import')) cli(); else server.listen(PORT, () => console.log(`Retail Engineering Tracker running on http://localhost:${PORT}`));
