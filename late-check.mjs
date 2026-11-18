// Avaibe – late clock-in reminder
// Runs from GitHub Actions. Emails anyone with a shift today who has not clocked in
// by shift start + GRACE_MIN, from hr@avaibe.com with a copy to HR (Sujit).
//
// Env:
//   ZONE        IST | UAE        which cohort to check (set by the workflow schedule)
//   DRY_RUN     "true"           list who would be emailed, send nothing
//   TEST_TO     email address    send ONE sample email to this address only, nothing else
//   GRAPH_TOKEN (optional)       pre-fetched token; otherwise obtained via GitHub OIDC

const CONFIG = {
  tenantId: "41385911-2567-4873-bcf2-e28193046199",
  clientId: "55833e79-a769-4d81-897d-93824a9402f0",
  teamId:   "55a0dee6-69f7-447d-8a27-ae00decfad30",       // Avaibe – Attendance
  sender:   "hr@avaibe.com",
  cc:       ["sujit@avaibe.com"],
  graceMin: 10,
  exclude:  ["abhilash@avaibe.com", "director@avaibe.com", "sujit@avaibe.com"],
  zones: {
    IST: { tz: "Asia/Kolkata", label: "IST", offset: "+05:30" },
    UAE: { tz: "Asia/Dubai",   label: "UAE", offset: "+04:00" }
  },
  // Staff working from India; everyone else is UAE.
  indiaStaff: ["archana","arjun","athithya","ayisha","gayathri","mohammedafsal","mohan","rashid",
               "renganathan","sathyavani","sujit"].map(n => n + "@avaibe.com")
};

const GRAPH = "https://graph.microsoft.com/v1.0";
const env = process.env;
const ZONE = (env.ZONE || "").toUpperCase();
const DRY = String(env.DRY_RUN || "").toLowerCase() === "true";
const TEST_TO = (env.TEST_TO || "").trim();
const now = env.NOW ? new Date(env.NOW) : new Date();

const log = (...a) => console.log(...a);
const fail = m => { console.error("ERROR:", m); process.exit(1); };

/* ---------- auth: GitHub OIDC -> Entra (no stored secret) ---------- */
async function getToken() {
  if (env.GRAPH_TOKEN) return env.GRAPH_TOKEN;
  if (!env.ACTIONS_ID_TOKEN_REQUEST_URL) fail("No OIDC token available – the workflow needs 'permissions: id-token: write'.");
  const oidc = await fetch(`${env.ACTIONS_ID_TOKEN_REQUEST_URL}&audience=api://AzureADTokenExchange`,
    { headers: { Authorization: `Bearer ${env.ACTIONS_ID_TOKEN_REQUEST_TOKEN}` } }).then(r => r.json());
  if (!oidc.value) fail("Could not get GitHub OIDC token.");
  try { log("OIDC subject:", JSON.parse(Buffer.from(oidc.value.split(".")[1], "base64url").toString()).sub); } catch {}
  const body = new URLSearchParams({
    client_id: CONFIG.clientId, grant_type: "client_credentials", scope: "https://graph.microsoft.com/.default",
    client_assertion_type: "urn:ietf:params:oauth:client-assertion-type:jwt-bearer", client_assertion: oidc.value
  });
  const res = await fetch(`https://login.microsoftonline.com/${CONFIG.tenantId}/oauth2/v2.0/token`, { method: "POST", body });
  const j = await res.json();
  if (!j.access_token) fail("Entra token exchange failed: " + (j.error_description || JSON.stringify(j)).slice(0, 400));
  return j.access_token;
}

/* ---------- graph helpers ---------- */
let TOKEN;
async function graph(path, opts = {}) {
  const res = await fetch(path.startsWith("http") ? path : GRAPH + path, {
    ...opts, headers: { Authorization: `Bearer ${TOKEN}`, "Content-Type": "application/json", ...(opts.headers || {}) }
  });
  if (!res.ok) {
    const t = await res.text();
    throw new Error(`${res.status} ${path.split("?")[0]} – ${t.slice(0, 300)}`);
  }
  return res.status === 202 || res.status === 204 ? null : res.json();
}
async function graphAll(path) {
  let out = [], next = path, n = 0;
  while (next && n++ < 50) { const j = await graph(next); out = out.concat(j.value || []); next = j["@odata.nextLink"]; }
  return out;
}

/* ---------- time helpers ---------- */
const dateKey = (d, tz) => new Intl.DateTimeFormat("en-CA", { timeZone: tz, year: "numeric", month: "2-digit", day: "2-digit" }).format(d);
const fmtTime = (d, tz) => new Intl.DateTimeFormat("en-GB", { timeZone: tz, hour: "numeric", minute: "2-digit", hour12: true }).format(d).toUpperCase();
const fmtDate = (d, tz) => new Intl.DateTimeFormat("en-GB", { timeZone: tz, weekday: "long", day: "numeric", month: "long", year: "numeric" }).format(d);
const esc = s => String(s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));

/* ---------- email ---------- */
function buildEmail({ firstName, deadline, tzLabel, dayText }) {
  const subject = `Missed clock-in – ${dayText}`;
  const html = `
<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#1b1f24;max-width:560px">
  <p>Hi ${esc(firstName)},</p>
  <p>Our Shifts time clock shows that you had <b>not clocked in by ${esc(deadline)} ${esc(tzLabel)}</b> today, ${esc(dayText)}.</p>
  <p style="font-size:15px"><b>You have missed your clock-in time. Please contact HR.</b></p>
  <p>If you forgot to clock in, please submit a <b>Missed clock-in / clock-out</b> request in Teams
     (Approvals → New approval request) so HR can correct your record. Please do not edit your own time sheet.</p>
  <p>Thank you,<br>HR – Avaibe</p>
  <p style="color:#6b7280;font-size:12px">This is an automated message sent when no clock-in is recorded ${CONFIG.graceMin} minutes after your shift starts.</p>
</div>`;
  return { subject, html };
}
async function sendMail(to, cc, { subject, html }) {
  await graph(`/users/${encodeURIComponent(CONFIG.sender)}/sendMail`, {
    method: "POST",
    body: JSON.stringify({
      message: {
        subject, body: { contentType: "HTML", content: html },
        toRecipients: [{ emailAddress: { address: to } }],
        ccRecipients: cc.map(a => ({ emailAddress: { address: a } }))
      },
      saveToSentItems: true
    })
  });
}

/* ---------- main ---------- */
async function main() {
  const zone = CONFIG.zones[ZONE];
  if (!zone) fail(`ZONE must be IST or UAE (got "${ZONE}")`);
  TOKEN = await getToken();

  if (TEST_TO) {
    const mail = buildEmail({ firstName: "Test", deadline: "9:10 AM", tzLabel: zone.label, dayText: fmtDate(now, zone.tz) });
    mail.subject = "[TEST] " + mail.subject;
    await sendMail(TEST_TO, [], mail);
    log(`Test email sent from ${CONFIG.sender} to ${TEST_TO}.`);
    return;
  }

  const today = dateKey(now, zone.tz);
  const dayStart = new Date(`${today}T00:00:00${zone.offset}`), dayEnd = new Date(dayStart.getTime() + 864e5);
  log(`Zone ${zone.label} · ${today} · now ${now.toISOString()} · grace ${CONFIG.graceMin} min · ${DRY ? "DRY RUN" : "LIVE"}`);

  const T = CONFIG.teamId;
  const [shifts, timeOff, cards] = await Promise.all([
    graphAll(`/teams/${T}/schedule/shifts?$filter=sharedShift/startDateTime ge ${dayStart.toISOString()} and sharedShift/endDateTime le ${new Date(dayEnd.getTime() + 864e5).toISOString()}`),
    graphAll(`/teams/${T}/schedule/timesOff?$filter=sharedTimeOff/startDateTime le ${dayEnd.toISOString()} and sharedTimeOff/endDateTime ge ${dayStart.toISOString()}`).catch(e => { log("Time off lookup failed, continuing:", e.message); return []; }),
    graphAll(`/teams/${T}/schedule/timeCards?$filter=clockInEvent/dateTime ge ${new Date(dayStart.getTime() - 6 * 3600e3).toISOString()}&$top=200`)
      .catch(async e => { if (!/400/.test(e.message)) throw e; return graphAll(`/teams/${T}/schedule/timeCards?$top=200`); })
  ]);

  const ids = [...new Set(shifts.map(s => s.userId).filter(Boolean))];
  const users = {};
  for (let i = 0; i < ids.length; i += 15) {
    const chunk = ids.slice(i, i + 15).map(x => `'${x}'`).join(",");
    (await graphAll(`/users?$filter=id in (${chunk})&$select=id,displayName,givenName,mail,userPrincipalName`)).forEach(u => users[u.id] = u);
  }
  const emailOf = id => (users[id]?.mail || users[id]?.userPrincipalName || "").toLowerCase();
  const zoneOf = id => CONFIG.indiaStaff.includes(emailOf(id)) ? "IST" : "UAE";

  const results = [];
  for (const s of shifts) {
    const uid = s.userId, email = emailOf(uid);
    if (!uid || !email || CONFIG.exclude.includes(email) || zoneOf(uid) !== ZONE) continue;
    const start = new Date(s.sharedShift.startDateTime);
    if (dateKey(start, zone.tz) !== today) continue;                         // not today's shift
    const deadline = new Date(start.getTime() + CONFIG.graceMin * 60000);
    const name = users[uid]?.displayName || email;
    if (now < deadline) { results.push({ name, email, action: "skip – grace period not over yet" }); continue; }
    const off = timeOff.some(t => t.userId === uid && new Date(t.sharedTimeOff.startDateTime) <= deadline && new Date(t.sharedTimeOff.endDateTime) >= start);
    if (off) { results.push({ name, email, action: "skip – time off" }); continue; }
    const onTime = cards.some(c => c.userId === uid && c.clockInEvent?.dateTime &&
      new Date(c.clockInEvent.dateTime) >= new Date(start.getTime() - 6 * 3600e3) && new Date(c.clockInEvent.dateTime) <= deadline);
    if (onTime) { results.push({ name, email, action: "ok – clocked in on time" }); continue; }
    const lateCard = cards.find(c => c.userId === uid && new Date(c.clockInEvent?.dateTime) > deadline && new Date(c.clockInEvent?.dateTime) < dayEnd);
    const firstName = users[uid]?.givenName || name.split(" ")[0];
    const mail = buildEmail({ firstName, deadline: fmtTime(deadline, zone.tz), tzLabel: zone.label, dayText: fmtDate(start, zone.tz) });
    if (!DRY) await sendMail(email, CONFIG.cc, mail);
    results.push({ name, email, action: `${DRY ? "WOULD EMAIL" : "EMAILED"} – ${lateCard ? "clocked in late at " + fmtTime(new Date(lateCard.clockInEvent.dateTime), zone.tz) : "no clock-in"}` });
  }

  log(`\n${results.length} scheduled ${zone.label} staff checked:`);
  for (const r of results) log(` • ${r.name} <${r.email}>: ${r.action}`);
  const n = results.filter(r => r.action.startsWith("EMAILED") || r.action.startsWith("WOULD")).length;
  log(`\n${n} reminder(s) ${DRY ? "would be sent" : "sent"}.`);
  if (env.GITHUB_STEP_SUMMARY) {
    const fs = await import("node:fs");
    fs.appendFileSync(env.GITHUB_STEP_SUMMARY, `### Late clock-in check – ${zone.label} ${today}${DRY ? " (dry run)" : ""}\n\n| Staff | Result |\n|---|---|\n` +
      results.map(r => `| ${r.name} | ${r.action} |`).join("\n") + `\n\n**${n} reminder(s) ${DRY ? "would be sent" : "sent"}.**\n`);
  }
}

main().catch(e => fail(e.message));
