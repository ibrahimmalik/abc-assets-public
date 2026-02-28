/*******************************************************
 * ABC Hyde Park — Partner Discounts (QR Validation) v2.3
 * -----------------------------------------------------
 * WHAT THIS DOES:
 * - Partner QR opens: /exec?p=<partner_id>
 * - Guest enters Booking Ref + Last Name
 * - Server validates booking eligibility
 * - Returns a "VALID" proof screen with a live countdown
 *
 * PHASE 2 (anti-screenshot proof):
 * - On approval, server sends:
 *   - approval_code (unique code)
 *   - approved_at (server time)
 *   - expires_at (server time + TTL)
 *   - booking_ref_masked (safe partial ref)
 *
 * PHASE 2.3 (mistake-friendly rule):
 * - Default proof TTL = 10 minutes
 * - If Verify is pressed again within 2 minutes
 *   (same partner + booking), allow again BUT TTL = 5 minutes
 * - Re-issues do NOT consume extra "night uses"
 *
 * REPORTS:
 * - /exec?report=1&p=<partner_id>&pin=<report_pin>
 *
 * Abuse controls:
 * - Max 1 "real" approval per day per partner per booking
 *   (except the 2-minute re-issue window)
 * - Max total approvals per partner per booking = number of NIGHTS stayed
 *   (re-issues do not count towards this)
 *******************************************************/

/** Sheet tab names (must match exactly) */
const SHEET_PARTNERS = "Partners";
const SHEET_BOOKINGS = "Bookings";
const SHEET_REDEMPTIONS = "Redemptions";

/** Proof timing (Phase 2.3) */
const PROOF_TTL_SECONDS = 10 * 60;        // 10 minutes normal
const REVERIFY_WINDOW_SECONDS = 2 * 60;   // allow re-verify within 2 minutes
const REVERIFY_TTL_SECONDS = 5 * 60;      // re-issue proof lasts 5 minutes

/** Convert Sheets boolean/text to real boolean */
function isTrue(v) {
  return v === true || String(v).trim().toUpperCase() === "TRUE";
}

/** Normalize strings for matching */
function normalize(s) {
  return String(s || "").trim().toLowerCase();
}

/** Midnight date for comparisons (prevents time-of-day edge cases) */
function startOfDay(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

/**
 * Robust date parsing (Sheets may store Date or text)
 * Supports:
 * - Date objects
 * - YYYY-MM-DD
 * - DD/MM/YYYY
 * - JS fallback
 */
function parseSheetDate(value) {
  if (!value) return null;

  // Already a Date
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value)) {
    return value;
  }

  const s = String(value).trim();
  if (!s) return null;

  // YYYY-MM-DD
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) {
    const y = Number(m[1]), mo = Number(m[2]) - 1, d = Number(m[3]);
    const dt = new Date(y, mo, d);
    return isNaN(dt) ? null : dt;
  }

  // DD/MM/YYYY or D/M/YYYY
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) {
    const d = Number(m[1]), mo = Number(m[2]) - 1, y = Number(m[3]);
    const dt = new Date(y, mo, d);
    return isNaN(dt) ? null : dt;
  }

  // Fallback
  const dt = new Date(s);
  return isNaN(dt) ? null : dt;
}

/** Always use the bound spreadsheet in Web App execution */
function getSS() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * GET handler:
 * - /exec?p=<partner_id> => Validate page
 * - /exec?report=1&p=<partner_id>&pin=<pin> => Report page (PIN protected)
 */
function doGet(e) {
  const partnerId = (e.parameter.p || "").trim();
  const isReport = (e.parameter.report || "") === "1";

  if (!partnerId) return html("Missing partner id.");

  const partner = getPartner(partnerId);
  if (!partner || !isTrue(partner.is_active)) return html("Partner inactive or not found.");

  // Report (PIN protected)
  if (isReport) {
    const pin = String(e.parameter.pin || "").trim();
    const expected = String(partner.report_pin || "").trim();
    if (!pin || pin !== expected) return html("Wrong PIN.");

    const reportData = buildPartnerReport(partnerId);
    const t = HtmlService.createTemplateFromFile("Report");
    t.data = { partner, report: reportData };
    return t.evaluate().setTitle(`${partner.partner_name} — Report`);
  }

  // Validate page (guest)
  const t = HtmlService.createTemplateFromFile("Validate");
  t.data = {
    post_url: ScriptApp.getService().getUrl(), // always correct /exec
    partner_id: partner.partner_id,
    partner_name: partner.partner_name,

    // Keep these available for future geo-fencing, but client can ignore to avoid popups
    geo_enabled: isTrue(partner.geo_enabled),
    geo_lat: partner.geo_lat,
    geo_lng: partner.geo_lng,
    geo_radius_m: partner.geo_radius_m
  };

  return t.evaluate().setTitle("ABC Hyde Park — Guest Access Validation");
}

/**
 * POST handler:
 * Validates guest submission and returns JSON.
 *
 * NOTE:
 * We intentionally do NOT show discount % on screen.
 * Partners only need "VALID / NOT VALID".
 */
function doPost(e) {
  const partnerId  = (e.parameter.p || "").trim();
  const bookingRef = (e.parameter.booking_ref || "").trim();
  const lastName   = (e.parameter.last_name || "").trim();
  const deviceNote = (e.parameter.device_note || "").trim();
  const lat        = (e.parameter.lat || "").trim();
  const lng        = (e.parameter.lng || "").trim();

  if (!partnerId) return json({ ok:false, message:"Missing partner id." });
  if (!bookingRef || !lastName) return json({ ok:false, message:"Enter booking reference and last name." });

  const partner = getPartner(partnerId);
  if (!partner || !isTrue(partner.is_active)) return json({ ok:false, message:"Partner inactive or not found." });

  /******** OPTIONAL SERVER-SIDE GEOFENCE ********
   * Only used if partner.geo_enabled is TRUE.
   * Client can still send blank lat/lng if you don't want location prompts.
   */
  if (isTrue(partner.geo_enabled)) {
    const glat = Number(lat), glng = Number(lng);
    const plat = Number(partner.geo_lat), plng = Number(partner.geo_lng);
    const radius = Number(partner.geo_radius_m || 120);

    if (!isFinite(glat) || !isFinite(glng)) {
      logRedemption(false, "Location required (geofence enabled)", partner, bookingRef, lastName, deviceNote, lat, lng);
      return json({ ok:false, message:"Location required here. Please allow location access and try again." });
    }

    const dist = distanceMeters(glat, glng, plat, plng);
    if (dist > radius) {
      logRedemption(false, `Outside geofence (${Math.round(dist)}m > ${radius}m)`, partner, bookingRef, lastName, deviceNote, lat, lng);
      return json({ ok:false, message:"Not at the partner location. Please validate at the counter." });
    }
  }

  /******** FIND BOOKING ********/
  const booking = findBooking(bookingRef, lastName);
  if (!booking) {
    logRedemption(false, "Booking not found / name mismatch", partner, bookingRef, lastName, deviceNote, lat, lng);
    return json({ ok:false, message:"Not found. Check booking ref + last name spelling." });
  }

  /******** IN-HOUSE STATUS RULE ********/
  const st = normalize(booking.status);

  // Allowed statuses (Guestline-friendly)
  const allowed = new Set(["active","resident","inhouse","in-house","checked-in","checked in"]);

  // Blocked statuses
  const blocked = new Set(["cancelled","canceled","no-show","noshow","check-out","checked-out","checked out","departed"]);

  if (blocked.has(st) || (!allowed.has(st) && st !== "")) {
    logRedemption(false, `Not in-house (status=${booking.status})`, partner, bookingRef, lastName, deviceNote, lat, lng);
    return json({ ok:false, message:"Access is limited to in-house guests only." });
  }

  /******** DATE VALIDITY WINDOW ********
   * Valid from check-in date through (check-out + 1 day)
   */
  const now = startOfDay(new Date());
  const checkInRaw  = parseSheetDate(booking.check_in);
  const checkOutRaw = parseSheetDate(booking.check_out);

  if (!checkInRaw || !checkOutRaw) {
    logRedemption(false, "Invalid booking dates", partner, bookingRef, lastName, deviceNote, lat, lng);
    return json({ ok:false, message:"Booking dates invalid. Please contact reception." });
  }

  const checkIn  = startOfDay(checkInRaw);
  const checkOut = startOfDay(checkOutRaw);
  const expiry   = new Date(checkOut.getTime() + 24*60*60*1000); // checkout + 1 day

  if (now < checkIn) {
    logRedemption(false, "Before check-in", partner, bookingRef, lastName, deviceNote, lat, lng);
    return json({ ok:false, message:"Not valid yet (before check-in)." });
  }

  if (now > expiry) {
    logRedemption(false, "Expired (after checkout+1)", partner, bookingRef, lastName, deviceNote, lat, lng);
    return json({ ok:false, message:"Expired (after checkout + 1 day)." });
  }

  /******** RULE 1: MAX 1 APPROVAL PER DAY ********
   * Phase 2.3 exception:
   * - If they press verify again within 2 minutes,
   *   allow it as a "REISSUE" with shorter TTL.
   */
  let isReissue = false;

  if (alreadyRedeemedToday(partnerId, bookingRef)) {
    const lastTs = getLastApprovedTimestamp(partnerId, bookingRef);

    if (lastTs) {
      const secondsSince = (new Date().getTime() - lastTs.getTime()) / 1000;

      if (secondsSince <= REVERIFY_WINDOW_SECONDS) {
        isReissue = true; // allow re-verify due to likely typo/mistake
      } else {
        logRedemption(false, "Already redeemed today", partner, bookingRef, lastName, deviceNote, lat, lng);
        return json({ ok:false, message:"Already redeemed today at this partner." });
      }
    } else {
      // If we can't find timestamp, stay strict
      logRedemption(false, "Already redeemed today", partner, bookingRef, lastName, deviceNote, lat, lng);
      return json({ ok:false, message:"Already redeemed today at this partner." });
    }
  }

  /******** RULE 2: MAX USES PER STAY (NIGHTS) ********
   * IMPORTANT:
   * - Re-issues should NOT consume extra uses.
   * - We implement this by NOT counting REISSUE rows.
   */
  const maxUses = bookingNights(booking);
  const usedSoFar = countApprovedRedemptions(partnerId, bookingRef);

  if (!isReissue && usedSoFar >= maxUses) {
    logRedemption(false, `Max uses reached (${usedSoFar}/${maxUses})`, partner, bookingRef, lastName, deviceNote, lat, lng);
    return json({ ok:false, message:`Limit reached for this stay (${maxUses} uses).` });
  }

  /******** ✅ APPROVED (PROOF) ********/
  const serverNow = new Date();

  // Choose TTL based on normal approval vs re-issue
  const ttlSeconds = isReissue ? REVERIFY_TTL_SECONDS : PROOF_TTL_SECONDS;
  const expiresAt = new Date(serverNow.getTime() + (ttlSeconds * 1000));

  const approvalCode = generateApprovalCode(partner.partner_id);

  // Log:
  // - Normal approvals count as real uses
  // - Re-issues are marked so we can ignore them in "use counting"
  const reason = isReissue
    ? `REISSUE (CODE:${approvalCode})`
    : `Approved (CODE:${approvalCode})`;

  logRedemption(true, reason, partner, bookingRef, lastName, deviceNote, lat, lng);

  return json({
    ok: true,
    message: `VALID ✅ Guest verified for ${partner.partner_name}`,
    booking_ref_masked: maskBookingRef(bookingRef),
    approval_code: approvalCode,
    approved_at: serverNow.toISOString(),
    expires_at: expiresAt.toISOString(),
    proof_ttl_seconds: ttlSeconds
  });
}

/*******************************************************
 * Proof helpers (Phase 2)
 *******************************************************/

/** Safe display: don't show full booking ref on screen */
function maskBookingRef(ref) {
  const s = String(ref || "").trim();
  if (s.length <= 4) return s;
  const head = s.slice(0, 2);
  const tail = s.slice(-2);
  return `${head}***${tail}`;
}

/** Human-readable code for staff to quickly spot */
function generateApprovalCode(partnerId) {
  const p = String(partnerId || "ABC").trim().toUpperCase();
  const num = Math.floor(10000 + Math.random() * 90000);
  return `${p}-${num}`;
}

/*******************************************************
 * Booking / Redemptions logic
 *******************************************************/

/** Calculate nights from check_out - check_in (minimum 1) */
function bookingNights(booking) {
  const inD  = parseSheetDate(booking.check_in);
  const outD = parseSheetDate(booking.check_out);
  if (!inD || !outD) return 1;

  const checkIn = startOfDay(inD);
  const checkOut = startOfDay(outD);

  const ms = checkOut.getTime() - checkIn.getTime();
  const nights = Math.round(ms / (24*60*60*1000));
  return Math.max(1, nights);
}

/** True if a row is a re-issue (so it should not consume a stay use) */
function isReissueReason(reason) {
  return String(reason || "").toUpperCase().startsWith("REISSUE");
}

/**
 * Count total approved redemptions for booking at partner (across all time),
 * EXCLUDING re-issues (Phase 2.3)
 */
function countApprovedRedemptions(partnerId, bookingRef) {
  const sh = getSS().getSheetByName(SHEET_REDEMPTIONS);
  if (!sh) return 0;

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return 0;

  const headers = values[0];
  let count = 0;

  for (let i=1; i<values.length; i++){
    const row = asObj(headers, values[i]);
    if (!isTrue(row.approved)) continue;
    if (isReissueReason(row.reason)) continue; // key line for Phase 2.3
    if (String(row.partner_id).trim() !== String(partnerId).trim()) continue;
    if (normalize(row.booking_ref) !== normalize(bookingRef)) continue;
    count++;
  }
  return count;
}

/** Prevent multiple approvals in same day for same partner+booking */
function alreadyRedeemedToday(partnerId, bookingRef) {
  const sh = getSS().getSheetByName(SHEET_REDEMPTIONS);
  if (!sh) return false;

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return false;

  const headers = values[0];
  const today = startOfDay(new Date());

  for (let i=1; i<values.length; i++){
    const row = asObj(headers, values[i]);
    if (!isTrue(row.approved)) continue;
    if (String(row.partner_id).trim() !== String(partnerId).trim()) continue;
    if (normalize(row.booking_ref) !== normalize(bookingRef)) continue;

    const tsDay = startOfDay(new Date(row.timestamp));
    if (tsDay.getTime() === today.getTime()) return true;
  }
  return false;
}

/**
 * Finds the most recent APPROVED redemption timestamp for partner+booking.
 * Used to allow re-verify within 2 minutes.
 */
function getLastApprovedTimestamp(partnerId, bookingRef) {
  const sh = getSS().getSheetByName(SHEET_REDEMPTIONS);
  if (!sh) return null;

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return null;

  const headers = values[0];
  const refNorm = normalize(bookingRef);
  const pid = String(partnerId).trim();

  let latest = null;

  for (let i = 1; i < values.length; i++) {
    const row = asObj(headers, values[i]);
    if (!isTrue(row.approved)) continue;
    if (String(row.partner_id).trim() !== pid) continue;
    if (normalize(row.booking_ref) !== refNorm) continue;

    const ts = new Date(row.timestamp);
    if (isNaN(ts)) continue;

    if (!latest || ts > latest) latest = ts;
  }

  return latest;
}

/*******************************************************
 * Sheets access
 *******************************************************/

function getPartner(partnerId) {
  const sh = getSS().getSheetByName(SHEET_PARTNERS);
  if (!sh) return null;

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return null;

  const headers = values[0];
  for (let i=1; i<values.length; i++){
    const row = asObj(headers, values[i]);
    if (String(row.partner_id).trim() === String(partnerId).trim()) return row;
  }
  return null;
}

function findBooking(bookingRef, lastName) {
  const sh = getSS().getSheetByName(SHEET_BOOKINGS);
  if (!sh) return null;

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return null;

  const headers = values[0];
  const refNorm = normalize(bookingRef);
  const nameNorm = normalize(lastName);

  for (let i=1; i<values.length; i++){
    const row = asObj(headers, values[i]);
    if (normalize(row.booking_ref) === refNorm && normalize(row.last_name) === nameNorm) return row;
  }
  return null;
}

function logRedemption(approved, reason, partner, bookingRef, lastName, deviceNote, lat, lng) {
  const sh = getSS().getSheetByName(SHEET_REDEMPTIONS);
  if (!sh) return;

  // Your existing columns are preserved (no sheet changes required)
  sh.appendRow([
    new Date(),
    partner.partner_id,
    partner.partner_name,
    partner.discount_percent, // internal only
    bookingRef,
    lastName,
    approved,
    reason,
    deviceNote,
    lat || "",
    lng || ""
  ]);
}

function buildPartnerReport(partnerId) {
  const sh = getSS().getSheetByName(SHEET_REDEMPTIONS);
  if (!sh) return { today:0, week:0, month:0, total:0 };

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return { today:0, week:0, month:0, total:0 };

  const headers = values[0];
  const today = startOfDay(new Date());
  const weekAgo = new Date(today.getTime() - 6*24*60*60*1000);
  const monthAgo = new Date(today.getTime() - 29*24*60*60*1000);

  let total=0, t=0, w=0, m=0;

  for (let i=1; i<values.length; i++){
    const row = asObj(headers, values[i]);
    if (String(row.partner_id).trim() !== String(partnerId).trim()) continue;
    if (!isTrue(row.approved)) continue;

    total++;
    const ts = startOfDay(new Date(row.timestamp));
    if (ts.getTime() === today.getTime()) t++;
    if (ts >= weekAgo) w++;
    if (ts >= monthAgo) m++;
  }

  return { today:t, week:w, month:m, total };
}

/*******************************************************
 * Rendering / JSON helpers
 *******************************************************/

function asObj(headers, row){
  const o = {};
  headers.forEach((h, idx) => o[h] = row[idx]);
  return o;
}

function html(msg){
  return HtmlService.createHtmlOutput(
    `<div style="font-family:system-ui;padding:24px">${escapeHtml(msg)}</div>`
  );
}

function json(o){
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}

function escapeHtml(s){
  return String(s).replace(/[&<>"']/g, c => ({
    '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'
  }[c]));
}

/*******************************************************
 * Geo: Haversine distance (meters)
 *******************************************************/
function distanceMeters(lat1, lon1, lat2, lon2){
  const R = 6371000;
  const toRad = x => x * Math.PI/180;
  const dLat = toRad(lat2-lat1);
  const dLon = toRad(lon2-lon1);
  const a = Math.sin(dLat/2)**2 +
            Math.cos(toRad(lat1))*Math.cos(toRad(lat2))*
            Math.sin(dLon/2)**2;
  return 2*R*Math.asin(Math.sqrt(a));
}
