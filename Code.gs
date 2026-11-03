/**
 * Aakash SystemCraft — Website backend (Google Apps Script)
 * 1) Lead form  → "Received data" tab
 * 2) Portfolio  → "Website Portfolio" tab se site par apne aap dikhta hai
 *
 * SETUP (ek baar):
 *  - Sheet mein: Extensions > Apps Script > ye pura code paste karo > Save
 *  - Function dropdown se "setup" chuno > Run > Allow
 *  - Deploy > New deployment > Web app | Execute as: Me | Who has access: Anyone
 *  - Web App URL ko website HTML ke SHEET_URL mein paste karo
 * Code badalne ke baad: Deploy > Manage deployments > Edit (pencil) > Version: New version > Deploy (URL wahi rahega)
 */

const LEAD_SHEET = 'Received data';
const LEAD_HEADERS = ['Timestamp','Name','Mobile','Email','Business','City','Services Needed','Message','Source','Status','Follow-up Date','Remarks'];
const ALERT_EMAIL = true; // har naye lead par email

const PF_SHEET = 'Website Portfolio';
const PF_HEADERS = ['Show (Yes/No)','Category','Title','Type','Link','Image URL (optional)','Description (optional)','Order'];
const PF_TYPES = ['Software','Website','Video','Image','PPT','Other'];

/* ---------------- WEB APP ---------------- */
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';
  if (action === 'portfolio') return json_(getPortfolio_());
  return json_({ ok: true, message: 'Aakash SystemCraft backend is live' });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const p = (e && e.parameter) || {};
    const sh = leadSheet_();
    sh.appendRow([new Date(), clean_(p.name), "'" + clean_(p.phone), clean_(p.email), clean_(p.business),
                  clean_(p.city), clean_(p.services), clean_(p.message), clean_(p.source), 'New', '', '']);
    if (ALERT_EMAIL) {
      MailApp.sendEmail(Session.getEffectiveUser().getEmail(), 'New website lead: ' + clean_(p.name),
        'Naam: ' + clean_(p.name) + '\nMobile: ' + clean_(p.phone) + '\nEmail: ' + clean_(p.email) +
        '\nBusiness: ' + clean_(p.business) + '\nCity: ' + clean_(p.city) +
        '\nServices: ' + clean_(p.services) + '\nMessage: ' + clean_(p.message) +
        '\n\nWhatsApp: https://wa.me/91' + clean_(p.phone).replace(/\D/g, '').slice(-10));
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/* ---------------- PORTFOLIO ---------------- */
function getPortfolio_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('pf');
  if (hit) return JSON.parse(hit);
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PF_SHEET);
  if (!sh || sh.getLastRow() < 2) return { ok: true, items: [] };
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, PF_HEADERS.length).getValues();
  const items = rows
    .map((r, i) => ({ show: String(r[0]).trim().toLowerCase(), category: String(r[1]).trim(), title: String(r[2]).trim(),
                      type: String(r[3]).trim() || 'Other', link: String(r[4]).trim(), image: String(r[5]).trim(),
                      desc: String(r[6]).trim(), order: Number(r[7]) || 1000 + i }))
    .filter(it => it.show !== 'no' && it.title && (it.link || it.image))
    .sort((a, b) => a.order - b.order)
    .map(({ show, ...rest }) => rest);
  const out = { ok: true, items: items, updated: new Date().toISOString() };
  try { cache.put('pf', JSON.stringify(out), 60); } catch (e) {}
  return out;
}

// Sheet edit hote hi cache clear, taaki site par turant update dikhe
function onEdit(e) {
  if (e && e.range && e.range.getSheet().getName() === PF_SHEET) CacheService.getScriptCache().remove('pf');
}

/* ---------------- SETUP ---------------- */
function setup() {
  leadSheet_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(PF_SHEET);
  if (!sh) sh = ss.insertSheet(PF_SHEET);
  if (sh.getLastRow() === 0) {
    sh.appendRow(PF_HEADERS);
    styleHeader_(sh, PF_HEADERS.length);
    const seed = [
    ["Yes", "Inventory & Purchase", "Master ERP (IMS)", "Software", "https://script.google.com/macros/s/AKfycbz3naZPuGx5xMQz4Pxs1Yz4NG6d3VHrjQtIHMaZ6FFokDUvLUZ18RTptZ3Zaoqxaw_O3Q/exec", "", "", 1],
    ["Yes", "Inventory & Purchase", "Master IMS", "Software", "https://script.google.com/macros/s/AKfycbxIP-zh4XJyIjal1ydyHGdntqW1NUIjSXF0gZwZob_65XoAyfLt4iK_e1jQ4yRZL0XyTg/exec", "", "", 2],
    ["Yes", "Inventory & Purchase", "Inventory Software", "Software", "https://script.google.com/macros/s/AKfycbwuL061h1AOJeNfcgifcGa-26_O9CrQCdjBb8A4VLEcUVbnZ7kSNDG1hvgYqGOwBrGJEQ/exec", "", "", 3],
    ["Yes", "Inventory & Purchase", "Material Receiving", "Software", "https://script.google.com/macros/s/AKfycbxFOGDkily-vf8I6Q5pMskZ0Mbl2kw5QbTSZEDE4iAA0du_4B0EvtnmRvkTGfZ9AkRM/exec", "", "", 4],
    ["Yes", "Inventory & Purchase", "PO Management", "Software", "https://script.google.com/macros/s/AKfycbyocbF8nj-vOM1otMVrEeIwZ4M-Zi2m2hQaWsknSlLv2hHxzPFdZgB3u_MsYOpGdmNx/exec?page=po_management", "", "", 5],
    ["Yes", "Inventory & Purchase", "Requirement", "Software", "https://script.google.com/macros/s/AKfycbyN63_Cuzke7mEONWA9izAnH4wLbxcOx1P-GIyP_-Yh5LM2ZtlDeO2Q5D_eZWRnnx7LNA/exec", "", "", 6],
    ["Yes", "Inventory & Purchase", "Add Product", "Software", "https://script.google.com/macros/s/AKfycbxAteHxl2LhAkhSUCiv6cMzIdSgEePhkz3akbuMRX32yuKeKAdR_9TjyErFU8S1dzJmew/exec", "", "", 7],
    ["Yes", "Sales & Orders", "Order Entry", "Software", "https://script.google.com/macros/s/AKfycbwHrvJLiBzUoLuOgWpDdQRDyjil0wZo2OEFDQ__8YoDZhxjS3vDnxw1eQFX64lXTvVwaw/exec?page=order_form", "", "", 8],
    ["Yes", "Sales & Orders", "Sales System", "Software", "https://script.google.com/macros/s/AKfycbzd0qB3LkNe0i-GhmcTQ5IJ5c-E7LLs1exfDxvkjd7VI_0Zd4me2KsN_avE5p2DM1UBTQ/exec", "", "", 9],
    ["Yes", "Sales & Orders", "Sales Admin", "Software", "https://script.google.com/macros/s/AKfycbxvZOiXHNANKHEWjx25glaM1Qid84aXYmtSaagtyTtS8gwHAX-bK0FOa4_XOTzwYyAC/exec", "", "", 10],
    ["Yes", "Sales & Orders", "Sales Tracking", "Software", "https://script.google.com/macros/s/AKfycbwdg9QsymMYYrJ4gOjoccB16AXYobybunztloAEORT6E3HRtPkBvXB83QlrBf6sINTDBw/exec", "", "", 11],
    ["Yes", "Sales & Orders", "Sales Tracking App (AppSheet)", "Software", "https://www.appsheet.com/start/740d51d4-d6e6-4710-87bf-6b75582250a4?platform=desktop#appName=aakashapp-138827931&view=SALES%20APP%20TRACKING", "", "", 12],
    ["Yes", "Sales & Orders", "Sales Autopilot", "Software", "https://sites.google.com/view/sales-autopilot/home", "", "", 13],
    ["Yes", "Sales & Orders", "Production Autopilot", "Software", "https://sites.google.com/view/busines-growth/home", "", "", 14],
    ["Yes", "Dashboards, Tasks & HR", "PMS Admin Dashboard", "Software", "https://script.google.com/macros/s/AKfycbxDK5JoBsA08PNvPTDwec5ipBrThMPHsnaM3hJVC63NWo_6xxLu8e2JzDj65EYa3z-a/exec", "", "", 15],
    ["Yes", "Dashboards, Tasks & HR", "User Dashboard", "Software", "https://script.google.com/macros/s/AKfycbzOIIzviaY3p1dsf5BZKPie8xskcbia39B-s0jd1DQh8QiWrpgF3KZOHhxwIkgx1hq6/exec?page=userdashboard", "", "", 16],
    ["Yes", "Dashboards, Tasks & HR", "Pending Tasks", "Software", "https://script.google.com/macros/s/AKfycbxMWuT3yAo-uGpPS6WMCuRjefhXW6A6piPCds_-fPCQLJIep126grbLv7gPZm0g5bTJpw/exec", "", "", 17],
    ["Yes", "Dashboards, Tasks & HR", "New Joiner", "Software", "https://script.google.com/macros/s/AKfycbzTrDDPMmDimI4Lat5gjHyBQHAS4eMxOAD8BC0PCQgU2rN5gwPIiqAsPsBiAjdhWLSL/exec", "", "", 18],
    ["Yes", "Dashboards, Tasks & HR", "Employee Exit", "Software", "https://script.google.com/macros/s/AKfycbys4ftf1k-DWOdOF-WGWvRVlw5LT7auZJhCmTHdplo-lk1CBqQz7tafHzVkaAM5GJhz/exec?page=exit", "", "", 19],
    ["Yes", "Dashboards, Tasks & HR", "WhatsApp Sender", "Software", "https://script.google.com/macros/s/AKfycbxsCMIS5Twy1nwePrBHEtcqjfwh4suH6yD7cjhrUeimV4vWsKMvqLtNkIWSghbQyaaHVA/exec", "", "", 20],
    ["Yes", "Dashboards, Tasks & HR", "Aakash ERP", "Software", "https://script.google.com/macros/s/AKfycbxaNj2qle6I9ekGvNnDdnHvEvNS2b2VqvHJfQGWS-KkLz-W8U8AUFoZvT16iOjgeL00/exec", "", "", 21],
    ["Yes", "Websites & Marketplaces", "Business", "Website", "https://script.google.com/macros/s/AKfycbwWMgpcCRGbo13uU6E-uKpLnyM9zDlItFA8YHNHKdGhl2mVWwdMlv56TYAk74nrqovuFw/exec?page=business&id=client1", "", "", 22],
    ["Yes", "Websites & Marketplaces", "E-commerce", "Website", "https://script.google.com/macros/s/AKfycbwWMgpcCRGbo13uU6E-uKpLnyM9zDlItFA8YHNHKdGhl2mVWwdMlv56TYAk74nrqovuFw/exec?page=ecommerce&id=client1", "", "", 23],
    ["Yes", "Websites & Marketplaces", "Restaurant", "Website", "https://script.google.com/macros/s/AKfycbwWMgpcCRGbo13uU6E-uKpLnyM9zDlItFA8YHNHKdGhl2mVWwdMlv56TYAk74nrqovuFw/exec?page=restaurant&id=client1", "", "", 24],
    ["Yes", "Websites & Marketplaces", "Real Estate", "Website", "https://script.google.com/macros/s/AKfycbwWMgpcCRGbo13uU6E-uKpLnyM9zDlItFA8YHNHKdGhl2mVWwdMlv56TYAk74nrqovuFw/exec?page=realestate&id=client1", "", "", 25],
    ["Yes", "Websites & Marketplaces", "Wedding", "Website", "https://script.google.com/macros/s/AKfycbwWMgpcCRGbo13uU6E-uKpLnyM9zDlItFA8YHNHKdGhl2mVWwdMlv56TYAk74nrqovuFw/exec?page=wedding&id=client1", "", "", 26],
    ["Yes", "Websites & Marketplaces", "Event", "Website", "https://script.google.com/macros/s/AKfycbwWMgpcCRGbo13uU6E-uKpLnyM9zDlItFA8YHNHKdGhl2mVWwdMlv56TYAk74nrqovuFw/exec?page=event&id=client1", "", "", 27],
    ["Yes", "Websites & Marketplaces", "Resume", "Website", "https://script.google.com/macros/s/AKfycbwWMgpcCRGbo13uU6E-uKpLnyM9zDlItFA8YHNHKdGhl2mVWwdMlv56TYAk74nrqovuFw/exec?page=resume&id=client1", "", "", 28],
    ["Yes", "Websites & Marketplaces", "Logistics", "Website", "https://script.google.com/macros/s/AKfycbwWMgpcCRGbo13uU6E-uKpLnyM9zDlItFA8YHNHKdGhl2mVWwdMlv56TYAk74nrqovuFw/exec?page=log&id=client1", "", "", 29],
    ["Yes", "Websites & Marketplaces", "Automation Expert Site", "Website", "https://aakashsystemcraftspecialist.github.io/Businessautomationexpert/", "", "", 30],
    ["Yes", "Websites & Marketplaces", "Business Autopilot Site", "Website", "https://business-autopilot.netlify.app/", "", "", 31],
    ["Yes", "Websites & Marketplaces", "Bechlo Supplier", "Software", "https://script.google.com/macros/s/AKfycbwBctLhQlytd11Y5Iq6k83WBI-vYGcMA5kk-lL87mGhuwLAyPAfJCiSwAPVQ5gfGnm-/exec", "", "", 32],
    ["Yes", "Websites & Marketplaces", "Bechlo Customer", "Software", "https://script.google.com/macros/s/AKfycbwBctLhQlytd11Y5Iq6k83WBI-vYGcMA5kk-lL87mGhuwLAyPAfJCiSwAPVQ5gfGnm-/exec?page=customer", "", "", 33],
    ["Yes", "Presentations", "Company PPT", "PPT", "https://docs.google.com/presentation/d/1VLU5iOx3RiVAoph0fIUPGPIC74gaIDHx_peSZjRw3jA/present?slide=id.g2f8c9338f3f_0_0", "", "Services aur systems ka overview", 34]
    ];
    sh.getRange(2, 1, seed.length, PF_HEADERS.length).setValues(seed);
    sh.getRange('A2:A').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Yes','No'], true).build());
    sh.getRange('D2:D').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(PF_TYPES, true).build());
    sh.setColumnWidth(1, 110); sh.setColumnWidth(2, 200); sh.setColumnWidth(3, 220); sh.setColumnWidth(4, 100);
    sh.setColumnWidth(5, 320); sh.setColumnWidth(6, 260); sh.setColumnWidth(7, 280); sh.setColumnWidth(8, 70);
    sh.getRange(1, 1).setNote('Yes = site par dikhega, No = chhup jayega');
    sh.getRange(1, 6).setNote('Screenshot ka Google Drive link (sharing: Anyone with link) ya koi image URL');
    sh.getRange(1, 4).setNote('Video = YouTube / Drive video link. Image ya Video wale items "Showcase" gallery mein aate hain.');
  }
}

function leadSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(LEAD_SHEET) || ss.insertSheet(LEAD_SHEET);
  if (sh.getLastRow() === 0) {
    sh.appendRow(LEAD_HEADERS);
    styleHeader_(sh, LEAD_HEADERS.length);
    sh.getRange('J2:J').setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(['New','Called','Demo Done','Quote Sent','Won','Lost'], true).build());
    sh.setColumnWidths(1, LEAD_HEADERS.length, 140);
  }
  return sh;
}

function styleHeader_(sh, n) {
  sh.getRange(1, 1, 1, n).setFontWeight('bold').setBackground('#1546C9').setFontColor('#FFFFFF');
  sh.setFrozenRows(1);
}

function clean_(v) {
  v = String(v || '').trim().slice(0, 1000);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
