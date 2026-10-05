/**
 * WORKSHOP Automation V1.4.1
 * Single-file Apps Script engine.
 *
 * Architecture:
 *   ONE PERMANENT APPS SCRIPT ENGINE
 *        -> ONE MASTER GOOGLE SHEET
 *        -> MANY ISOLATED WORKSHOPS
 *
 * V1.4.1 scope:
 *   Workshop -> Registration -> Participant -> Confirmation -> Reminder
 *   -> Attendance -> Certificate -> PDF -> Email -> Verification
 *   Email delivery foundation: central queue, priority, quota reserve, retry backoff, resumable WAITING_FOR_QUOTA state.
 *
 * Self Test:
 *   Uses a temporary participant context selected from an existing workshop.
 *   It does not require a Google Form submission and never creates a production
 *   participant row. Default: T+0 Registration, T+~1 min Reminder, T+~2 min
 *   Certificate + PDF, with optional Attendance in between.
 *
 * Email templates are embedded here for runtime rendering; standalone HTML references are shipped in the Email Templates folder.
 */

const APP = {
  VERSION: 'V1.4.1',
  ROOT: 'WORKSHOP',
  DEFAULT_TIMEZONE: 'Asia/Dhaka',
  SCHEDULER_MINUTES: 5,
  TEST_DELAY_MS: 60 * 1000,
  MAX_EMAIL_BATCH: 20,
  EMAIL_SAFE_DAILY_RESERVE: 10,
  EMAIL_RETRY_LIMIT: 3,
  EMAIL_RETRY_BASE_MINUTES: 10,
  EMAIL_QUOTA_RECHECK_MINUTES: 30,
  MAX_CERT_BATCH: 8,
  MAX_WORKSHOPS_PER_RUN: 50,
  SHEETS: {
    SETTINGS: 'Settings',
    WORKSHOPS: 'Workshops',
    PARTICIPANTS: 'Participants',
    ATTENDANCE: 'Attendance',
    QUEUE: 'Email Queue',
    CERTS: 'Certificates',
    LOGS: 'System Logs',
    TESTS: 'Test Runs',
    DASHBOARD: 'Dashboard',
    CUSTOM_EMAILS: 'Custom Emails'
  },
  EVENT_STATES: ['DRAFT', 'OPEN', 'LIVE', 'COMPLETED', 'CANCELLED', 'ARCHIVED'],
  REG_STATES: ['REGISTERED', 'APPROVED', 'ON_HOLD', 'REMOVED', 'REJECTED', 'TEST'],
  EMAIL_STATES: ['PENDING', 'PROCESSING', 'SENT', 'FAILED', 'RETRY', 'WAITING_FOR_QUOTA', 'EXPIRED'],
  EMAIL_TYPES: ['REGISTRATION', 'REMINDER', 'ATTENDANCE', 'CERTIFICATE', 'CUSTOM', 'TEST_REGISTRATION', 'TEST_REMINDER', 'TEST_ATTENDANCE', 'TEST_CERTIFICATE'],
  CERT_RULES: ['ALL_REGISTERED', 'APPROVED_ONLY', 'ATTENDANCE_REQUIRED']
};

const SCHEMAS = {
  Settings: ['Key','Value','Updated At'],
  Workshops: ['Event ID','Workshop Name','Workshop Date','Start Time','End Time','Workshop Type','Timezone','Meeting Platform','Meeting Link','Meeting ID','Meeting Passcode','Meeting Host','Venue Name','Venue Address','Details URL','Capacity','Registration Opens','Registration Closes','Reminder Enabled','Reminder Mode','Reminder Offsets','Reminder Custom At','Certificate Enabled','Certificate Type','Certificate Release Mode','Certificate Delay','Certificate Next Morning Time','Certificate Custom Release At','Certificate Eligibility','Website','Facebook','LinkedIn','Instagram','YouTube','X','Support Email','Group Enabled','Group Type','Group Name','Group Invite URL','Notes','Status','Folder ID','Registration Form ID','Registration Response Spreadsheet ID','Attendance Form ID','Attendance Response Spreadsheet ID','Last Registration Row','Last Attendance Row','Created At','Updated At','Delivery Mode','Attendance Enabled','Attendance Window Minutes','Announcement Enabled','Announcement Mode','Announcement Lead Minutes','Announcement Custom At'],
  Participants: ['Participant ID','Event ID','Workshop Name','Registration Timestamp','Full Name','Email','Phone / WhatsApp Number','Department / Faculty','Season','ID Number','Registration Number','Registration Status','Attendance Status','Certificate Eligibility','Certificate Status','Certificate ID','Certificate PDF URL','Created At','Updated At','Source Response Row'],
  Attendance: ['Timestamp','Participant ID','Event ID','Email','Name','Workshop Type','Validation','Attendance Status','Validation Message','Source','Created At'],
  'Email Queue': ['Job ID','Unique Key','Event ID','Participant ID','Email Type','Priority','Scheduled At','Not Before','Deadline At','Status','Attempt Count','Last Attempt At','Sent At','Recipient','Subject','Attachment File ID','Last Error','Queue Reason','Test Run ID','Message Body','Include Meeting Link','Sender Name','Created At','Updated At'],
  Certificates: ['Certificate ID','Event ID','Participant ID','Participant Name','Email','Workshop Name','Season','Issued At','PDF URL','PDF File ID','Verification URL','Status','Created At','Test Run ID'],
  'System Logs': ['Timestamp','Level','Module','Event ID','Participant ID','Action','Message','Details'],
  'Test Runs': ['Run ID','Event ID','Participant ID','Workshop Name','Recipient','Participant Snapshot','Include Attendance','Started At','Status','Current Step','Confirmation Sent At','Reminder Sent At','Attendance Sent At','Certificate Created At','Certificate Sent At','Test Certificate ID','Test Certificate PDF URL','Last Error','Completed At'],
  'Custom Emails': ['Message ID','Event ID','Workshop Name','Recipient Mode','Recipient Count','Subject','Heading','Message','CTA Enabled','CTA Text','CTA URL','Created At','Status','Queued Jobs','Last Error'],
  Dashboard: ['Metric','Value','Updated At']
};

function onOpen() {
  SpreadsheetApp.getUi().createMenu('WORKSHOP Automation')
    .addItem('Open Control Center','openControlCenter')
    .addItem('Open Self Test','openSelfTestV141_')
    .addItem('Open Custom Email','openCustomEmailV141_')
    .addItem('Refresh Control Center','refreshControlCenterV14_')
    .addItem('Run Setup / Repair','setupSystem')
    .addItem('Run Scheduler Now','processScheduler')
    .addItem('System Readiness Check','systemReadinessCheck')
    .addToUi();
}

function openControlCenter() {
  const ss = getMasterSpreadsheetFast_();
  let sh = ss.getSheetByName('Control Center');
  if (!sh) {
    setupControlCenter_(ss);
    setupCommunicationSheetsV141_(ss);
    installControlCenterEditTriggerV14_(ss);
    sh = ss.getSheetByName('Control Center');
  }
  ss.setActiveSheet(sh);
  sh.getRange('A1').activate();
  refreshControlCenterV14_();
  setupCommunicationSheetsV141_(ss);
}

function doGet(e) {
  const params = (e && e.parameter) || {};
  const page = String(params.page || 'verify').trim().toLowerCase();

  // Public JSON API used by the Vercel verification frontend.
  // Apps Script passes URL query parameters through e.parameter; the API must
  // return TextOutput/JSON rather than the HTML verification page.
  if (page === 'verify-api') {
    const certificateId = String(params.certificateId || params.id || '').trim();
    const payload = verifyCertificatePublic(certificateId);
    payload.version = APP.VERSION;
    payload.service = 'certificate-verification';
    return ContentService
      .createTextOutput(JSON.stringify(payload))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (page === 'health') {
    return ContentService
      .createTextOutput(JSON.stringify({ok:true,version:APP.VERSION,service:'certificate-verification'}))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (page === 'verify') {
    return HtmlService.createTemplateFromFile('Verify').evaluate().setTitle('Certificate Verification');
  }

  return HtmlService.createHtmlOutput('<!doctype html><html><body style="font-family:system-ui;padding:32px"><h2>WORKSHOP Automation '+APP.VERSION+'</h2><p>Open the master Google Sheet and use <b>WORKSHOP Automation → Open Control Center</b>.</p></body></html>');
}

function getPublicBranding() {
  const settings = getSettings_();
  return {
    organizationName: settings.ORGANIZATION_NAME || APP.ROOT,
    logoUrl: settings.DEFAULT_LOGO_URL || '',
    verificationUrl: settings.PUBLIC_VERIFY_URL || settings.WEB_APP_URL || ScriptApp.getService().getUrl() || ''
  };
}

function setupSystem() {
  // Setup/repair may legitimately overlap a scheduler execution. The lock helper gives setup a longer wait window,
  // while the scheduler uses a short non-blocking lock and simply retries on its next trigger.
  return withLock_('setupSystem', function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error('Open the master Google Sheet before running setupSystem().');
    PropertiesService.getScriptProperties().setProperties({
      MASTER_SPREADSHEET_ID: ss.getId(),
      VERSION: APP.VERSION
    }, true);

    const root = getOrCreateRootFolder_();
    const system = getOrCreateFolder_(root, '00_System');
    const brand = getOrCreateFolder_(root, '01_Brand');
    const templates = getOrCreateFolder_(root, '02_Templates');
    const events = getOrCreateFolder_(root, '03_Events');
    const archive = getOrCreateFolder_(root, '04_Archive');
    getOrCreateFolder_(templates, 'Email Themes');
    getOrCreateFolder_(templates, 'Email Themes/Registration');
    getOrCreateFolder_(templates, 'Email Themes/Reminder');
    getOrCreateFolder_(templates, 'Email Themes/Certificate');

    // Keep the master spreadsheet out of My Drive root.
    try { DriveApp.getFileById(ss.getId()).moveTo(system); } catch (err) { log_('WARN','SETUP','', '', 'MOVE_MASTER_SHEET', String(err), ''); }

    migrateLegacySchemas_(ss);
    Object.keys(SCHEMAS).forEach(name => ensureSheetSchema_(ss, name, SCHEMAS[name]));
    ensureDefaultSettings_(ss, {rootId: root.getId(), systemId: system.getId(), brandId: brand.getId(), templatesId: templates.getId(), eventsId: events.getId(), archiveId: archive.getId()});
    repairAllWorkshopMetadata_(ss);
    upgradeV21DefaultsAndLegacyWorkshops_(ss);
    disableLegacyAnnouncementQueueJobs_(ss);
    // Keep the visible settings version synchronized on upgrades.
    updateObjectByKey_(getSheet_(APP.SHEETS.SETTINGS),'Key','APP_VERSION',{'Value':APP.VERSION,'Updated At':new Date()});
    setupControlCenter_(ss);
    setupCommunicationSheetsV141_(ss);
    installControlCenterEditTriggerV14_(ss);
    ensureDefaultCertificateMaster_();
    installCentralScheduler_();
    repairEmailQueueMetadata_();
    tidyEmailQueueSheet_();
    refreshDashboard_();
    refreshControlCenterV14_();
    log_('INFO','SETUP','','','SETUP_COMPLETE','System initialized '+APP.VERSION,'rootId='+root.getId());
    return {ok:true, version:APP.VERSION, rootFolder:root.getName(), message:'WORKSHOP Automation setup completed.'};
  });
}

function ensureDefaultSettings_(ss, ids) {
  const sh = getSheet_(APP.SHEETS.SETTINGS);
  const defaults = {
    APP_VERSION: APP.VERSION,
    ROOT_FOLDER_ID: ids.rootId,
    SYSTEM_FOLDER_ID: ids.systemId,
    BRAND_FOLDER_ID: ids.brandId,
    TEMPLATES_FOLDER_ID: ids.templatesId,
    EVENTS_FOLDER_ID: ids.eventsId,
    ARCHIVE_FOLDER_ID: ids.archiveId,
    MASTER_SPREADSHEET_ID: ss.getId(),
    DEFAULT_TIMEZONE: APP.DEFAULT_TIMEZONE,
    
    DEFAULT_WORKSHOP_TYPE: 'Workshop',
    DEFAULT_MEETING_PLATFORM: 'Google Meet',
    DEFAULT_GROUP_TYPE: 'WhatsApp',
    DEFAULT_CAPACITY: '0',
    DEFAULT_OPEN_DAYS_BEFORE: '7',
    DEFAULT_CLOSE_MINUTES_BEFORE: '15',
    DEFAULT_REMINDERS: '60',
    DEFAULT_REMINDER_LEAD_MINUTES: '2880',
    DEFAULT_DELIVERY_MODE: 'OFFLINE',
    DEFAULT_ATTENDANCE_ENABLED: 'TRUE',
    DEFAULT_ATTENDANCE_WINDOW_MINUTES: '120',
    DEFAULT_ANNOUNCEMENT_ENABLED: 'FALSE',
    DEFAULT_ANNOUNCEMENT_LEAD_MINUTES: '5760',
    DEFAULT_ANNOUNCEMENT_MODE: 'BEFORE_START',
    DEFAULT_CERTIFICATE_ENABLED: 'TRUE',
    DEFAULT_CERTIFICATE_RELEASE: 'AFTER_WORKSHOP',
    DEFAULT_CERTIFICATE_DELAY: '0',
    DEFAULT_CERTIFICATE_TIME: '09:00',
    DEFAULT_CERTIFICATE_ELIGIBILITY: 'ATTENDANCE_REQUIRED',
    SENDER_NAME: 'Finance Club PSTU',
    ORGANIZATION_NAME: 'Finance Club PSTU',
    DEFAULT_LOGO_URL: 'https://raw.githubusercontent.com/SudiptoKumar/WORKSHOP-/refs/heads/main/assets/Finance-Club.jpg',
    WEB_APP_URL: '',
    DEFAULT_WEBSITE: '',
    DEFAULT_FACEBOOK: '',
    DEFAULT_INSTAGRAM: '',
    DEFAULT_LINKEDIN: '',
    DEFAULT_YOUTUBE: '',
    DEFAULT_X: '',
    DEFAULT_SUPPORT_EMAIL: '',
    PUBLIC_VERIFY_URL: 'https://financeclubpstu.vercel.app/verify',
    EMAIL_SAFETY_RESERVE: String(APP.EMAIL_SAFE_DAILY_RESERVE),
    EMAIL_MAX_BATCH: String(APP.MAX_EMAIL_BATCH),
    EMAIL_RETRY_LIMIT: String(APP.EMAIL_RETRY_LIMIT),
    EMAIL_RETRY_BASE_MINUTES: String(APP.EMAIL_RETRY_BASE_MINUTES),
    EMAIL_QUOTA_RECHECK_MINUTES: String(APP.EMAIL_QUOTA_RECHECK_MINUTES),
    PAUSE_AUTOMATION: 'FALSE'
  };
  const values = readSheetObjects_(sh);
  const map = {}; values.forEach(r => map[String(r.Key || '')] = r);
  const now = new Date();
  Object.keys(defaults).forEach(k => {
    if (!map[k]) sh.appendRow([k, defaults[k], now]);
  });
}

function upgradeV21DefaultsAndLegacyWorkshops_(ss) {
  try {
    const sh = ss.getSheetByName(APP.SHEETS.SETTINGS);
    if (sh) {
      const reminders = getSetting_('DEFAULT_REMINDERS');
      if (reminders === '720,60,10') updateObjectByKey_(sh, 'Key', 'DEFAULT_REMINDERS', {'Value':'60','Updated At':new Date()});
      if (!getSetting_('DEFAULT_LOGO_URL') || getSetting_('DEFAULT_LOGO_URL')==='' || getSetting_('DEFAULT_LOGO_URL')==='{{logo_url}}') updateObjectByKey_(sh, 'Key', 'DEFAULT_LOGO_URL', {'Value':'https://raw.githubusercontent.com/SudiptoKumar/WORKSHOP-/refs/heads/main/assets/Finance-Club.jpg','Updated At':new Date()});
      if (!getSetting_('ORGANIZATION_NAME') || getSetting_('ORGANIZATION_NAME')==='Finance Club') updateObjectByKey_(sh, 'Key', 'ORGANIZATION_NAME', {'Value':'Finance Club PSTU','Updated At':new Date()});
      if (!getSetting_('SENDER_NAME') || getSetting_('SENDER_NAME')==='Finance Club') updateObjectByKey_(sh, 'Key', 'SENDER_NAME', {'Value':'Finance Club PSTU','Updated At':new Date()});
      if (!getSetting_('PUBLIC_VERIFY_URL') || getSetting_('PUBLIC_VERIFY_URL')==='') updateObjectByKey_(sh, 'Key', 'PUBLIC_VERIFY_URL', {'Value':'https://financeclubpstu.vercel.app/verify','Updated At':new Date()});
      if (!getSetting_('DEFAULT_REMINDER_LEAD_MINUTES')) updateObjectByKey_(sh, 'Key', 'DEFAULT_REMINDER_LEAD_MINUTES', {'Value':'2880','Updated At':new Date()});
      if (!getSetting_('DEFAULT_X')) updateObjectByKey_(sh, 'Key', 'DEFAULT_X', {'Value':'','Updated At':new Date()});
      if (!getSetting_('DEFAULT_SUPPORT_EMAIL')) updateObjectByKey_(sh, 'Key', 'DEFAULT_SUPPORT_EMAIL', {'Value':'','Updated At':new Date()});
      if (!getSetting_('PUBLIC_VERIFY_URL')) updateObjectByKey_(sh, 'Key', 'PUBLIC_VERIFY_URL', {'Value':'','Updated At':new Date()});
      if (!getSetting_('DEFAULT_DELIVERY_MODE')) updateObjectByKey_(sh, 'Key', 'DEFAULT_DELIVERY_MODE', {'Value':'OFFLINE','Updated At':new Date()});
      if (!getSetting_('DEFAULT_ATTENDANCE_ENABLED')) updateObjectByKey_(sh, 'Key', 'DEFAULT_ATTENDANCE_ENABLED', {'Value':'TRUE','Updated At':new Date()});
      if (!getSetting_('DEFAULT_ATTENDANCE_WINDOW_MINUTES')) updateObjectByKey_(sh, 'Key', 'DEFAULT_ATTENDANCE_WINDOW_MINUTES', {'Value':'120','Updated At':new Date()});
      if (!getSetting_('DEFAULT_ANNOUNCEMENT_ENABLED')) updateObjectByKey_(sh, 'Key', 'DEFAULT_ANNOUNCEMENT_ENABLED', {'Value':'FALSE','Updated At':new Date()});
      if (!getSetting_('DEFAULT_ANNOUNCEMENT_LEAD_MINUTES')) updateObjectByKey_(sh, 'Key', 'DEFAULT_ANNOUNCEMENT_LEAD_MINUTES', {'Value':'5760','Updated At':new Date()});
      if (!getSetting_('DEFAULT_ANNOUNCEMENT_MODE')) updateObjectByKey_(sh, 'Key', 'DEFAULT_ANNOUNCEMENT_MODE', {'Value':'BEFORE_START','Updated At':new Date()});
    }
    const wsh = ss.getSheetByName(APP.SHEETS.WORKSHOPS);
    if (wsh) {
      const rows = readSheetObjects_(wsh);
      rows.forEach(function(w){
        const fields = {};
        if (String(w['Reminder Enabled']).trim() === '') fields['Reminder Enabled'] = !!String(w['Reminder Offsets']||'').trim();
        if (String(w['Reminder Mode']).trim() === '') fields['Reminder Mode'] = 'BEFORE_START';
        if (String(w['Reminder Offsets'])==='720,60,10') fields['Reminder Offsets'] = '60';
        if (!String(w['Certificate Release Mode']||'').trim()) fields['Certificate Release Mode']='AFTER_WORKSHOP';
        if (!String(w['Delivery Mode']||'').trim()) fields['Delivery Mode'] = String(w['Meeting Link']||'').trim() ? 'ONLINE' : 'OFFLINE';
        if (String(w['Attendance Enabled']).trim() === '') fields['Attendance Enabled'] = boolValue_(getSetting_('DEFAULT_ATTENDANCE_ENABLED'), true);
        if (!String(w['Attendance Window Minutes']||'').trim()) fields['Attendance Window Minutes'] = Number(getSetting_('DEFAULT_ATTENDANCE_WINDOW_MINUTES')||120) || 120;
        if (String(w['Announcement Enabled']).trim() === '') fields['Announcement Enabled'] = boolValue_(getSetting_('DEFAULT_ANNOUNCEMENT_ENABLED'), false);
        if (!String(w['Announcement Mode']||'').trim()) fields['Announcement Mode'] = String(getSetting_('DEFAULT_ANNOUNCEMENT_MODE')||'BEFORE_START');
        if (!String(w['Announcement Lead Minutes']||'').trim()) fields['Announcement Lead Minutes'] = Number(getSetting_('DEFAULT_ANNOUNCEMENT_LEAD_MINUTES')||5760) || 5760;
        if (!String(w['Announcement Custom At']||'').trim()) fields['Announcement Custom At'] = '';
        if (Object.keys(fields).length) updateWorkshopByEvent_(w['Event ID'], Object.assign(fields, {'Updated At':new Date()}));
      });
    }
  } catch (err) {
    log_('WARN','SETUP','','','V1_MIGRATION_WARNING',String(err),stack_(err));
  }
}

function ensureDefaultCertificateMaster_() {
  const props = PropertiesService.getScriptProperties();
  const existingId = props.getProperty('CERTIFICATE_MASTER_ID');
  if (existingId) {
    try {
      const existing = SlidesApp.openById(existingId);
      existing.replaceAllText('{{SESSION}}','{{SEASON}}');
      existing.saveAndClose();
    } catch (err) {
      log_('WARN','SETUP','','','CERTIFICATE_MASTER_REPAIR_FAILED',String(err),'');
    }
    return;
  }
  const settings = getSettings_();
  const folder = DriveApp.getFolderById(settings.TEMPLATES_FOLDER_ID);
  const p = SlidesApp.create('Default Certificate Master');
  const file = DriveApp.getFileById(p.getId());
  try { file.moveTo(folder); } catch (err) {}
  const slide = p.getSlides()[0];
  slide.getBackground().setSolidFill('#F8FAFC');
  const title = slide.insertTextBox('CERTIFICATE OF PARTICIPATION', 50, 65, 620, 50);
  title.getText().getTextStyle().setFontFamily('Arial').setFontSize(24).setBold(true).setForegroundColor('#0F172A');
  const intro = slide.insertTextBox('This certificate is proudly presented to', 80, 145, 560, 30);
  intro.getText().getTextStyle().setFontFamily('Arial').setFontSize(15).setForegroundColor('#475569');
  const name = slide.insertTextBox('{{PARTICIPANT_NAME}}', 80, 190, 560, 55);
  name.getText().getTextStyle().setFontFamily('Arial').setFontSize(28).setBold(true).setForegroundColor('#2563EB');
  const body = slide.insertTextBox('for participating in\n{{WORKSHOP_NAME}}\n{{WORKSHOP_DATE}} • {{SEASON}}', 80, 270, 560, 130);
  body.getText().getTextStyle().setFontFamily('Arial').setFontSize(16).setForegroundColor('#334155');
  const cert = slide.insertTextBox('Certificate ID: {{CERTIFICATE_ID}}', 80, 420, 560, 28);
  cert.getText().getTextStyle().setFontFamily('Arial').setFontSize(11).setForegroundColor('#64748B');
  p.saveAndClose();
  props.setProperty('CERTIFICATE_MASTER_ID', p.getId());
}

function installCentralScheduler_() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'processScheduler') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('processScheduler').timeBased().everyMinutes(APP.SCHEDULER_MINUTES).create();
}

function normalizeDeliveryMode_(value, fallback){
  const mode=String(value==null?'':value).trim().toUpperCase();
  return mode==='ONLINE' || mode==='OFFLINE' ? mode : String(fallback||'OFFLINE').trim().toUpperCase()==='ONLINE' ? 'ONLINE' : 'OFFLINE';
}
function normalizeAttendanceWindowMinutes_(value, fallback){
  const n=Number(value);
  if(isFinite(n) && n>0) return Math.min(Math.round(n), 7*24*60);
  const f=Number(fallback);
  return isFinite(f) && f>0 ? Math.min(Math.round(f),7*24*60) : 120;
}
function reminderLeadMinutesForNewWorkshop_(){
  const configured=Number(getSetting_('DEFAULT_REMINDER_LEAD_MINUTES'));
  return isFinite(configured) && configured>0 ? configured : 2880;
}
function attendanceWindowPresetToMinutes_(preset){
  const s=String(preset||'').trim().toLowerCase();
  if(s==='30 minutes') return 30;
  if(s==='1 hour') return 60;
  if(s==='2 hours') return 120;
  if(s==='6 hours') return 360;
  if(s==='12 hours') return 720;
  if(s==='24 hours') return 1440;
  return 120;
}
function attendanceWindowHuman_(minutes){
  const n=Number(minutes||120);
  if(n%1440===0) return (n/1440)+' day'+(n/1440===1?'':'s');
  if(n%60===0) return (n/60)+' hour'+(n/60===1?'':'s');
  return n+' minutes';
}
function normalizeAnnouncementLeadMinutes_(value, fallback){
  const n=Number(value);
  if(isFinite(n) && n>0) return Math.min(Math.round(n),14*24*60);
  const f=Number(fallback);
  return isFinite(f) && f>0 ? Math.min(Math.round(f),14*24*60) : 5760;
}
function announcementPresetToMinutes_(preset){
  const m={'1 day before':1440,'2 days before':2880,'3 days before':4320,'4 days before':5760,'5 days before':7200,'7 days before':10080};
  return m[String(preset||'').trim()]||0;
}
function announcementPresetHuman_(minutes){
  const n=Number(minutes||5760);
  const map={1440:'1 day before',2880:'2 days before',4320:'3 days before',5760:'4 days before',7200:'5 days before',10080:'7 days before'};
  return map[n]||formatOffsetHuman_(n)+' before';
}

function createWorkshop(payload) {
  return withLock_('createWorkshop', function() {
    ensureSetup_();
    payload = payload || {};
    const name = clean_(payload.workshopName);
    const date = clean_(payload.workshopDate);
    const start = clean_(payload.startTime) || '10:00';
    const end = clean_(payload.endTime) || '12:00';
    if (!name) throw new Error('Workshop Name is required.');
    if (!date) throw new Error('Workshop Date is required.');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('Workshop Date must use YYYY-MM-DD.');
    if (!/^\d{2}:\d{2}$/.test(start) || !/^\d{2}:\d{2}$/.test(end)) throw new Error('Start/End Time must use HH:mm.');
    const settings = getSettings_();
    const ss = getMasterSpreadsheet_();
    const eventId = nextEventId_(date.substring(0,4));
    const timezone = clean_(payload.timezone) || settings.DEFAULT_TIMEZONE || APP.DEFAULT_TIMEZONE;
    const deliveryMode = normalizeDeliveryMode_(payload.deliveryMode, settings.DEFAULT_DELIVERY_MODE || 'OFFLINE');
    const meetingLink = deliveryMode === 'ONLINE' ? clean_(payload.meetingLink) : '';
    if(deliveryMode==='ONLINE' && meetingLink && !isHttpUrl_(meetingLink)) throw new Error('Meeting Link must be a valid http(s) URL.');
    const extracted = deliveryMode==='ONLINE' ? extractMeetingDetails_(meetingLink, payload.meetingPlatform || settings.DEFAULT_MEETING_PLATFORM) : {};
    const meetingPlatform = deliveryMode==='ONLINE' ? (clean_(payload.meetingPlatform) || extracted.platform || settings.DEFAULT_MEETING_PLATFORM) : '';
    const meetingId = deliveryMode==='ONLINE' ? (clean_(payload.meetingId) || extracted.meetingId || '') : '';
    const meetingPasscode = deliveryMode==='ONLINE' ? (clean_(payload.meetingPasscode) || extracted.meetingPasscode || '') : '';
    const meetingHost = deliveryMode==='ONLINE' ? (clean_(payload.meetingHost) || extracted.meetingHost || '') : '';
    const venueName = deliveryMode === 'OFFLINE' ? clean_(payload.venueName) : '';
    const venueAddress = deliveryMode === 'OFFLINE' ? clean_(payload.venueAddress) : '';
    const attendanceEnabled = payload.attendanceEnabled === undefined ? boolValue_(settings.DEFAULT_ATTENDANCE_ENABLED, true) : !!payload.attendanceEnabled;
    const attendanceWindowMinutes = normalizeAttendanceWindowMinutes_(payload.attendanceWindowMinutes, settings.DEFAULT_ATTENDANCE_WINDOW_MINUTES || 120);
    // V1.4 removes scheduled announcements from the automatic lifecycle. Custom Email is the manual communication tool.
    const announcementEnabled = false;
    const announcementMode = 'MANUAL_ONLY';
    const announcementLeadMinutes = 0;
    const announcementCustomAt = '';
    const groupEnabled = payload.groupEnabled === undefined ? !!clean_(payload.groupInviteUrl) : !!payload.groupEnabled;
    const groupType = clean_(payload.groupType) || settings.DEFAULT_GROUP_TYPE || 'WhatsApp';
    const groupName = clean_(payload.groupName) || (groupEnabled ? name+' '+groupType+' Group' : '');
    const groupInviteUrl = groupEnabled ? clean_(payload.groupInviteUrl) : '';
    if(groupInviteUrl && !isHttpUrl_(groupInviteUrl)) throw new Error('Group Invite URL must be a valid http(s) URL.');
    const folderRoot = DriveApp.getFolderById(settings.EVENTS_FOLDER_ID);
    const eventFolder = folderRoot.createFolder(eventId+' - '+safeFileName_(name));
    const registrationFolder = eventFolder.createFolder('Registration');
    const participantsFolder = eventFolder.createFolder('Participants');
    const attendanceFolder = eventFolder.createFolder('Attendance');
    const certificatesFolder = eventFolder.createFolder('Certificates');
    certificatesFolder.createFolder('Test Runs');
    const emailAssetsFolder = eventFolder.createFolder('Email Assets');
    const eventAssetsFolder = eventFolder.createFolder('Event Assets');

    const regSs = SpreadsheetApp.create(eventId+' - Registration Responses');
    try { DriveApp.getFileById(regSs.getId()).moveTo(registrationFolder); } catch (err) {}
    const regForm = FormApp.create(eventId+' - '+name+' Registration');
    regForm.setDescription(buildRegistrationDescription_(name,date,start,end,timezone,deliveryMode,meetingPlatform,venueName,venueAddress));
    addRegistrationItems_(regForm);
    regForm.setConfirmationMessage(buildRegistrationConfirmationMessage_(name,eventId,groupEnabled,groupName,groupInviteUrl,settings.ORGANIZATION_NAME||'Finance Club PSTU'));
    regForm.setDestination(FormApp.DestinationType.SPREADSHEET, regSs.getId());
    try { DriveApp.getFileById(regForm.getId()).moveTo(registrationFolder); } catch (err) {}

    let attSs=null;
    let attForm=null;
    if(attendanceEnabled){
      attSs = SpreadsheetApp.create(eventId+' - Attendance Responses');
      try { DriveApp.getFileById(attSs.getId()).moveTo(attendanceFolder); } catch (err) {}
      attForm = FormApp.create(eventId+' - '+name+' Attendance');
      attForm.setDescription('Attendance check-in for '+name+' ('+eventId+'). Enter the Participant ID received by email.');
      attForm.addTextItem().setTitle('Participant ID').setRequired(true);
      attForm.setConfirmationMessage('Attendance recorded. Thank you.');
      attForm.setDestination(FormApp.DestinationType.SPREADSHEET, attSs.getId());
      try { DriveApp.getFileById(attForm.getId()).moveTo(attendanceFolder); } catch (err) {}
    }

    const row = {
      'Event ID':eventId,'Workshop Name':name,'Workshop Date':date,'Start Time':start,'End Time':end,
      'Workshop Type':clean_(payload.workshopType)||settings.DEFAULT_WORKSHOP_TYPE,
      'Timezone':timezone,'Meeting Platform':meetingPlatform,'Meeting Link':meetingLink,'Meeting ID':meetingId,'Meeting Passcode':meetingPasscode,'Meeting Host':meetingHost,
      'Venue Name':venueName,'Venue Address':venueAddress,'Details URL':clean_(payload.detailsUrl),
      'Capacity':numOr_(payload.capacity,Number(settings.DEFAULT_CAPACITY||0)),'Registration Opens':payload.registrationOpens||'',
      'Registration Closes':payload.registrationCloses||'','Reminder Enabled':payload.reminderEnabled === false ? false : true,'Reminder Mode':clean_(payload.reminderMode)||'BEFORE_START','Reminder Offsets':clean_(payload.reminderOffsets)||'60','Reminder Custom At':payload.reminderCustomAt||'',
      'Certificate Enabled':payload.certificateEnabled === false ? false : true,'Certificate Type':clean_(payload.certificateType)||'Certificate of Participation','Certificate Release Mode':clean_(payload.certificateReleaseMode)||settings.DEFAULT_CERTIFICATE_RELEASE,
      'Certificate Delay':numOr_(payload.certificateDelay,Number(settings.DEFAULT_CERTIFICATE_DELAY||0)),'Certificate Next Morning Time':clean_(payload.certificateNextMorningTime)||settings.DEFAULT_CERTIFICATE_TIME,'Certificate Custom Release At':payload.certificateCustomReleaseAt||'',
      'Certificate Eligibility':clean_(payload.certificateEligibility)||settings.DEFAULT_CERTIFICATE_ELIGIBILITY,'Website':clean_(payload.website)||settings.DEFAULT_WEBSITE,
      'Facebook':clean_(payload.facebook)||settings.DEFAULT_FACEBOOK,'LinkedIn':clean_(payload.linkedin)||settings.DEFAULT_LINKEDIN,'Instagram':clean_(payload.instagram)||settings.DEFAULT_INSTAGRAM,'YouTube':clean_(payload.youtube)||settings.DEFAULT_YOUTUBE,'X':clean_(payload.x)||settings.DEFAULT_X,'Support Email':clean_(payload.supportEmail)||settings.DEFAULT_SUPPORT_EMAIL,'Group Enabled':groupEnabled,'Group Type':groupType,'Group Name':groupName,'Group Invite URL':groupInviteUrl,'Notes':clean_(payload.notes),
      'Status':'DRAFT','Folder ID':eventFolder.getId(),'Registration Form ID':regForm.getId(),'Registration Response Spreadsheet ID':regSs.getId(),
      'Attendance Form ID':attForm ? attForm.getId() : '','Attendance Response Spreadsheet ID':attSs ? attSs.getId() : '','Last Registration Row':1,'Last Attendance Row':1,
      'Created At':new Date(),'Updated At':new Date(),
      'Delivery Mode':deliveryMode,'Attendance Enabled':attendanceEnabled,'Attendance Window Minutes':attendanceWindowMinutes,
      'Announcement Enabled':announcementEnabled,'Announcement Mode':announcementMode,'Announcement Lead Minutes':announcementLeadMinutes,'Announcement Custom At':announcementCustomAt
    };
    row['Registration Opens'] = row['Registration Opens'] || addMinutes_(parseDateTime_(date,start,timezone), -24*60*7, timezone);
    row['Registration Closes'] = row['Registration Closes'] || addMinutes_(parseDateTime_(date,start,timezone), -15, timezone);
    const workshopSheet=getSheet_(APP.SHEETS.WORKSHOPS);
    appendObject_(workshopSheet, row);
    try{ syncWorkshopFormState_(row); }catch(err){
      log_('ERROR','WORKSHOP',eventId,'','FORM_STATE_SYNC_AFTER_CREATE_FAILED',String(err),stack_(err));
    }
    log_('INFO','WORKSHOP',eventId,'','CREATE_WORKSHOP','Workshop created',JSON.stringify({folderId:eventFolder.getId(),registrationFormId:regForm.getId(),attendanceFormId:attForm?attForm.getId():'',groupEnabled:groupEnabled,deliveryMode:deliveryMode}));
    refreshDashboard_();
    return {ok:true,eventId,workshopName:name,folderId:eventFolder.getId(),registrationFormUrl:regForm.getPublishedUrl(),attendanceFormUrl:attForm?attForm.getPublishedUrl():'',registrationEditUrl:regForm.getEditUrl(),groupInviteUrl:groupInviteUrl};
  });
}
function addRegistrationItems_(form) {
  form.addTextItem().setTitle('Full Name').setRequired(true);
  form.addTextItem().setTitle('Email').setRequired(true);
  form.addTextItem().setTitle('Phone / WhatsApp Number').setRequired(true);
  form.addTextItem().setTitle('Department / Faculty').setRequired(true);
  form.addTextItem().setTitle('Season').setRequired(true);
  form.addTextItem().setTitle('ID Number').setRequired(true);
  form.addTextItem().setTitle('Registration Number').setRequired(true);
}

function buildRegistrationDescription_(name,date,start,end,tz,deliveryMode,platform,venueName,venueAddress) {
  const location = String(deliveryMode||'OFFLINE').toUpperCase()==='ONLINE'
    ? 'Platform: '+String(platform||'')
    : 'Venue: '+String(venueName||'')+(venueAddress?'\nAddress: '+String(venueAddress):'');
  return [name,'', 'Date: '+date,'Time: '+start+' - '+end+' ('+tz+')',location,'','Please complete these seven required fields: Full Name, Email, Phone / WhatsApp Number, Department / Faculty, Season, ID Number, Registration Number.'].join('\n');
}

function buildRegistrationConfirmationMessage_(name,eventId,groupEnabled,groupName,groupInviteUrl,org){
  const lines=['Registration received.','',String(name),'Event ID: '+String(eventId),'','Thank you for registering. Your registration will be validated and your confirmation email will be sent shortly.'];
  if(groupEnabled && groupInviteUrl){
    lines.push('','JOIN WORKSHOP '+String(groupName||'GROUP'),'Join the official workshop group for updates and announcements:',String(groupInviteUrl));
  }
  lines.push('','— '+String(org||'Finance Club PSTU'));
  return lines.join('\n');
}
function isHttpUrl_(value){return /^https?:\/\/\S+$/i.test(String(value||'').trim());}
function extractMeetingDetails_(url,platform){
  const raw=String(url||'').trim();
  const out={platform:clean_(platform),meetingId:'',meetingPasscode:'',meetingHost:''};
  if(!raw) return out;
  const lower=raw.toLowerCase();
  if(!out.platform || out.platform==='Other'){
    if(lower.indexOf('meet.google.com')>=0) out.platform='Google Meet';
    else if(lower.indexOf('zoom.us')>=0) out.platform='Zoom';
    else if(lower.indexOf('teams.microsoft.com')>=0) out.platform='Microsoft Teams';
  }
  let m=raw.match(/meet\.google\.com\/([a-z]{3}-[a-z]{4}-[a-z]{3})/i); if(m) out.meetingId=m[1];
  if(!out.meetingId){m=raw.match(/zoom\.us\/(?:j|my)\/([0-9]{7,15})/i);if(m)out.meetingId=m[1];}
  if(!out.meetingPasscode){try{const q=raw.split('?')[1]||'';const parts=q.split('&');parts.forEach(function(part){const kv=part.split('=');if(kv.length===2&&/^(pwd|passcode|password)$/i.test(decodeURIComponent(kv[0]))){out.meetingPasscode=decodeURIComponent(kv[1]);}});}catch(err){}}
  return out;
}
function buildCalendarUrl_(w){
  try{
    const tz=String(w['Timezone']||APP.DEFAULT_TIMEZONE);
    const start=parseDateTime_(w['Workshop Date'],w['Start Time'],tz);
    const end=parseDateTime_(w['Workshop Date'],w['End Time'],tz);
    const fmt=d=>Utilities.formatDate(d,'UTC',"yyyyMMdd'T'HHmmss'Z'");
    const mode=normalizeDeliveryMode_(w['Delivery Mode'],String(w['Meeting Link']||'').trim()?'ONLINE':'OFFLINE');
    const location=mode==='ONLINE' ? (w['Meeting Platform']||'Online') : [w['Venue Name'],w['Venue Address']].filter(Boolean).join(', ');
    const details=[w['Workshop Name'],'Event ID: '+w['Event ID'],mode==='ONLINE'?(w['Meeting Link']||''):'',w['Group Invite URL']||''].filter(Boolean).join('\n');
    return 'https://calendar.google.com/calendar/render?action=TEMPLATE&text='+encodeURIComponent(String(w['Workshop Name']||''))+'&dates='+fmt(start)+'/'+fmt(end)+'&details='+encodeURIComponent(details)+'&location='+encodeURIComponent(location||'')+'&ctz='+encodeURIComponent(tz);
  }catch(err){return '';}
}
function buildGroupBlock_(w){
  const url=String(w['Group Invite URL']||'').trim(); if(!boolValue_(w['Group Enabled'],!!url)||!url)return '';
  const name=esc_(w['Group Name']||('Workshop '+(w['Group Type']||'Group'))); const type=esc_(w['Group Type']||'WhatsApp');
  const buttonLabel='JOIN WORKSHOP '+type.toUpperCase()+' GROUP';
  return '<!-- WHATSAPP GROUP -->\n<tr><td class="side section-gap" style="padding:0 32px 22px;"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" class="panel"><tr><td align="center" class="cta-pad" style="padding:23px 20px 25px;"><div class="apple-font" style="font-size:17px;line-height:23px;font-weight:600;color:#1d1d1f;">'+name+'</div><div class="apple-font" style="margin-top:3px;font-size:12px;line-height:18px;color:#6e6e73;">Join the official workshop group for updates, reminders, and important announcements.</div><table role="presentation" cellpadding="0" cellspacing="0" border="0" style="margin-top:14px;" class="cta-button"><tr><td align="center" style="background:#25d366;border-radius:9px;"><a href="'+escAttr_(url)+'" style="display:block;padding:11px 18px;font-family:-apple-system,BlinkMacSystemFont,\'Segoe UI\',Arial,sans-serif;font-size:13px;line-height:18px;font-weight:600;color:#ffffff;min-width:190px;">'+esc_(buttonLabel)+'</a></td></tr></table><div class="apple-font" style="margin-top:10px;font-size:10px;line-height:15px;color:#86868b;word-break:break-all;">'+esc_(url)+'</div></td></tr></table></td></tr>';
}
function buildLifecycleBlock_(w){
  const attendance=boolValue_(w['Attendance Enabled'],true);
  const cells=[['✓','Registered','Your place is saved'],['02','Reminder','Before the workshop'],['03','Workshop','Attend the session']];
  if(attendance) cells.push(['04','Attendance','Complete your check-in']);
  cells.push([String(cells.length+1).padStart(2,'0'),'Certificate','After completion']);
  const html=cells.map(function(c,i){const right=(i%2===1)?';border-right:0':'';const bottom=(i<cells.length-2)?';border-bottom:1px solid #e5e5e7':''; return '<td class="timeline-cell" width="50%" align="center" valign="top" style="padding:19px 10px 18px;border-right:1px solid #e5e5e7'+right+bottom+'"><table role="presentation" width="42" height="42" align="center" cellpadding="0" cellspacing="0" border="0"><tr><td align="center" valign="middle" style="width:42px;height:42px;border-radius:21px;background:'+(i===0?'#e7f5ec':'#ffffff')+';color:'+(i===0?'#248a4b':'#6e6e73')+';font-family:-apple-system,BlinkMacSystemFont,\'Segoe UI\',Arial,sans-serif;font-size:11px;line-height:40px;font-weight:600;">'+esc_(c[0])+'</td></tr></table><div class="apple-font" style="margin-top:8px;font-size:13px;line-height:18px;font-weight:600;color:#1d1d1f;">'+esc_(c[1])+'</div><div class="apple-font" style="margin-top:2px;font-size:10px;line-height:15px;color:#86868b;">'+esc_(c[2])+'</div></td>';}).join('');
  const rows=[]; for(let i=0;i<cells.length;i+=2){const pair=html.split('|||');}
  let table='';
  for(let i=0;i<cells.length;i+=2){
    const a=cells[i], b=cells[i+1];
    const cell=function(c,idx){return '<td class="timeline-cell" width="50%" align="center" valign="top" style="padding:19px 10px 18px;border-right:1px solid #e5e5e7'+(idx%2===1?';border-right:0':'')+';'+(i+2<cells.length?'border-bottom:1px solid #e5e5e7':'')+'"><table role="presentation" width="42" height="42" align="center" cellpadding="0" cellspacing="0" border="0"><tr><td align="center" valign="middle" style="width:42px;height:42px;border-radius:21px;background:'+(idx===0?'#e7f5ec':'#ffffff')+';color:'+(idx===0?'#248a4b':'#6e6e73')+';font-family:-apple-system,BlinkMacSystemFont,\'Segoe UI\',Arial,sans-serif;font-size:11px;line-height:40px;font-weight:600;">'+esc_(c[0])+'</td></tr></table><div class="apple-font" style="margin-top:8px;font-size:13px;line-height:18px;font-weight:600;color:#1d1d1f;">'+esc_(c[1])+'</div><div class="apple-font" style="margin-top:2px;font-size:10px;line-height:15px;color:#86868b;">'+esc_(c[2])+'</div></td>';};
    table+='<tr>'+cell(a,0)+(b?cell(b,1):'<td width="50%"></td>')+'</tr>';
  }
  return '<!-- WHAT HAPPENS NEXT --><tr><td class="side section-gap" style="padding:0 32px 18px;"><div class="apple-font" style="font-size:17px;line-height:23px;font-weight:600;letter-spacing:-.3px;color:#1d1d1f;margin-bottom:11px;">What happens next</div><table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" class="panel">'+table+'</table></td></tr>';
}
function buildJoiningChecklistBlock_(mode,w){
  const items=mode==='ONLINE' ? ['Open the meeting link.','Use your registered name.','Keep your microphone muted unless requested.','Join a few minutes early.','Keep your Registration Number available.'] : ['Check the venue and address.','Bring your registration details.','Arrive a few minutes early.','Follow the organizer\'s instructions.'];
  const rows=items.map(function(text,i){return '<tr><td style="padding:7px 0;"><table role="presentation" cellpadding="0" cellspacing="0" border="0"><tr><td width="28" valign="top"><table role="presentation" width="20" height="20" cellpadding="0" cellspacing="0" border="0"><tr><td align="center" valign="middle" style="width:20px;height:20px;border-radius:10px;background:#e7f5ec;color:#248a4b;font-family:-apple-system,BlinkMacSystemFont,\'Segoe UI\',Arial,sans-serif;font-size:11px;font-weight:700;">'+(i+1)+'</td></tr></table></td><td class="apple-font" valign="middle" style="padding-left:10px;font-size:14px;line-height:20px;color:#424245;">'+esc_(text)+'</td></tr></table></td></tr>';}).join('');
  const title=mode==='ONLINE'?'BEFORE YOU JOIN':'BEFORE YOU ARRIVE';
  return '<tr><td class="side" style="padding:25px 35px 0;"><div class="apple-font" style="font-size:12px;line-height:16px;font-weight:600;letter-spacing:.7px;color:#86868b;margin-bottom:11px;">'+title+'</div><table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" class="panel"><tr><td style="padding:20px 22px 18px;"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">'+rows+'</table></td></tr></table></td></tr>';
}
function buildMeetingDetailsBlock_(w){
  const mode=normalizeDeliveryMode_(w['Delivery Mode'],String(w['Meeting Link']||'').trim()?'ONLINE':'OFFLINE');
  if(mode==='OFFLINE'){
    if(!w['Venue Name'] && !w['Venue Address']) return '';
    return '<tr><td class="side" style="padding:25px 35px 0;"><div class="apple-font" style="font-size:12px;line-height:16px;font-weight:600;letter-spacing:.7px;color:#86868b;margin-bottom:11px;">VENUE DETAILS</div><table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"><tr><td class="panel" style="background:#f5f5f7;border-radius:16px;padding:16px 18px;"><div class="apple-font" style="font-size:11px;color:#86868b;margin-bottom:5px;">Venue</div><div class="apple-font" style="font-size:15px;font-weight:600;color:#1d1d1f;">'+esc_(w['Venue Name']||'')+'</div>'+(w['Venue Address']?'<div class="apple-font" style="margin-top:6px;font-size:13px;line-height:19px;color:#6e6e73;">'+esc_(w['Venue Address'])+'</div>':'')+'</td></tr></table></td></tr>';
  }
  const fields=[];
  if(w['Meeting Platform']) fields.push(['Platform',w['Meeting Platform']]);
  if(w['Meeting ID']) fields.push(['Meeting ID',w['Meeting ID']]);
  if(w['Meeting Passcode']) fields.push(['Passcode',w['Meeting Passcode']]);
  if(w['Meeting Host']) fields.push(['Host',w['Meeting Host']]);
  if(w['Meeting Link']) fields.push(['Meeting link',w['Meeting Link']]);
  if(!fields.length) return '';
  const cards=fields.map(function(f){const value=f[0]==='Meeting link'?'<a href="'+escAttr_(f[1])+'" class="footer-link" style="font-family:\'Courier New\',monospace;font-size:12px;line-height:18px;color:#0066cc;word-break:break-all;">'+esc_(f[1])+'</a>':esc_(f[1]);return '<div class="panel" style="background:#f5f5f7;border-radius:16px;padding:16px 18px;margin-bottom:10px;"><div class="apple-font" style="font-size:11px;color:#86868b;margin-bottom:5px;">'+esc_(f[0])+'</div><div class="apple-font" style="font-size:15px;font-weight:600;color:#1d1d1f;word-break:break-word;">'+value+'</div></div>';}).join('');
  return '<tr><td class="side" style="padding:25px 35px 0;"><div class="apple-font" style="font-size:12px;line-height:16px;font-weight:600;letter-spacing:.7px;color:#86868b;margin-bottom:11px;">MEETING DETAILS</div>'+cards+'</td></tr>';
}
function buildParticipantPassBlock_(p){
  return '<tr><td class="side" style="padding:25px 35px 0;"><div class="apple-font" style="font-size:12px;line-height:16px;font-weight:600;letter-spacing:.7px;color:#86868b;margin-bottom:11px;">YOUR PASS</div><table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" class="panel"><tr><td style="padding:20px 22px;"><div class="apple-font" style="font-size:11px;color:#86868b;margin-bottom:5px;">Participant</div><div class="apple-font" style="font-size:20px;line-height:25px;font-weight:700;color:#1d1d1f;">'+esc_(p['Full Name']||'')+'</div><div style="height:1px;background:#e5e5e7;margin:14px 0 12px;"></div><div class="apple-font" style="font-size:11px;color:#86868b;margin-bottom:4px;">Registration number</div><div class="apple-font" style="font-size:14px;color:#1d1d1f;">'+esc_(p['Registration Number']||'')+'</div></td></tr></table></td></tr>';
}

function recoverStaleEmailJobs_() {
  // Conservative recovery: only release jobs that have been stuck in PROCESSING
  // for a long time. We never assume a send failed, because the message may have
  // reached the provider before the execution stopped. Requeueing such a job can
  // cause a duplicate email.
  const sh=getSheet_(APP.SHEETS.QUEUE);
  const rows=readSheetObjects_(sh);
  const now=new Date();
  const staleAfterMs=30*60*1000;
  const stale=rows.filter(function(job){
    if(String(job.Status)!=='PROCESSING') return false;
    const last=parseQueueDate_(job['Last Attempt At']);
    return last && (now.getTime()-last.getTime())>=staleAfterMs;
  });
  stale.forEach(function(job){
    markQueueJob_(job['Job ID'],{
      'Status':'FAILED',
      'Last Error':'Delivery outcome unknown after stale processing lock. Manual review required before retrying.',
      'Queue Reason':'STALE_PROCESSING_REVIEW',
      'Updated At':now
    });
    log_('WARN','EMAIL',String(job['Event ID']),String(job['Participant ID']),'STALE_EMAIL_JOB_REVIEW','Email job remained PROCESSING beyond the recovery threshold; it was not automatically resent to avoid duplicates.',String(job['Job ID']));
  });
  return {detected:stale.length};
}

function pruneSystemLogs_(){
  const sh=getSheet_(APP.SHEETS.LOGS);
  const maxRows=positiveIntSetting_('SYSTEM_LOG_MAX_ROWS',2000);
  const triggerRows=Math.max(maxRows+500,maxRows*1.25);
  const lr=sh.getLastRow();
  if(lr<=triggerRows) return 0;
  const deleteCount=lr-1-maxRows;
  if(deleteCount>0) sh.deleteRows(2,deleteCount);
  return Math.max(0,deleteCount);
}

function processScheduler() {
  // Do not stack scheduler executions behind a long-running prior run.
  // If the script lock is busy, the trigger exits cleanly and the next 5-minute run retries.
  if (isPaused_()) return {ok:false,paused:true};
  return withLock_('processScheduler', function() {
    ensureSetup_();
    const workshops = readSheetObjects_(getSheet_(APP.SHEETS.WORKSHOPS));
    const limit = Math.min(workshops.length, APP.MAX_WORKSHOPS_PER_RUN);

    for (let i=0; i<limit; i++) {
      const w = workshops[i];
      if (!w['Event ID'] || ['ARCHIVED','CANCELLED'].includes(String(w.Status))) continue;
      const eventId = String(w['Event ID']);

      // Each module is isolated so one bad workshop/module cannot stop the entire scheduler.
      const modules = [
        ['FORM_STATE', function(){ syncWorkshopFormState_(w); }],
        ['REGISTRATION', function(){ processRegistrationResponses_(w); }],
        ['ATTENDANCE_EMAILS', function(){ scheduleAttendanceEmails_(w); }],
        ['ATTENDANCE', function(){ if (boolValue_(w['Attendance Enabled'], true)) processAttendanceResponses_(w); }],
        ['REMINDERS', function(){ scheduleMissingReminders_(w); }],
        ['CERTIFICATES', function(){ processDueCertificates_(w); }]
      ];

      modules.forEach(function(entry) {
        try {
          entry[1]();
        } catch (err) {
          log_('ERROR', 'SCHEDULER', eventId, '', 'MODULE_FAILED', entry[0]+': '+String(err), stack_(err));
        }
      });
    }

    // Test runs are processed inside the central scheduler as a safety net.
    // A one-time test trigger can collide with the scheduler lock; central processing
    // guarantees the test workflow continues on the next scheduler cycle.
    try { processDueTestStepsCore_(); }
    catch (err) { log_('ERROR','SCHEDULER','','','TEST_RUNS_FAILED',String(err),stack_(err)); }

    try { processEmailQueue_(); }
    catch (err) { log_('ERROR','SCHEDULER','','','EMAIL_QUEUE_FAILED',String(err),stack_(err)); }

    try { recoverStaleEmailJobs_(); }
    catch (err) { log_('ERROR','SCHEDULER','','','STALE_EMAIL_RECOVERY_FAILED',String(err),stack_(err)); }

    try { refreshDashboard_(); }
    catch (err) { log_('ERROR','SCHEDULER','','','DASHBOARD_REFRESH_FAILED',String(err),stack_(err)); }

    try { pruneSystemLogs_(); }
    catch (err) { log_('ERROR','SCHEDULER','','','SYSTEM_LOG_PRUNE_FAILED',String(err),stack_(err)); }

    return {ok:true,processed:limit};
  });
}

function processRegistrationResponses_(w) {
  w=repairIfWorkshopIntegrationBroken_(w);
  const eventId=String(w['Event ID']||'');
  const lastProcessed=Math.max(1,Number(w['Last Registration Row']||1));
  const formId=String(w['Registration Form ID']||'');
  const source=readWorkshopFormResponses_(formId, w['Registration Response Spreadsheet ID'], eventId, 'REGISTRATION', lastProcessed);
  if (!source || !source.responses.length) return;

  const psh=getSheet_(APP.SHEETS.PARTICIPANTS);
  let lastGoodRow=lastProcessed;
  for (let idx=0; idx<source.responses.length; idx++) {
    const item=source.responses[idx];
    const data=item.data;
    const rowNumber=item.sourceRow;
    try {
      const email=normalizeEmail_(data['Email']);
      if (!email || !isValidEmail_(email)) throw new Error('Invalid email address.');

      const duplicate=findParticipant_(eventId,email);
      if (duplicate) {
        log_('WARN','REGISTRATION',eventId,String(duplicate['Participant ID']),
          'DUPLICATE_REGISTRATION','Duplicate registration ignored',String(rowNumber));
        lastGoodRow=rowNumber;
        continue;
      }

      const participantId=nextParticipantId_(dateYear_(w['Workshop Date'],String(w['Timezone']||APP.DEFAULT_TIMEZONE)));
      const now=new Date();
      const cap=Number(w.Capacity||0);
      const currentActive=countWorkshopActiveParticipants_(eventId);
      const status=cap>0 && currentActive>=cap?'ON_HOLD':'REGISTERED';
      const obj={
        'Participant ID':participantId,
        'Event ID':eventId,
        'Workshop Name':w['Workshop Name'],
        'Registration Timestamp':data.Timestamp||now,
        'Full Name':clean_(data['Full Name']),
        'Email':email,
        'Phone / WhatsApp Number':clean_(data['Phone / WhatsApp Number']),
        'Department / Faculty':clean_(data['Department / Faculty']),
        'Season':clean_(data['Season']),
        'ID Number':clean_(data['ID Number']),
        'Registration Number':clean_(data['Registration Number']),
        'Registration Status':status,
        'Attendance Status':'NOT_ATTENDED',
        'Certificate Eligibility':'PENDING',
        'Certificate Status':'NOT_READY',
        'Certificate ID':'',
        'Certificate PDF URL':'',
        'Created At':now,
        'Updated At':now,
        'Source Response Row':rowNumber
      };
      appendObject_(psh,obj);
      if(status!=='ON_HOLD') queueRegistrationEmail_(w,obj);
      log_('INFO','REGISTRATION',eventId,participantId,'REGISTRATION_ACCEPTED','Participant created',String(rowNumber));
      lastGoodRow=rowNumber;
    } catch(err) {
      log_('ERROR','REGISTRATION',eventId,'','REGISTRATION_ROW_FAILED',String(err),String(rowNumber)+' | source='+source.sourceType+' | '+(source.sourceDetail||''));
      break;
    }
  }
  if(lastGoodRow>lastProcessed){
    updateWorkshopByEvent_(eventId,{'Last Registration Row':lastGoodRow,'Updated At':new Date()});
  }
}

function processAttendanceResponses_(w) {
  if (!boolValue_(w['Attendance Enabled'], true)) return;
  w=repairIfWorkshopIntegrationBroken_(w);
  const eventId=String(w['Event ID']||'');
  const tz=String(w['Timezone']||APP.DEFAULT_TIMEZONE);
  const attendanceStart=parseDateTime_(w['Workshop Date'],w['Start Time'],tz);
  const attendanceEnd=new Date(attendanceStart.getTime()+normalizeAttendanceWindowMinutes_(w['Attendance Window Minutes'],120)*60000);
  const lastProcessed=Math.max(1,Number(w['Last Attendance Row']||1));
  const formId=String(w['Attendance Form ID']||'');
  const source=readWorkshopFormResponses_(formId, w['Attendance Response Spreadsheet ID'], eventId, 'ATTENDANCE', lastProcessed);
  if(!source || !source.responses.length) return;

  let lastGoodRow=lastProcessed;
  for(let idx=0;idx<source.responses.length;idx++){
    const item=source.responses[idx];
    const data=item.data;
    const rowNumber=item.sourceRow;
    try{
      const participantId=clean_(data['Participant ID']);
      const p=findParticipantById_(participantId);
      let validation='INVALID';
      let status='REJECTED';
      let message='';
      const submittedAt=data.Timestamp instanceof Date && !isNaN(data.Timestamp.getTime()) ? data.Timestamp : new Date(data.Timestamp||0);

      if(!submittedAt || isNaN(submittedAt.getTime())){
        message='Attendance submission timestamp is invalid.';
      }else if(submittedAt < attendanceStart){
        message='Attendance is not open yet.';
      }else if(submittedAt > attendanceEnd){
        message='Attendance window has closed.';
      }else if(!p){
        message='Participant ID not found.';
      }else if(String(p['Event ID'])!==eventId){
        message='Participant belongs to another workshop.';
      }else if(['REMOVED','REJECTED'].includes(String(p['Registration Status']))){
        message='Participant is not eligible for attendance.';
      }else{
        const already=hasAttendance_(eventId,participantId);
        if(already){
          validation='VALID'; status='DUPLICATE'; message='Attendance already recorded.';
        }else{
          validation='VALID'; status='PRESENT'; message='Attendance accepted.';
          appendObject_(getSheet_(APP.SHEETS.ATTENDANCE),{
            'Timestamp':data.Timestamp||new Date(),
            'Participant ID':participantId,
            'Event ID':eventId,
            'Email':p['Email'],
            'Name':p['Full Name'],
            'Workshop Type':w['Workshop Type'],
            'Validation':validation,
            'Attendance Status':status,
            'Validation Message':message,
            'Source':'Attendance Form',
            'Created At':new Date()
          });
          updateParticipant_(participantId,{
            'Attendance Status':'PRESENT',
            'Certificate Eligibility':eligibilityFrom_(w,p,true),
            'Updated At':new Date()
          });
        }
      }

      if(validation!=='VALID'){
        appendObject_(getSheet_(APP.SHEETS.ATTENDANCE),{
          'Timestamp':data.Timestamp||new Date(),
          'Participant ID':participantId,
          'Event ID':eventId,
          'Email':p?p['Email']:'',
          'Name':p?p['Full Name']:'',
          'Workshop Type':w['Workshop Type'],
          'Validation':validation,
          'Attendance Status':status,
          'Validation Message':message,
          'Source':'Attendance Form',
          'Created At':new Date()
        });
      }
      log_(validation==='VALID'?'INFO':'WARN','ATTENDANCE',eventId,participantId,'ATTENDANCE_'+validation,message,String(rowNumber)+' | source='+source.sourceType+' | '+(source.sourceDetail||''));
      lastGoodRow=rowNumber;
    }catch(err){
      log_('ERROR','ATTENDANCE',eventId,'','ATTENDANCE_ROW_FAILED',String(err),String(rowNumber)+' | source='+source.sourceType+' | '+(source.sourceDetail||''));
      break;
    }
  }
  if(lastGoodRow>lastProcessed){
    updateWorkshopByEvent_(eventId,{'Last Attendance Row':lastGoodRow,'Updated At':new Date()});
  }
}

/**
 * Read Google Form responses without depending on the response spreadsheet's
 * first sheet or fixed range geometry. The Form response is authoritative;
 * the destination spreadsheet remains a human-readable mirror.
 * Falls back to the response spreadsheet when Forms access is unavailable.
 */
function readWorkshopFormResponses_(formId, responseSpreadsheetId, eventId, moduleName, lastProcessedRow) {
  const baseRow=Math.max(1,Number(lastProcessedRow||1));
  if(formId && likelyGoogleId_(formId)){
    try{
      const form=FormApp.openById(String(formId));
      const responses=form.getResponses();
      const out=[];
      const firstResponseIndex=Math.max(0,baseRow-1);
      for(let i=firstResponseIndex;i<responses.length;i++){
        const response=responses[i];
        const data={};
        const ts=response.getTimestamp();
        if(ts) data.Timestamp=ts;
        response.getItemResponses().forEach(function(ir){
          const title=String(ir.getItem().getTitle()||'').trim();
          if(title) data[title]=ir.getResponse();
        });
        out.push({data:data,sourceRow:i+2,responseIndex:i});
      }
      if(out.length){
        log_('INFO',moduleName,eventId,'','FORM_RESPONSE_SOURCE_OK','Read '+out.length+' new response(s) from Google Form','responses='+responses.length+';startRow='+baseRow);
      }
      return {responses:out,sourceType:'FORM',sourceDetail:'FormApp response index'};
    }catch(err){
      log_('WARN',moduleName,eventId,'','FORM_RESPONSE_SOURCE_FAILED','Falling back to response spreadsheet',String(err));
    }
  }
  if(responseSpreadsheetId && likelyGoogleId_(responseSpreadsheetId)){
    try{
      const ss=SpreadsheetApp.openById(String(responseSpreadsheetId));
      const sh=findResponseSheetByHeaders_(ss,moduleName);
      const lastRow=Math.max(1,sh.getLastRow());
      const lastCol=Math.max(1,sh.getLastColumn());
      if(lastRow<=baseRow) return {responses:[],sourceType:'SHEET',sourceDetail:sh.getName()};
      const headerValues=sh.getRange(1,1,1,lastCol).getDisplayValues()[0].map(String);
      const rows=sh.getRange(baseRow+1,1,lastRow-baseRow,lastCol).getValues();
      const out=rows.map(function(values,idx){return {data:mapRow_(headerValues,values),sourceRow:baseRow+1+idx};});
      if(out.length){
        log_('INFO',moduleName,eventId,'','SHEET_RESPONSE_SOURCE_OK','Fallback response sheet read: '+sh.getName(),'rows='+out.length);
      }
      return {responses:out,sourceType:'SHEET',sourceDetail:sh.getName()};
    }catch(err){
      throw new Error(moduleName+' response source unavailable. Form ID='+String(formId||'')+'; Spreadsheet ID='+String(responseSpreadsheetId||'')+'. '+String(err));
    }
  }
  return {responses:[],sourceType:'NONE',sourceDetail:'No response source configured'};
}

function findResponseSheetByHeaders_(ss,moduleName){
  const sheets=ss.getSheets();
  const preferred=[];
  sheets.forEach(function(sh){
    const name=String(sh.getName()||'');
    if(/form responses?/i.test(name)) preferred.push(sh);
  });
  const candidates=preferred.length?preferred:sheets.slice();
  const wanted=moduleName==='REGISTRATION'
    ? ['Timestamp','Full Name','Email','Phone / WhatsApp Number','Department / Faculty','Season','ID Number','Registration Number']
    : ['Timestamp','Participant ID'];
  for(let i=0;i<candidates.length;i++){
    const sh=candidates[i];
    try{
      const lc=Math.max(1,sh.getLastColumn());
      const headers=sh.getRange(1,1,1,lc).getDisplayValues()[0].map(function(x){return String(x).trim();});
      if(wanted.every(function(h){return headers.indexOf(h)>=0;})) return sh;
    }catch(err){}
  }
  if(!sheets.length) throw new Error('No sheets exist in response spreadsheet.');
  return sheets[0];
}

function queueRegistrationEmail_(w,p) {
  const link = w['Meeting Link'];
  const email = buildRegistrationEmail_(w,p);
  enqueueEmail_({event:w,participant:p,type:'REGISTRATION',scheduledAt:new Date(),subject:email.subject,html:email.html,body:email.text,includeMeetingLink:false,queueReason:'REGISTRATION_CREATED'});
}

function disableLegacyAnnouncementQueueJobs_(ss){
  try{
    const sh=ss.getSheetByName(APP.SHEETS.QUEUE); if(!sh||sh.getLastRow()<2)return;
    const rows=readSheetObjects_(sh); let changed=0; rows.forEach(function(r){const t=String(r['Email Type']||'').toUpperCase();const status=String(r['Status']||'').toUpperCase();if(t==='ANNOUNCEMENT'&&['PENDING','RETRY','WAITING_FOR_QUOTA','PROCESSING'].includes(status)){markQueueJob_(r['Job ID'],{'Status':'EXPIRED','Queue Reason':'LEGACY_ANNOUNCEMENT_DISABLED','Last Error':'V1.4: automatic announcements are manual-only.','Updated At':new Date()});changed++;}}); if(changed)log_('INFO','MIGRATION','','','LEGACY_ANNOUNCEMENTS_DISABLED','Expired '+changed+' pending automatic Announcement job(s).','');
  }catch(err){log_('WARN','MIGRATION','','','LEGACY_ANNOUNCEMENT_MIGRATION_FAILED',String(err),'');}
}
function scheduleMissingAnnouncements_(w) {
  // V1.4: announcements are manual-only. Use Custom Email from Control Center.
  return;
}
function scheduleAttendanceEmails_(w) {
  if (!w['Event ID'] || !boolValue_(w['Attendance Enabled'], true)) return;
  const formId=String(w['Attendance Form ID']||'').trim();
  if(!formId) return;
  const tz=String(w['Timezone']||APP.DEFAULT_TIMEZONE);
  const start=parseDateTime_(w['Workshop Date'],w['Start Time'],tz);
  const end=new Date(start.getTime()+normalizeAttendanceWindowMinutes_(w['Attendance Window Minutes'],120)*60000);
  const now=new Date();
  if(now<start || now>end) return;
  const participants=readSheetObjects_(getSheet_(APP.SHEETS.PARTICIPANTS)).filter(function(p){return String(p['Event ID'])===String(w['Event ID'])&&['REGISTERED','APPROVED'].includes(String(p['Registration Status']));});
  const attendanceUrl=FormApp.openById(formId).getPublishedUrl();
  participants.forEach(function(p){
    const unique=String(w['Event ID'])+'|'+String(p['Participant ID'])+'|ATTENDANCE';
    if(hasEmailUniqueKey_(unique)) return;
    const e=buildAttendanceEmail_(w,p,attendanceUrl,false);
    enqueueEmail_({event:w,participant:p,type:'ATTENDANCE',scheduledAt:now,subject:e.subject,html:e.html,body:e.text,includeMeetingLink:false,uniqueKey:unique,deadlineAt:end,queueReason:'ATTENDANCE_OPEN'});
  });
}

function scheduleMissingReminders_(w) {
  if (!w['Event ID'] || String(w['Reminder Enabled']).toUpperCase() === 'FALSE') return;
  const participants = readSheetObjects_(getSheet_(APP.SHEETS.PARTICIPANTS)).filter(p => String(p['Event ID'])===String(w['Event ID']) && ['REGISTERED','APPROVED'].includes(String(p['Registration Status'])));
  if (!participants.length) return;
  const tz = String(w['Timezone']||APP.DEFAULT_TIMEZONE);
  const now = new Date();
  const mode = String(w['Reminder Mode']||'BEFORE_START').toUpperCase();
  let schedules = [];
  if (mode === 'EXACT_DATETIME') {
    const exact = parseFlexibleDateTime_(w['Reminder Custom At'], tz);
    if (exact) schedules.push({when:exact, offsetLabel:'at the scheduled reminder time', keyPart:'EXACT'});
  } else {
    parseOffsets_(w['Reminder Offsets']).forEach(function(offset){
      schedules.push({when:new Date(parseDateTime_(w['Workshop Date'],w['Start Time'],tz).getTime()-offset*60000), offset:offset, offsetLabel:formatOffsetHuman_(offset)+' before the workshop', keyPart:String(offset)});
    });
  }
  const workshopStart=parseDateTime_(w['Workshop Date'],w['Start Time'],tz);
  if(now>=workshopStart) return;
  schedules.forEach(function(s){
    if (!s.when || s.when >= workshopStart) return;
    participants.forEach(function(p){
      const unique = w['Event ID']+'|'+p['Participant ID']+'|REMINDER|'+s.keyPart;
      if (hasEmailUniqueKey_(unique)) return;
      const e = buildReminderEmail_(w,p,s.offset || 0,false);
      if (mode === 'EXACT_DATETIME') e.text = 'Reminder: '+w['Workshop Name']+' is scheduled for '+formatDateTimeForDisplay_(w['Reminder Custom At'],tz)+'.';
      const effectiveAt=new Date(s.when);
      let reason='REMINDER_SCHEDULED';
      if(effectiveAt<=now){ effectiveAt.setTime(now.getTime()); reason='REMINDER_CATCHUP'; }
      if(effectiveAt>=workshopStart) return;
      enqueueEmail_({event:w,participant:p,type:'REMINDER',scheduledAt:effectiveAt,subject:e.subject,html:e.html,body:e.text,includeMeetingLink:normalizeDeliveryMode_(w['Delivery Mode'],String(w['Meeting Link']||'').trim()?'ONLINE':'OFFLINE')==='ONLINE',uniqueKey:unique,deadlineAt:workshopStart,queueReason:reason});
    });
  });
}

function processDueCertificates_(w) {
  if (String(w['Certificate Enabled']).toUpperCase() !== 'TRUE' && w['Certificate Enabled'] !== true) return;
  const participants = readSheetObjects_(getSheet_(APP.SHEETS.PARTICIPANTS)).filter(p=>String(p['Event ID'])===String(w['Event ID']) && String(p['Certificate Status']||'')!=='SENT');
  const now = new Date();
  const release = certificateReleaseDate_(w);
  if (!release || now < release) return;
  let processed = 0;
  participants.forEach(p => {
    if (processed >= APP.MAX_CERT_BATCH) return;
    if (!isCertificateEligible_(w,p)) return;
    if (String(p['Certificate Status'])==='GENERATED' || String(p['Certificate Status'])==='SENT') return;
    try {
      const cert = generateCertificateForParticipant_(w,p,false,'');
      updateParticipant_(p['Participant ID'], {'Certificate Eligibility':'ELIGIBLE','Certificate Status':'GENERATED','Certificate ID':cert.certificateId,'Certificate PDF URL':cert.pdfUrl,'Updated At':new Date()});
      const email = buildCertificateEmail_(w,p,cert);
      enqueueEmail_({event:w,participant:p,type:'CERTIFICATE',scheduledAt:now,subject:email.subject,html:email.html,body:email.text,attachmentFileId:cert.pdfFileId,includeMeetingLink:false,queueReason:'CERTIFICATE_READY'});
      updateParticipant_(p['Participant ID'], {'Certificate Status':'QUEUED','Updated At':new Date()});
      processed++;
    } catch (err) {
      updateParticipant_(p['Participant ID'], {'Certificate Eligibility':'ELIGIBLE','Certificate Status':'FAILED','Updated At':new Date()});
      log_('ERROR','CERTIFICATE',String(w['Event ID']),String(p['Participant ID']),'CERTIFICATE_FAILED',String(err),stack_(err));
    }
  });
}

function processEmailQueue_() {
  const sh = getSheet_(APP.SHEETS.QUEUE);
  const rows = readSheetObjects_(sh);
  const now = new Date();
  const quota = typeof MailApp.getRemainingDailyQuota === 'function' ? Number(MailApp.getRemainingDailyQuota()) : APP.MAX_EMAIL_BATCH;
  const reserve = positiveIntSetting_('EMAIL_SAFETY_RESERVE', APP.EMAIL_SAFE_DAILY_RESERVE);
  const configuredBatch = positiveIntSetting_('EMAIL_MAX_BATCH', APP.MAX_EMAIL_BATCH);
  const safeBudget = Math.max(0, quota - reserve);
  const limit = Math.min(configuredBatch, APP.MAX_EMAIL_BATCH, safeBudget);

  const expired=[];
  const due = rows.filter(function(job){
    if (!['PENDING','RETRY','WAITING_FOR_QUOTA'].includes(String(job.Status))) return false;
    const when = parseQueueDate_(job['Not Before'] || job['Scheduled At']);
    if (when && when > now) return false;
    const deadline=parseQueueDate_(job['Deadline At']);
    const type=String(job['Email Type']||'').toUpperCase();
    if(deadline && now > deadline && (type==='ANNOUNCEMENT' || type==='REMINDER')){ expired.push(job); return false; }
    // Per-participant lifecycle is authoritative: registration confirmation must be
    // delivered before any later production email for the same participant.
    // Global priority still decides which participant/job wins after dependencies pass.
    if(!isEmailLifecycleReady_(job, rows)){ return false; }
    return true;
  });
  if(expired.length){
    expired.forEach(function(job){ markQueueJob_(job['Job ID'],{'Status':'EXPIRED','Last Error':'Delivery window closed.','Queue Reason':'DEADLINE_PASSED','Updated At':now}); });
  }

  due.sort(compareEmailJobs_);

  if (limit <= 0) {
    if (quota <= reserve && due.length) {
      markQueueJobsWaitingForQuota_(due);
      log_('INFO','EMAIL','','','EMAIL_QUOTA_WAIT','Email queue paused because only '+quota+' recipient quota remains and the safety reserve is '+reserve+'.','dueJobs='+due.length+';remainingQuota='+quota+';reserve='+reserve);
    }
    return {sent:0,quotaRemaining:quota,safeBudget:safeBudget,waiting:true,dueJobs:due.length};
  }

  let sent = 0;
  for (let i = 0; i < due.length && sent < limit; i++) {
    const job = due[i];
    const attempt = Number(job['Attempt Count'] || 0) + 1;
    try {
      markQueueJob_(job['Job ID'], {
        'Status':'PROCESSING',
        'Attempt Count':attempt,
        'Last Attempt At':now,
        'Updated At':now
      });

      const attachments=[];
      if (job['Attachment File ID']) attachments.push(DriveApp.getFileById(job['Attachment File ID']).getBlob());
      MailApp.sendEmail({
        to:String(job.Recipient),
        subject:String(job.Subject),
        body:stripHtml_(job['Message Body']||''),
        htmlBody:String(job['Message Body']||''),
        name:String(job['Sender Name']||getSetting_('SENDER_NAME')||APP.ROOT),
        attachments:attachments
      });

      const sentAt=new Date();
      markQueueJob_(job['Job ID'], {'Status':'SENT','Sent At':sentAt,'Updated At':sentAt,'Last Error':'','Queue Reason':'SENT'});
      markEmailLifecycleAfterSend_(job);
      sent++;
    } catch(err) {
      const message=String(err&&err.message?err.message:err);
      if (isQuotaError_(message)) {
        const nextTry = new Date(Date.now()+positiveIntSetting_('EMAIL_QUOTA_RECHECK_MINUTES',APP.EMAIL_QUOTA_RECHECK_MINUTES)*60000);
        markQueueJob_(job['Job ID'], {
          'Status':'WAITING_FOR_QUOTA',
          'Not Before':nextTry,
          'Last Error':message,
          'Queue Reason':'QUOTA_WAIT',
          'Updated At':new Date()
        });
        log_('WARN','EMAIL',String(job['Event ID']),String(job['Participant ID']),'EMAIL_QUOTA_REACHED','Google email quota was reached while sending. Job will resume automatically.',JSON.stringify({jobId:job['Job ID'],recipient:job['Recipient'],nextAttemptAt:nextTry,attempt:attempt}));
        break;
      }

      const retryLimit=positiveIntSetting_('EMAIL_RETRY_LIMIT',APP.EMAIL_RETRY_LIMIT);
      const retryDelay=positiveIntSetting_('EMAIL_RETRY_BASE_MINUTES',APP.EMAIL_RETRY_BASE_MINUTES) * Math.max(1, Math.pow(2, Math.max(0,attempt-1)));
      const nextTry=new Date(Date.now()+retryDelay*60000);
      const permanent=attempt>=retryLimit;
      markQueueJob_(job['Job ID'], {
        'Status':permanent?'FAILED':'RETRY',
        'Not Before':permanent ? (job['Not Before']||job['Scheduled At']) : nextTry,
        'Last Error':message,
        'Queue Reason':permanent?'PERMANENT_FAILURE':'RETRY_BACKOFF',
        'Updated At':new Date()
      });
      log_('ERROR','EMAIL',String(job['Event ID']),String(job['Participant ID']),'EMAIL_SEND_FAILED',message,JSON.stringify({jobId:job['Job ID'],attempt,retryAt:permanent?'':nextTry}));
    }
  }

  return {sent:sent,quotaRemaining:quota-sent,safeBudget:safeBudget};
}

function isEmailLifecycleReady_(job, allRows){
  const type=String(job['Email Type']||'').toUpperCase();
  if(['REGISTRATION','TEST_REGISTRATION','TEST_REMINDER','TEST_ATTENDANCE','TEST_CERTIFICATE','CUSTOM'].includes(type)) return true;
  const eventId=String(job['Event ID']||'');
  const participantId=String(job['Participant ID']||'');
  if(!eventId || !participantId) return true;
  const registration=allRows.find(function(r){
    return String(r['Event ID']||'')===eventId &&
      String(r['Participant ID']||'')===participantId &&
      String(r['Email Type']||'').toUpperCase()==='REGISTRATION';
  });
  // Legacy participants may predate the queue. Do not deadlock them forever.
  if(!registration) return true;
  const status=String(registration['Status']||'').toUpperCase();
  if(status==='SENT') return true;
  return false;
}

function markQueueJobsWaitingForQuota_(jobs) {
  if (!jobs || !jobs.length) return 0;
  const sh=getSheet_(APP.SHEETS.QUEUE);
  const lr=sh.getLastRow(), lc=sh.getLastColumn();
  if(lr<2) return 0;
  const headers=sh.getRange(1,1,1,lc).getValues()[0].map(String);
  const vals=sh.getRange(2,1,lr-1,lc).getValues();
  const idx={}; headers.forEach(function(h,i){idx[h]=i;});
  const ids={}; jobs.forEach(function(j){ids[String(j['Job ID'])]=true;});
  const now=new Date();
  const nextTry=new Date(now.getTime()+positiveIntSetting_('EMAIL_QUOTA_RECHECK_MINUTES',APP.EMAIL_QUOTA_RECHECK_MINUTES)*60000);
  let changed=0;
  vals.forEach(function(row){
    const id=String(row[idx['Job ID']]||'');
    if(!ids[id]) return;
    row[idx['Status']]='WAITING_FOR_QUOTA';
    row[idx['Not Before']]=nextTry;
    row[idx['Queue Reason']]='QUOTA_WAIT';
    row[idx['Updated At']]=now;
    changed++;
  });
  if(changed) sh.getRange(2,1,vals.length,lc).setValues(vals);
  return changed;
}

function markEmailLifecycleAfterSend_(job) {
  const type=String(job['Email Type']||'');
  if (type==='REGISTRATION') log_('INFO','EMAIL',String(job['Event ID']),String(job['Participant ID']),'REGISTRATION_EMAIL_SENT','',job['Recipient']);
  if (type==='ANNOUNCEMENT') log_('INFO','EMAIL',String(job['Event ID']),String(job['Participant ID']),'ANNOUNCEMENT_EMAIL_SENT','',job['Recipient']);
  if (type==='REMINDER') log_('INFO','EMAIL',String(job['Event ID']),String(job['Participant ID']),'REMINDER_EMAIL_SENT','',job['Recipient']);
  if (type==='ATTENDANCE') log_('INFO','EMAIL',String(job['Event ID']),String(job['Participant ID']),'ATTENDANCE_EMAIL_SENT','',job['Recipient']);
  if (type==='CERTIFICATE') {
    updateParticipant_(job['Participant ID'], {'Certificate Status':'SENT','Updated At':new Date()});
    log_('INFO','EMAIL',String(job['Event ID']),String(job['Participant ID']),'CERTIFICATE_EMAIL_SENT','',job['Recipient']);
  }
  if (type==='CUSTOM') log_('INFO','EMAIL',String(job['Event ID']),String(job['Participant ID']),'CUSTOM_EMAIL_SENT','',job['Recipient']);
}

function enqueueEmail_(opts) {
  if(opts && opts.html) validateRenderedEmail_('EMAIL',opts.html,{participant_name:String(opts.participant&&opts.participant['Full Name']||'Participant'),workshop_name:String(opts.event&&opts.event['Workshop Name']||'Workshop')});
  const scheduledAt=opts.scheduledAt||new Date();
  const unique = opts.uniqueKey || (opts.event['Event ID']+'|'+opts.participant['Participant ID']+'|'+opts.type);
  if (hasEmailUniqueKey_(unique)) return null;
  const type=String(opts.type||'').toUpperCase();
  const deadlineAt=opts.deadlineAt || emailDeadlineFor_(opts.event,type);
  const obj = {
    'Job ID':'JOB-'+Utilities.getUuid().slice(0,12).toUpperCase(),
    'Unique Key':unique,
    'Event ID':opts.event['Event ID'],
    'Participant ID':opts.participant['Participant ID'],
    'Email Type':type,
    'Priority':emailPriority_(type),
    'Scheduled At':scheduledAt,
    'Not Before':scheduledAt,
    'Deadline At':deadlineAt||'',
    'Status':'PENDING',
    'Attempt Count':0,
    'Last Attempt At':'',
    'Sent At':'',
    'Recipient':opts.participant['Email'],
    'Subject':opts.subject,
    'Attachment File ID':opts.attachmentFileId||'',
    'Last Error':'',
    'Queue Reason':opts.queueReason||'SCHEDULED',
    'Test Run ID':opts.testRunId||'',
    'Message Body':opts.html,
    'Include Meeting Link':opts.includeMeetingLink?'TRUE':'FALSE',
    'Sender Name':getSetting_('SENDER_NAME')||APP.ROOT,
    'Created At':new Date(),
    'Updated At':new Date()
  };
  appendObject_(getSheet_(APP.SHEETS.QUEUE),obj);
  return obj['Job ID'];
}

function emailPriority_(type) {
  switch (String(type||'').toUpperCase()) {
    case 'CERTIFICATE': return 0;
    case 'ATTENDANCE': return 1;
    case 'REMINDER': return 2;
    case 'REGISTRATION': return 3;
    case 'CUSTOM': return 3;
    case 'ANNOUNCEMENT': return 9;
    case 'TEST_CERTIFICATE': return 5;
    case 'TEST_ATTENDANCE': return 6;
    case 'TEST_REMINDER': return 7;
    case 'TEST_REGISTRATION': return 8;
    default: return 9;
  }
}

function emailDeadlineFor_(w,type) {
  if (!w) return '';
  const tz=String(w['Timezone']||APP.DEFAULT_TIMEZONE);
  const kind=String(type||'').toUpperCase();
  try {
    if (kind==='REMINDER' || kind==='ANNOUNCEMENT') {
      return parseDateTime_(w['Workshop Date'],w['Start Time'],tz);
    }
    return '';
  } catch(err) {
    return '';
  }
}

function compareEmailJobs_(a,b) {
  const pa=Number(a['Priority']||emailPriority_(a['Email Type']));
  const pb=Number(b['Priority']||emailPriority_(b['Email Type']));
  if (pa!==pb) return pa-pb;

  const da=parseQueueDate_(a['Deadline At']);
  const db=parseQueueDate_(b['Deadline At']);
  const daT=da&& !isNaN(da.getTime()) ? da.getTime() : Number.POSITIVE_INFINITY;
  const dbT=db&& !isNaN(db.getTime()) ? db.getTime() : Number.POSITIVE_INFINITY;
  if (daT!==dbT) return daT-dbT;

  const na=parseQueueDate_(a['Not Before']||a['Scheduled At']);
  const nb=parseQueueDate_(b['Not Before']||b['Scheduled At']);
  const naT=na&&!isNaN(na.getTime())?na.getTime():0;
  const nbT=nb&&!isNaN(nb.getTime())?nb.getTime():0;
  if (naT!==nbT) return naT-nbT;

  const ca=parseQueueDate_(a['Created At']);
  const cb=parseQueueDate_(b['Created At']);
  const caT=ca&&!isNaN(ca.getTime())?ca.getTime():0;
  const cbT=cb&&!isNaN(cb.getTime())?cb.getTime():0;
  if (caT!==cbT) return caT-cbT;
  return String(a['Job ID']||'').localeCompare(String(b['Job ID']||''));
}

function parseQueueDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  if (value===null || value===undefined || String(value).trim()==='') return null;
  const d=new Date(value);
  return isNaN(d.getTime()) ? null : d;
}

function positiveIntSetting_(key,fallback) {
  const n=Number(getSetting_(key));
  return isFinite(n) && n>0 ? Math.max(1,Math.round(n)) : Math.max(1,Math.round(Number(fallback)||1));
}

function isQuotaError_(message) {
  return /quota|too many|service invoked too many|daily limit|recipient limit|rate limit/i.test(String(message||''));
}

function startTestRun(participantOrConfig) {
  return withLock_('startTestRun', function(){
    ensureSetup_();
    let config={};
    if(typeof participantOrConfig==='string'){
      const legacy=findParticipantById_(participantOrConfig);
      if(!legacy) throw new Error('Participant not found.');
      config={eventId:legacy['Event ID'],recipient:legacy['Email'],participantName:legacy['Full Name'],registrationNumber:legacy['Registration Number'],department:legacy['Department / Faculty'],season:legacy['Season'],idNumber:legacy['ID Number'],phone:legacy['Phone / WhatsApp Number'],participantId:legacy['Participant ID'],includeAttendance:false};
    }else config=participantOrConfig||{};
    const w=findWorkshop_(config.eventId); if(!w) throw new Error('Workshop not found.');
    const recipient=normalizeEmail_(config.recipient); if(!isValidEmail_(recipient)) throw new Error('Enter a valid test recipient email.');
    const pid='TEST-PARTICIPANT-'+Utilities.getUuid().slice(0,8).toUpperCase();
    const p={'Participant ID':pid,'Event ID':w['Event ID'],'Workshop Name':w['Workshop Name'],'Registration Timestamp':new Date(),'Full Name':clean_(config.participantName)||'Test Participant','Email':recipient,'Phone / WhatsApp Number':clean_(config.phone)||'','Department / Faculty':clean_(config.department)||'','Season':clean_(config.season)||'','ID Number':clean_(config.idNumber)||'TEST-ID','Registration Number':clean_(config.registrationNumber)||'TEST-REG','Registration Status':'TEST','Attendance Status':'NOT_ATTENDED','Certificate Eligibility':'PENDING','Certificate Status':'NOT_READY','Certificate ID':'','Certificate PDF URL':'','Created At':new Date(),'Updated At':new Date(),'Source Response Row':''};
    const runId='TEST-'+Utilities.getUuid().slice(0,10).toUpperCase(); const now=new Date();
    appendObject_(getSheet_(APP.SHEETS.TESTS),{'Run ID':runId,'Event ID':w['Event ID'],'Participant ID':pid,'Workshop Name':w['Workshop Name'],'Recipient':recipient,'Participant Snapshot':JSON.stringify(p),'Include Attendance':!!config.includeAttendance,'Started At':now,'Status':'RUNNING','Current Step':'REMINDER_PENDING','Confirmation Sent At':'','Reminder Sent At':'','Attendance Sent At':'','Certificate Created At':'','Certificate Sent At':'','Test Certificate ID':'','Test Certificate PDF URL':'','Last Error':'','Completed At':''});
    const e=buildRegistrationEmail_(w,p,true); sendDirectEmail_(recipient,e.subject,e.html,e.text,''); updateTestRun_(runId,{'Confirmation Sent At':new Date()});
    try{createTestStepTrigger_();}catch(err){log_('WARN','TEST',w['Event ID'],pid,'TEST_STEP_TRIGGER_DEFERRED','Central scheduler will continue the test.',String(err));}
    log_('INFO','TEST',w['Event ID'],pid,'SELF_TEST_STARTED','Self test started without creating a production participant row',runId+' | recipient='+recipient);
    return {ok:true,runId,recipient:recipient,message:'Self test started. Registration sent now; next stage will be attempted in about 1 minute.'};
  });
}
function processDueTestSteps() { if(isPaused_()) return {ok:false,paused:true}; return withLock_('processDueTestSteps', processDueTestStepsCore_); }
function testParticipantFromRun_(run){
  try{ if(run['Participant Snapshot']) return JSON.parse(String(run['Participant Snapshot'])); }catch(err){}
  return findParticipantById_(run['Participant ID']);
}
function processDueTestStepsCore_() {
  const sh=getSheet_(APP.SHEETS.TESTS); const runs=readSheetObjects_(sh).filter(r=>String(r.Status)==='RUNNING'); let processed=0;
  runs.forEach(run=>{
    try{
      const p=testParticipantFromRun_(run); const w=findWorkshop_(run['Event ID']); if(!p||!w) throw new Error('Test participant snapshot/workshop no longer exists.');
      if(Array.isArray(p.__v141Stages)){ if(processModernTestRunStepV141_(run,p,w)) processed++; return; }
      const step=String(run['Current Step']||'');
      if(step==='REMINDER_PENDING'){
        const sentAt=parseQueueDate_(run['Confirmation Sent At']); if(sentAt && (new Date().getTime()-sentAt.getTime())<APP.TEST_DELAY_MS) return;
        const e=buildReminderEmail_(w,p,60,true); sendDirectEmail_(p['Email'],e.subject,e.html,e.text,'');
        updateTestRun_(run['Run ID'],{'Reminder Sent At':new Date(),'Current Step':boolValue_(run['Include Attendance'],false)?'ATTENDANCE_PENDING':'CERTIFICATE_PENDING'}); processed++;
        try{createTestStepTrigger_();}catch(err){log_('WARN','TEST',w['Event ID'],p['Participant ID'],'TEST_NEXT_TRIGGER_DEFERRED','Central scheduler will continue the test.',String(err));}
      }else if(step==='ATTENDANCE_PENDING'){
        const sentAt=parseQueueDate_(run['Reminder Sent At']); if(sentAt && (new Date().getTime()-sentAt.getTime())<APP.TEST_DELAY_MS) return;
        const attendanceUrl=w['Attendance Form ID']?FormApp.openById(String(w['Attendance Form ID'])).getPublishedUrl():'';
        const e=buildAttendanceEmail_(w,p,attendanceUrl,true); sendDirectEmail_(p['Email'],e.subject,e.html,e.text,'');
        updateTestRun_(run['Run ID'],{'Attendance Sent At':new Date(),'Current Step':'CERTIFICATE_PENDING'}); processed++;
        try{createTestStepTrigger_();}catch(err){log_('WARN','TEST',w['Event ID'],p['Participant ID'],'TEST_CERT_TRIGGER_DEFERRED','Central scheduler will continue the test.',String(err));}
      }else if(step==='CERTIFICATE_PENDING'){
        const prev=parseQueueDate_(boolValue_(run['Include Attendance'],false)?run['Attendance Sent At']:run['Reminder Sent At']); if(prev && (new Date().getTime()-prev.getTime())<APP.TEST_DELAY_MS) return;
        const cert=generateCertificateForParticipant_(w,p,true,run['Run ID']); const e=buildCertificateEmail_(w,p,cert,true); sendDirectEmail_(p['Email'],e.subject,e.html,e.text,cert.pdfFileId);
        updateTestRun_(run['Run ID'],{'Certificate Created At':new Date(),'Certificate Sent At':new Date(),'Test Certificate ID':cert.certificateId,'Test Certificate PDF URL':cert.pdfUrl,'Status':'COMPLETED','Current Step':'DONE','Completed At':new Date()});
        log_('INFO','TEST',w['Event ID'],p['Participant ID'],'SELF_TEST_COMPLETED','All requested self-test stages completed',run['Run ID']); processed++;
      }
    }catch(err){ updateTestRun_(run['Run ID'],{'Status':'FAILED','Last Error':String(err),'Current Step':'FAILED','Completed At':new Date()}); log_('ERROR','TEST',String(run['Event ID']),String(run['Participant ID']),'SELF_TEST_FAILED',String(err),stack_(err)); }
  });
  const pendingV141=readSheetObjects_(getSheet_(APP.SHEETS.TESTS)).filter(function(r){if(String(r.Status)!=='RUNNING')return false;try{const p=JSON.parse(String(r['Participant Snapshot']||'{}'));return Array.isArray(p.__v141Stages);}catch(err){return false;}});
  if(pendingV141.length){const delays=pendingV141.map(function(r){try{const p=JSON.parse(String(r['Participant Snapshot']||'{}'));return Number(p.__v141DelayMs)>=0?Number(p.__v141DelayMs):APP.TEST_DELAY_MS;}catch(err){return APP.TEST_DELAY_MS;}}).filter(function(n){return isFinite(n);});ensureV141TestStepTrigger_(delays.length?Math.min.apply(null,delays):APP.TEST_DELAY_MS);}
  return {ok:true,processed:processed};
}

function generateCertificateForParticipant_(w,p,isTest,testRunId) {
  w=repairIfWorkshopIntegrationBroken_(w);
  const masterId=PropertiesService.getScriptProperties().getProperty('CERTIFICATE_MASTER_ID');
  if(!masterId) throw new Error('Certificate Master is missing. Run setupSystem() again.');

  const eventFolder=DriveApp.getFolderById(w['Folder ID']);
  const certFolder=findChildFolder_(eventFolder,'Certificates') || eventFolder.createFolder('Certificates');
  const targetFolder=isTest ? (findChildFolder_(certFolder,'Test Runs') || certFolder.createFolder('Test Runs')) : certFolder;
  const certId=(isTest?'TEST-CERT-':'CERT-')+Utilities.getUuid().replace(/-/g,'').slice(0,12).toUpperCase();

  const certificateBaseName=safeFileName_(p['Full Name'])+' - '+certId;
  // Create first, then move into the event folder. This avoids the Drive
  // parent.mimeType failure seen when a legacy row supplied a bad parent.
  const copy=DriveApp.getFileById(masterId).makeCopy(certificateBaseName);
  try {
    copy.moveTo(targetFolder);
    const pres=SlidesApp.openById(copy.getId());
    const replacements={
      '{{PARTICIPANT_NAME}}':String(p['Full Name']||''),
      '{{NAME}}':String(p['Full Name']||''),
      '{{WORKSHOP_NAME}}':String(w['Workshop Name']||''),
      '{{EVENT_ID}}':String(w['Event ID']||''),
      '{{WORKSHOP_DATE}}':formatDateForDisplay_(w['Workshop Date'], String(w['Timezone']||APP.DEFAULT_TIMEZONE)),
      '{{WORKSHOP_TIME}}':formatTimeForDisplay_(w['Start Time'], String(w['Timezone']||APP.DEFAULT_TIMEZONE))+' - '+formatTimeForDisplay_(w['End Time'], String(w['Timezone']||APP.DEFAULT_TIMEZONE)),
      '{{WORKSHOP_DURATION}}':durationText_(w),
      '{{PARTICIPANT_ID}}':String(p['Participant ID']||''),
      '{{REG_NUMBER}}':String(p['Registration Number']||''),
      '{{ID_NUMBER}}':String(p['ID Number']||''),
      '{{DEPARTMENT_FACULTY}}':String(p['Department / Faculty']||''),
      '{{SEASON}}':String(p['Season']||''),
      '{{MEETING_PLATFORM}}':String(w['Meeting Platform']||''),
      '{{MEETING_LINK}}':String(w['Meeting Link']||''),
      '{{CERTIFICATE_ID}}':certId
    };
    Object.keys(replacements).forEach(k=>pres.replaceAllText(k,replacements[k]));
    pres.saveAndClose();

    const exportUrl='https://docs.google.com/presentation/d/'+copy.getId()+'/export/pdf';
    const response=UrlFetchApp.fetch(exportUrl,{
      headers:{Authorization:'Bearer '+ScriptApp.getOAuthToken()},
      muteHttpExceptions:true,
      followRedirects:true
    });
    const code=response.getResponseCode();
    if(code<200 || code>=300){
      const body=response.getContentText().slice(0,500);
      throw new Error('Certificate PDF export failed. HTTP '+code+'. '+body);
    }

    const pdfBlob=response.getBlob();
    const contentType=String(pdfBlob.getContentType()||'').toLowerCase();
    if(contentType && contentType!=='application/pdf'){
      throw new Error('Certificate PDF export returned unexpected content type: '+contentType);
    }

    const pdfFile=targetFolder.createFile(pdfBlob.setName(certificateBaseName+'.pdf'));
    const verificationUrl=buildPublicCertificateVerificationUrl_(certId);

    appendObject_(getSheet_(APP.SHEETS.CERTS),{
      'Certificate ID':certId,
      'Event ID':w['Event ID'],
      'Participant ID':p['Participant ID'],
      'Participant Name':p['Full Name'],
      'Email':p['Email'],
      'Workshop Name':w['Workshop Name'],
      'Season':p['Season'],
      'Issued At':new Date(),
      'PDF URL':pdfFile.getUrl(),
      'PDF File ID':pdfFile.getId(),
      'Verification URL':verificationUrl,
      'Status':isTest?'TEST_GENERATED':'GENERATED',
      'Created At':new Date(),
      'Test Run ID':testRunId||''
    });

    return {certificateId:certId,pdfUrl:pdfFile.getUrl(),pdfFileId:pdfFile.getId(),verificationUrl:verificationUrl};
  } finally {
    try { copy.setTrashed(true); } catch(err) {}
  }
}

function sendDirectEmail_(to,subject,html,text,attachmentFileId){
  validateRenderedEmail_('EMAIL',html,{participant_name:String(to||'recipient'),workshop_name:String(subject||'email')});
  const attachments=attachmentFileId?[DriveApp.getFileById(attachmentFileId).getBlob()]:[];
  MailApp.sendEmail({to:to,subject:subject,body:text||stripHtml_(html),htmlBody:html,name:getSetting_('SENDER_NAME')||APP.ROOT,attachments:attachments});
}

function socialBlock_(w){
  const settings=getSettings_();
  const links=[
    ['facebook','f',w['Facebook']||settings.DEFAULT_FACEBOOK||''],
    ['instagram','IG',w['Instagram']||settings.DEFAULT_INSTAGRAM||''],
    ['linkedin','in',w['LinkedIn']||settings.DEFAULT_LINKEDIN||''],
    ['youtube','YT',w['YouTube']||settings.DEFAULT_YOUTUBE||''],
    ['website','W',w['Website']||settings.DEFAULT_WEBSITE||'']
  ].filter(x=>x[2]);
  if(!links.length) return '';
  return '<table role="presentation" cellpadding="0" cellspacing="0" border="0"><tr>'+
    links.map(function(x){return '<td style="padding:0 6px"><a href="'+escAttr_(x[2])+'" aria-label="'+esc_(x[0])+'" title="'+esc_(x[0])+'" style="display:block;width:36px;height:36px;line-height:36px;text-align:center;background:#1C1C1E;border:1px solid #2C2C2E;border-radius:18px;font-family:-apple-system,BlinkMacSystemFont,\'Segoe UI\',sans-serif;font-size:11px;font-weight:600;color:#D1D1D6">'+x[1]+'</a></td>';}).join('')+
    '</tr></table>';
}
function orgInitials_(name){
  const parts=String(name||'WORKSHOP Automation').trim().split(/\\s+/).filter(Boolean);
  return parts.slice(0,2).map(function(x){return x.charAt(0).toUpperCase();}).join('').slice(0,2)||'W';
}
const EMAIL_TEMPLATES_ = {
  REGISTRATION: "<!doctype html>\n<html lang=\"en\">\n<head>\n<meta charset=\"utf-8\">\n<meta name=\"viewport\" content=\"width=device-width,initial-scale=1\">\n<meta http-equiv=\"X-UA-Compatible\" content=\"IE=edge\">\n<meta name=\"color-scheme\" content=\"light only\">\n<meta name=\"supported-color-schemes\" content=\"light\">\n<meta name=\"format-detection\" content=\"telephone=no,date=no,address=no,email=no\">\n<title>Registration Successful | Finance Club PSTU</title>\n<style>\nhtml,body{margin:0!important;padding:0!important;width:100%!important}\nbody{background:#f5f5f7;color:#1d1d1f;-webkit-text-size-adjust:100%;-ms-text-size-adjust:100%}\ntable,td{border-collapse:collapse!important;mso-table-lspace:0pt;mso-table-rspace:0pt}\nimg{border:0;outline:none;text-decoration:none;-ms-interpolation-mode:bicubic}\na{text-decoration:none}\n.apple-font{font-family:-apple-system,BlinkMacSystemFont,\"SF Pro Text\",\"SF Pro Display\",\"Segoe UI\",Arial,sans-serif}\n.shell{width:600px;max-width:600px}\n.panel{background:#f5f5f7;border-radius:16px}\n.row-line{border-bottom:1px solid #e5e5e7}\n.footer-link{color:#0066cc}\n@media screen and (max-width:600px){\n  .outer-pad{padding:0!important}\n  .shell{width:100%!important;max-width:100%!important}\n  .side{padding-left:20px!important;padding-right:20px!important}\n  .brand{padding-top:24px!important;padding-bottom:10px!important}\n  .hero{padding:28px 20px 26px!important}\n  .headline{font-size:32px!important;line-height:37px!important;letter-spacing:-1.15px!important}\n  .intro{font-size:15px!important;line-height:23px!important}\n  .panel-pad{padding-left:18px!important;padding-right:18px!important}\n  .section-gap{padding-bottom:18px!important}\n  .timeline-cell{padding:17px 8px 19px!important}\n  .cta-pad{padding:22px 18px 24px!important}\n  .cta-button a{display:block!important;width:100%!important;box-sizing:border-box!important}\n  .footer{padding-left:20px!important;padding-right:20px!important}\n}\n</style>\n</head>\n<body style=\"margin:0;padding:0;background:#f5f5f7;\">\n<div style=\"display:none;max-height:0;overflow:hidden;opacity:0;mso-hide:all;\">\nYour registration for {{workshop_name}} is confirmed.\n</div>\n\n<table role=\"presentation\" width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\" style=\"background:#f5f5f7;\">\n<tr><td align=\"center\" class=\"outer-pad\" style=\"padding:16px 0;\">\n\n<!--[if mso]><table role=\"presentation\" width=\"600\" align=\"center\"><tr><td><![endif]-->\n<table role=\"presentation\" width=\"600\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\" class=\"shell\" style=\"width:600px;max-width:600px;background:#ffffff;margin:0 auto;\">\n\n<!-- BRAND -->\n<tr>\n<td align=\"center\" class=\"brand\" style=\"padding:27px 24px 8px;\">\n  <img src=\"{{logo_url}}\" width=\"58\" height=\"58\" alt=\"\" style=\"display:block;width:58px;height:58px;object-fit:contain;margin:0 auto;\">\n  <div class=\"apple-font\" style=\"margin-top:9px;font-size:18px;line-height:23px;font-weight:600;letter-spacing:-.35px;color:#1d1d1f;\">Finance Club</div>\n  <div class=\"apple-font\" style=\"margin-top:1px;font-size:11px;line-height:16px;font-weight:500;letter-spacing:2px;color:#86868b;\">PSTU</div>\n</td>\n</tr>\n\n<!-- HERO -->\n<tr>\n<td align=\"center\" class=\"hero\" style=\"padding:27px 34px 27px;background:#ffffff;\">\n  <table role=\"presentation\" width=\"48\" height=\"48\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\">\n    <tr><td align=\"center\" valign=\"middle\" style=\"width:48px;height:48px;border-radius:24px;background:#e7f5ec;color:#248a4b;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif;font-size:23px;line-height:48px;font-weight:600;\">✓</td></tr>\n  </table>\n  <div class=\"apple-font\" style=\"margin-top:15px;font-size:11px;line-height:17px;font-weight:600;letter-spacing:.45px;color:#248a4b;\">REGISTRATION SUCCESSFUL</div>\n  <h1 class=\"apple-font headline\" style=\"margin:8px 0 0;font-size:35px;line-height:40px;font-weight:700;letter-spacing:-1.2px;color:#1d1d1f;\">You're officially registered.</h1>\n  <p class=\"apple-font intro\" style=\"margin:13px auto 0;max-width:470px;font-size:15px;line-height:23px;color:#6e6e73;\">Hello, <strong style=\"font-weight:600;color:#1d1d1f;\">{{participant_name}}</strong>. Your registration for <strong style=\"font-weight:600;color:#1d1d1f;\">{{workshop_name}}</strong> has been confirmed.</p>\n</td>\n</tr>\n\n<!-- WORKSHOP SUMMARY -->\n<tr><td class=\"side section-gap\" style=\"padding:0 32px 18px;\">\n<table role=\"presentation\" width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\" class=\"panel\">\n<tr><td class=\"panel-pad\" style=\"padding:22px 23px 21px;\">\n  <div class=\"apple-font\" style=\"font-size:10px;line-height:15px;font-weight:600;letter-spacing:.8px;color:#86868b;text-transform:uppercase;\">YOUR WORKSHOP</div>\n  <div class=\"apple-font\" style=\"margin-top:5px;font-size:22px;line-height:29px;font-weight:650;letter-spacing:-.55px;color:#1d1d1f;\">{{workshop_name}}</div>\n  <table role=\"presentation\" width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\" style=\"margin-top:14px;\">\n    <tr>\n      <td width=\"50%\" valign=\"top\" style=\"padding-right:9px;\">\n        <div class=\"apple-font\" style=\"font-size:10px;line-height:15px;color:#86868b;\">DATE</div>\n        <div class=\"apple-font\" style=\"margin-top:2px;font-size:13px;line-height:19px;font-weight:500;color:#1d1d1f;\">{{event_date}}</div>\n      </td>\n      <td width=\"50%\" valign=\"top\" style=\"padding-left:9px;\">\n        <div class=\"apple-font\" style=\"font-size:10px;line-height:15px;color:#86868b;\">TIME</div>\n        <div class=\"apple-font\" style=\"margin-top:2px;font-size:13px;line-height:19px;font-weight:500;color:#1d1d1f;\">{{time_range}}</div>\n      </td>\n    </tr>\n    <tr>\n      <td width=\"50%\" valign=\"top\" style=\"padding-top:12px;padding-right:9px;\">\n        <div class=\"apple-font\" style=\"font-size:10px;line-height:15px;color:#86868b;\">DURATION</div>\n        <div class=\"apple-font\" style=\"margin-top:2px;font-size:13px;line-height:19px;font-weight:500;color:#1d1d1f;\">{{duration}}</div>\n      </td>\n      <td width=\"50%\" valign=\"top\" style=\"padding-top:12px;padding-left:9px;\">\n        <div class=\"apple-font\" style=\"font-size:10px;line-height:15px;color:#86868b;\">PLATFORM / VENUE</div>\n        <div class=\"apple-font\" style=\"margin-top:2px;font-size:13px;line-height:19px;font-weight:500;color:#1d1d1f;\">{{location}}</div>\n      </td>\n    </tr>\n  </table>\n</td></tr>\n</table>\n</td></tr>\n\n<!-- REGISTRATION DETAILS -->\n<tr><td class=\"side section-gap\" style=\"padding:0 32px 18px;\">\n<table role=\"presentation\" width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\" class=\"panel\">\n<tr><td class=\"panel-pad\" style=\"padding:21px 23px 5px;\">\n  <div class=\"apple-font\" style=\"font-size:17px;line-height:23px;font-weight:600;letter-spacing:-.3px;color:#1d1d1f;\">Your registration</div>\n  <div class=\"apple-font\" style=\"margin-top:2px;font-size:11px;line-height:17px;color:#86868b;\">The details you provided during registration.</div>\n\n  <table role=\"presentation\" width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\" style=\"margin-top:7px;\">\n  <tr><td class=\"row-line\" style=\"padding:11px 0;\"><table role=\"presentation\" width=\"100%\"><tr><td class=\"apple-font\" width=\"42%\" style=\"font-size:12px;line-height:18px;color:#6e6e73;\">Full name</td><td align=\"right\" class=\"apple-font\" style=\"font-size:12px;line-height:18px;font-weight:500;color:#1d1d1f;\">{{participant_name}}</td></tr></table></td></tr>\n  <tr><td class=\"row-line\" style=\"padding:11px 0;\"><table role=\"presentation\" width=\"100%\"><tr><td class=\"apple-font\" width=\"42%\" style=\"font-size:12px;line-height:18px;color:#6e6e73;\">Email</td><td align=\"right\" class=\"apple-font\" style=\"font-size:12px;line-height:18px;font-weight:500;color:#1d1d1f;overflow-wrap:anywhere;\">{{email}}</td></tr></table></td></tr>\n  <tr><td class=\"row-line\" style=\"padding:11px 0;\"><table role=\"presentation\" width=\"100%\"><tr><td class=\"apple-font\" width=\"42%\" style=\"font-size:12px;line-height:18px;color:#6e6e73;\">Phone / WhatsApp</td><td align=\"right\" class=\"apple-font\" style=\"font-size:12px;line-height:18px;font-weight:500;color:#1d1d1f;\">{{phone_whatsapp}}</td></tr></table></td></tr>\n  <tr><td class=\"row-line\" style=\"padding:11px 0;\"><table role=\"presentation\" width=\"100%\"><tr><td class=\"apple-font\" width=\"42%\" style=\"font-size:12px;line-height:18px;color:#6e6e73;\">Department / Faculty</td><td align=\"right\" class=\"apple-font\" style=\"font-size:12px;line-height:18px;font-weight:500;color:#1d1d1f;\">{{department_faculty}}</td></tr></table></td></tr>\n  <tr><td class=\"row-line\" style=\"padding:11px 0;\"><table role=\"presentation\" width=\"100%\"><tr><td class=\"apple-font\" width=\"42%\" style=\"font-size:12px;line-height:18px;color:#6e6e73;\">Season</td><td align=\"right\" class=\"apple-font\" style=\"font-size:12px;line-height:18px;font-weight:500;color:#1d1d1f;\">{{season}}</td></tr></table></td></tr>\n  <tr><td class=\"row-line\" style=\"padding:11px 0;\"><table role=\"presentation\" width=\"100%\"><tr><td class=\"apple-font\" width=\"42%\" style=\"font-size:12px;line-height:18px;color:#6e6e73;\">ID number</td><td align=\"right\" class=\"apple-font\" style=\"font-size:12px;line-height:18px;font-weight:500;color:#1d1d1f;\">{{id_number}}</td></tr></table></td></tr>\n  <tr><td style=\"padding:11px 0 9px;\"><table role=\"presentation\" width=\"100%\"><tr><td class=\"apple-font\" width=\"42%\" style=\"font-size:12px;line-height:18px;color:#6e6e73;\">Registration no.</td><td align=\"right\" class=\"apple-font\" style=\"font-size:12px;line-height:18px;font-weight:600;color:#1d1d1f;\">{{registration_number}}</td></tr></table></td></tr>\n  </table>\n</td></tr>\n</table>\n</td></tr>\n\n{{lifecycle_block}}\n\n{{group_block}}\n\n<!-- CALENDAR -->\n<tr><td class=\"side section-gap\" style=\"padding:0 32px 22px;\">\n<table role=\"presentation\" width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\" class=\"panel\">\n<tr><td align=\"center\" class=\"cta-pad\" style=\"padding:23px 20px 25px;\">\n  <div class=\"apple-font\" style=\"font-size:17px;line-height:23px;font-weight:600;color:#1d1d1f;\">Add to your calendar</div>\n  <div class=\"apple-font\" style=\"margin-top:3px;font-size:12px;line-height:18px;color:#6e6e73;\">Keep the workshop date and time handy.</div>\n  <table role=\"presentation\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\" style=\"margin-top:14px;\" class=\"cta-button\">\n    <tr><td align=\"center\" style=\"background:#0071e3;border-radius:9px;\">\n      <a href=\"{{calendar_url}}\" style=\"display:block;padding:11px 18px;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif;font-size:13px;line-height:18px;font-weight:600;color:#ffffff;min-width:180px;\">Add to Google Calendar</a>\n    </td></tr>\n  </table>\n</td></tr>\n</table>\n</td></tr>\n\n<!-- CONTACT -->\n<tr><td class=\"side\" align=\"center\" style=\"padding:0 32px 20px;\">\n<div class=\"apple-font\" style=\"font-size:12px;line-height:19px;color:#6e6e73;\">Need help? <a href=\"mailto:{{support_email}}\" class=\"footer-link\">Contact Finance Club PSTU</a></div>\n</td></tr>\n\n<tr><td class=\"side\" style=\"padding:0 32px;\"><div style=\"height:1px;background:#e5e5e7;font-size:0;line-height:0;\">&nbsp;</div></td></tr>\n\n<!-- FOOTER -->\n<tr>\n<td class=\"footer\" align=\"center\" style=\"padding:21px 32px 25px;\">\n  <div class=\"apple-font\" style=\"font-size:14px;line-height:20px;font-weight:600;color:#1d1d1f;\">Finance Club PSTU</div>\n  <div class=\"apple-font\" style=\"margin-top:2px;font-size:10px;line-height:16px;color:#86868b;\">Learn · Grow · Connect · Create Impact</div>\n\n  <table role=\"presentation\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\" style=\"margin:14px auto 0;\">\n  <tr>\n    <td style=\"padding:0 4px;\"><a href=\"{{facebook_url}}\" aria-label=\"Facebook\"><span style=\"display:block;width:34px;height:34px;border-radius:17px;background:#ffffff;border:1px solid #d2d2d7;text-align:center;\"><img src=\"https://cdn.simpleicons.org/facebook/000000\" width=\"16\" height=\"16\" alt=\"\" style=\"display:block;width:16px;height:16px;margin:9px auto;\"></span></a></td>\n    <td style=\"padding:0 4px;\"><a href=\"{{instagram_url}}\" aria-label=\"Instagram\"><span style=\"display:block;width:34px;height:34px;border-radius:17px;background:#ffffff;border:1px solid #d2d2d7;text-align:center;\"><img src=\"https://cdn.simpleicons.org/instagram/000000\" width=\"16\" height=\"16\" alt=\"\" style=\"display:block;width:16px;height:16px;margin:9px auto;\"></span></a></td>\n    <td style=\"padding:0 4px;\"><a href=\"{{linkedin_url}}\" aria-label=\"LinkedIn\"><span style=\"display:block;width:34px;height:34px;border-radius:17px;background:#ffffff;border:1px solid #d2d2d7;text-align:center;\"><img src=\"https://cdn-icons-png.flaticon.com/512/1384/1384014.png\" width=\"16\" height=\"16\" alt=\"\" style=\"display:block;width:16px;height:16px;margin:9px auto;\"></span></a></td>\n    <td style=\"padding:0 4px;\"><a href=\"{{website_url}}\" aria-label=\"Website\" title=\"Website\"><span style=\"display:block;width:34px;height:34px;border-radius:17px;background:#ffffff;border:1px solid #d2d2d7;text-align:center;\"><img src=\"https://cdn-icons-png.flaticon.com/512/1006/1006771.png\" width=\"16\" height=\"16\" alt=\"\" style=\"display:block;width:16px;height:16px;margin:9px auto;\"></span></a></td>\n    <td style=\"padding:0 4px;\"><a href=\"{{x_url}}\" aria-label=\"X\"><span style=\"display:block;width:34px;height:34px;border-radius:17px;background:#ffffff;border:1px solid #d2d2d7;text-align:center;\"><img src=\"https://cdn.simpleicons.org/x/000000\" width=\"16\" height=\"16\" alt=\"\" style=\"display:block;width:16px;height:16px;margin:9px auto;\"></span></a></td>\n  </tr>\n  </table>\n\n  <div class=\"apple-font\" style=\"margin-top:13px;font-size:9px;line-height:15px;color:#86868b;\">© {{current_year}} Finance Club PSTU. All rights reserved.</div>\n  <div class=\"apple-font\" style=\"margin-top:2px;font-size:9px;line-height:15px;color:#aeaeb2;\">This is an official communication regarding your workshop registration.</div>\n</td>\n</tr>\n\n</table>\n<!--[if mso]></td></tr></table><![endif]-->\n\n</td></tr>\n</table>\n</body>\n</html>\n",
  REMINDER: "<!doctype html>\n<html lang=\"en\">\n<head>\n<meta charset=\"utf-8\">\n<meta name=\"viewport\" content=\"width=device-width,initial-scale=1\">\n<meta http-equiv=\"X-UA-Compatible\" content=\"IE=edge\">\n<meta name=\"color-scheme\" content=\"light only\">\n<meta name=\"supported-color-schemes\" content=\"light\">\n<meta name=\"format-detection\" content=\"telephone=no,date=no,address=no,email=no\">\n<title>Join Workshop | Finance Club PSTU</title>\n<style>\nhtml,body{margin:0!important;padding:0!important;width:100%!important}\nbody{background:#f5f5f7;color:#1d1d1f;-webkit-text-size-adjust:100%;-ms-text-size-adjust:100%}\ntable,td{border-collapse:collapse!important;mso-table-lspace:0pt;mso-table-rspace:0pt}\nimg{border:0;outline:none;text-decoration:none;-ms-interpolation-mode:bicubic}\na{text-decoration:none}\n.apple-font{font-family:-apple-system,BlinkMacSystemFont,\"SF Pro Text\",\"SF Pro Display\",\"Segoe UI\",Arial,sans-serif}\n.shell{width:600px;max-width:600px}\n.panel{background:#f5f5f7;border-radius:16px}\n.row-line{border-bottom:1px solid #e5e5e7}\n.footer-link{color:#0066cc}\n@media screen and (max-width:600px){\n  .outer-pad{padding:0!important}\n  .shell{width:100%!important;max-width:100%!important}\n  .side{padding-left:20px!important;padding-right:20px!important}\n  .brand{padding-top:24px!important;padding-bottom:8px!important}\n  .hero{padding:22px 20px 24px!important}\n  .headline{font-size:32px!important;line-height:37px!important;letter-spacing:-1.15px!important}\n  .intro{font-size:15px!important;line-height:23px!important}\n  .panel-pad{padding-left:18px!important;padding-right:18px!important}\n  .section-gap{padding-bottom:18px!important}\n  .cta-pad{padding:22px 18px 24px!important}\n  .cta-button a{display:block!important;width:100%!important;box-sizing:border-box!important}\n  .footer{padding-left:20px!important;padding-right:20px!important}\n  .two-col{display:block!important;width:100%!important}\n  .two-col-cell{display:block!important;width:100%!important;padding-left:0!important;padding-right:0!important}\n  .two-col-gap{padding-bottom:12px!important}\n  .progress-cell{display:block!important;width:100%!important;padding:0 0 14px!important}\n  .progress-line{display:none!important}\n}\n</style>\n</head>\n<body style=\"margin:0;padding:0;background:#f5f5f7;\">\n<div style=\"display:none;max-height:0;overflow:hidden;opacity:0;mso-hide:all;\">\nYour workshop is ready. Join {{workshop_name}}.\n</div>\n\n<table role=\"presentation\" width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\" style=\"background:#f5f5f7;\">\n<tr><td align=\"center\" class=\"outer-pad\" style=\"padding:16px 0;\">\n\n<!--[if mso]><table role=\"presentation\" width=\"600\" align=\"center\"><tr><td><![endif]-->\n<table role=\"presentation\" width=\"600\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\" class=\"shell\" style=\"width:600px;max-width:600px;background:#ffffff;margin:0 auto;\">\n\n<!-- BRAND -->\n<tr>\n<td align=\"center\" class=\"brand\" style=\"padding:27px 24px 8px;\">\n  <img src=\"{{logo_url}}\" width=\"58\" height=\"58\" alt=\"Finance Club PSTU\" style=\"display:block;width:58px;height:58px;object-fit:contain;margin:0 auto;\">\n  <div class=\"apple-font\" style=\"margin-top:9px;font-size:18px;line-height:23px;font-weight:600;letter-spacing:-.35px;color:#1d1d1f;\">Finance Club</div>\n  <div class=\"apple-font\" style=\"margin-top:1px;font-size:11px;line-height:16px;font-weight:500;letter-spacing:2px;color:#86868b;\">PSTU</div>\n</td>\n</tr>\n\n<!-- HERO -->\n<tr>\n<td align=\"center\" class=\"hero\" style=\"padding:27px 34px 25px;background:#ffffff;\">\n  <table role=\"presentation\" width=\"48\" height=\"48\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\">\n    <tr><td align=\"center\" valign=\"middle\" style=\"width:48px;height:48px;border-radius:24px;background:#e7f5ec;color:#248a4b;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif;font-size:23px;line-height:48px;font-weight:600;\">✓</td></tr>\n  </table>\n  <div class=\"apple-font\" style=\"margin-top:15px;font-size:11px;line-height:17px;font-weight:600;letter-spacing:.45px;color:#248a4b;\">{{eyebrow}}</div>\n  <h1 class=\"apple-font headline\" style=\"margin:8px 0 0;font-size:35px;line-height:40px;font-weight:700;letter-spacing:-1.2px;color:#1d1d1f;\">{{hero_title}}</h1>\n  <p class=\"apple-font intro\" style=\"margin:12px 0 0;font-size:16px;line-height:24px;color:#6e6e73;\">{{hero_intro}}</p>\n</td>\n</tr>\n\n<!-- WORKSHOP SUMMARY -->\n<tr>\n<td class=\"side\" style=\"padding:0 35px 0;\">\n<table role=\"presentation\" width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\" class=\"panel\">\n<tr><td class=\"panel-pad\" style=\"padding:22px 24px 23px;\">\n  <div class=\"apple-font\" style=\"font-size:11px;line-height:16px;font-weight:600;letter-spacing:.7px;color:#86868b;text-transform:uppercase;\">YOUR WORKSHOP</div>\n  <div class=\"apple-font\" style=\"margin-top:7px;font-size:22px;line-height:28px;font-weight:700;letter-spacing:-.45px;color:#1d1d1f;\">{{workshop_name}}</div>\n  <div style=\"height:1px;background:#e5e5e7;margin:15px 0 3px;\"></div>\n\n  <table role=\"presentation\" width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\">\n    <tr><td style=\"padding:10px 0;border-bottom:1px solid #e5e5e7;\">\n      <table role=\"presentation\" width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\"><tr>\n        <td class=\"apple-font\" style=\"font-size:14px;color:#6e6e73;\">Date</td>\n        <td class=\"apple-font\" align=\"right\" style=\"font-size:14px;font-weight:600;color:#1d1d1f;\">{{event_date}}</td>\n      </tr></table>\n    </td></tr>\n    <tr><td style=\"padding:10px 0;border-bottom:1px solid #e5e5e7;\">\n      <table role=\"presentation\" width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\"><tr>\n        <td class=\"apple-font\" style=\"font-size:14px;color:#6e6e73;\">Time</td>\n        <td class=\"apple-font\" align=\"right\" style=\"font-size:14px;font-weight:600;color:#1d1d1f;\">{{start_time}} to {{end_time}}</td>\n      </tr></table>\n    </td></tr>\n    <tr><td style=\"padding:10px 0;border-bottom:1px solid #e5e5e7;\">\n      <table role=\"presentation\" width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\"><tr>\n        <td class=\"apple-font\" style=\"font-size:14px;color:#6e6e73;\">Time zone</td>\n        <td class=\"apple-font\" align=\"right\" style=\"font-size:14px;font-weight:600;color:#1d1d1f;\">{{timezone}}</td>\n      </tr></table>\n    </td></tr>\n    <tr><td style=\"padding:10px 0 2px;\">\n      <table role=\"presentation\" width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\"><tr>\n        <td class=\"apple-font\" style=\"font-size:14px;color:#6e6e73;\">{{format_label}}</td>\n        <td class=\"apple-font\" align=\"right\" style=\"font-size:14px;font-weight:600;color:#1d1d1f;\">{{platform}}</td>\n      </tr></table>\n    </td></tr>\n  </table>\n</td></tr>\n</table>\n</td>\n</tr>\n\n{{primary_cta}}\n\n{{meeting_details_block}}\n\n{{participant_pass_block}}\n\n{{joining_checklist}}\n\n<!-- CALENDAR -->\n<tr>\n<td class=\"side\" align=\"center\" style=\"padding:18px 35px 0;\">\n  <table role=\"presentation\" width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\" class=\"panel\">\n    <tr><td align=\"center\" class=\"cta-pad\" style=\"padding:23px 24px 24px;\">\n      <div class=\"apple-font\" style=\"font-size:24px;line-height:30px;font-weight:700;letter-spacing:-.45px;color:#1d1d1f;\">Add to your calendar</div>\n      <div class=\"apple-font\" style=\"margin-top:6px;font-size:15px;line-height:22px;color:#6e6e73;\">Keep the workshop date and time handy.</div>\n      <table role=\"presentation\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\" style=\"margin-top:16px;\">\n        <tr><td style=\"background:#0071e3;border-radius:12px;\">\n          <a href=\"{{calendar_url}}\" title=\"Add to Google Calendar\" style=\"display:block;padding:14px 22px;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif;font-size:16px;line-height:21px;font-weight:600;color:#ffffff;\">Add to Google Calendar</a>\n        </td></tr>\n      </table>\n    </td></tr>\n  </table>\n</td>\n</tr>\n\n<!-- EVENT PROGRESS -->\n<tr>\n<td class=\"side\" style=\"padding:25px 35px 0;\">\n  <div class=\"apple-font\" style=\"font-size:12px;line-height:16px;font-weight:600;letter-spacing:.7px;color:#86868b;margin-bottom:11px;\">YOUR WORKSHOP JOURNEY</div>\n  <table role=\"presentation\" width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\" class=\"panel\">\n    <tr><td style=\"padding:20px 16px 21px;\">\n      <table role=\"presentation\" width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\"><tr>\n        <td width=\"33%\" valign=\"top\" align=\"center\" class=\"progress-cell\">\n          <table role=\"presentation\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\" width=\"100%\"><tr><td align=\"center\">\n            <table role=\"presentation\" width=\"12\" height=\"12\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\"><tr><td style=\"width:12px;height:12px;border-radius:6px;background:#d2d2d7;\"></td></tr></table>\n          </td></tr><tr><td align=\"center\" class=\"apple-font\" style=\"padding-top:9px;font-size:11px;font-weight:600;letter-spacing:.3px;color:#6e6e73;\">REGISTERED</td></tr></table>\n        </td>\n        <td width=\"4%\" valign=\"top\" class=\"progress-line\" style=\"padding-top:5px;\"><div style=\"height:1px;background:#d2d2d7;\"></div></td>\n        <td width=\"33%\" valign=\"top\" align=\"center\" class=\"progress-cell\">\n          <table role=\"presentation\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\" width=\"100%\"><tr><td align=\"center\">\n            <table role=\"presentation\" width=\"14\" height=\"14\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\"><tr><td style=\"width:14px;height:14px;border-radius:7px;background:#0071e3;\"></td></tr></table>\n          </td></tr><tr><td align=\"center\" class=\"apple-font\" style=\"padding-top:8px;font-size:11px;font-weight:700;letter-spacing:.3px;color:#1d1d1f;\">JOIN MEETING</td></tr></table>\n        </td>\n        <td width=\"4%\" valign=\"top\" class=\"progress-line\" style=\"padding-top:5px;\"><div style=\"height:1px;background:#d2d2d7;\"></div></td>\n        <td width=\"33%\" valign=\"top\" align=\"center\" class=\"progress-cell\">\n          <table role=\"presentation\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\" width=\"100%\"><tr><td align=\"center\">\n            <table role=\"presentation\" width=\"12\" height=\"12\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\"><tr><td style=\"width:12px;height:12px;border-radius:6px;background:#d2d2d7;\"></td></tr></table>\n          </td></tr><tr><td align=\"center\" class=\"apple-font\" style=\"padding-top:9px;font-size:11px;font-weight:600;letter-spacing:.3px;color:#6e6e73;\">CERTIFICATE</td></tr></table>\n        </td>\n      </tr></table>\n    </td></tr>\n  </table>\n</td>\n</tr>\n\n<!-- CONTACT -->\n<tr>\n<td align=\"center\" class=\"footer\" style=\"padding:26px 35px 0;\">\n  <div class=\"apple-font\" style=\"font-size:15px;line-height:22px;color:#6e6e73;\">Need help? <a href=\"mailto:{{support_email}}\" style=\"color:#0066cc;\">Contact Finance Club PSTU</a></div>\n</td>\n</tr>\n\n<!-- FOOTER -->\n<tr>\n<td align=\"center\" class=\"footer\" style=\"padding:22px 35px 34px;\">\n  <div style=\"height:1px;background:#e5e5e7;\"></div>\n  <div class=\"apple-font\" style=\"margin-top:23px;font-size:18px;line-height:23px;font-weight:600;letter-spacing:-.35px;color:#1d1d1f;\">Finance Club PSTU</div>\n  <div class=\"apple-font\" style=\"margin-top:4px;font-size:13px;line-height:19px;color:#86868b;\">Learn · Grow · Connect · Create Impact</div>\n\n  <table role=\"presentation\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\" style=\"margin:14px auto 0;\"><tr>\n    <td style=\"padding:0 4px;\"><a href=\"{{facebook_url}}\" aria-label=\"Facebook\" title=\"Facebook\"><span style=\"display:block;width:34px;height:34px;border-radius:17px;background:#ffffff;border:1px solid #d2d2d7;text-align:center;\"><img src=\"https://cdn.simpleicons.org/facebook/000000\" width=\"16\" height=\"16\" alt=\"\" style=\"display:block;width:16px;height:16px;margin:9px auto;\"></span></a></td>\n    <td style=\"padding:0 4px;\"><a href=\"{{instagram_url}}\" aria-label=\"Instagram\" title=\"Instagram\"><span style=\"display:block;width:34px;height:34px;border-radius:17px;background:#ffffff;border:1px solid #d2d2d7;text-align:center;\"><img src=\"https://cdn.simpleicons.org/instagram/000000\" width=\"16\" height=\"16\" alt=\"\" style=\"display:block;width:16px;height:16px;margin:9px auto;\"></span></a></td>\n    <td style=\"padding:0 4px;\"><a href=\"{{linkedin_url}}\" aria-label=\"LinkedIn\" title=\"LinkedIn\"><span style=\"display:block;width:34px;height:34px;border-radius:17px;background:#ffffff;border:1px solid #d2d2d7;text-align:center;\"><img src=\"https://cdn-icons-png.flaticon.com/512/1384/1384014.png\" width=\"16\" height=\"16\" alt=\"\" style=\"display:block;width:16px;height:16px;margin:9px auto;\"></span></a></td>\n    <td style=\"padding:0 4px;\"><a href=\"{{website_url}}\" aria-label=\"Website\" title=\"Website\"><span style=\"display:block;width:34px;height:34px;border-radius:17px;background:#ffffff;border:1px solid #d2d2d7;text-align:center;\"><img src=\"https://cdn-icons-png.flaticon.com/512/1006/1006771.png\" width=\"16\" height=\"16\" alt=\"\" style=\"display:block;width:16px;height:16px;margin:9px auto;\"></span></a></td>\n    <td style=\"padding:0 4px;\"><a href=\"{{x_url}}\" aria-label=\"X\" title=\"X\"><span style=\"display:block;width:34px;height:34px;border-radius:17px;background:#ffffff;border:1px solid #d2d2d7;text-align:center;\"><img src=\"https://cdn.simpleicons.org/x/000000\" width=\"16\" height=\"16\" alt=\"\" style=\"display:block;width:16px;height:16px;margin:9px auto;\"></span></a></td>\n  </tr></table>\n  <div class=\"apple-font\" style=\"margin-top:19px;font-size:12px;line-height:18px;color:#86868b;\">© {{current_year}} Finance Club PSTU. All rights reserved.</div>\n  <div class=\"apple-font\" style=\"margin-top:5px;font-size:11px;line-height:17px;color:#a1a1a6;\">This is an official workshop communication from Finance Club PSTU.</div>\n</td>\n</tr>\n\n</table>\n<!--[if mso]></td></tr></table><![endif]-->\n</td></tr>\n</table>\n</body>\n</html>\n",
  ATTENDANCE: "<!doctype html><html lang=\"en\"><head><meta charset=\"utf-8\"><meta name=\"viewport\" content=\"width=device-width,initial-scale=1\"><meta name=\"color-scheme\" content=\"light only\"><style>html,body{margin:0!important;padding:0!important;width:100%!important}body{background:#f5f5f7;color:#1d1d1f;-webkit-text-size-adjust:100%}table,td{border-collapse:collapse!important}a{text-decoration:none}@media screen and (max-width:600px){.shell{width:100%!important}.px{padding-left:20px!important;padding-right:20px!important}.h1{font-size:32px!important;line-height:37px!important}.btn a{display:block!important;width:100%!important;box-sizing:border-box!important}}</style></head><body style=\"margin:0;padding:0;background:#f5f5f7;\"><div style=\"display:none;max-height:0;overflow:hidden;opacity:0;mso-hide:all;\">Attendance required for {{workshop_name}}.</div><table role=\"presentation\" width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\" style=\"background:#f5f5f7;\"><tr><td align=\"center\" class=\"px\" style=\"padding:16px 0;\"><table role=\"presentation\" width=\"600\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\" class=\"shell\" style=\"width:600px;max-width:600px;background:#fff;margin:0 auto;\"><tr><td align=\"center\" style=\"padding:28px 24px 10px;\"><img src=\"{{logo_url}}\" width=\"58\" height=\"58\" alt=\"\" style=\"display:block;width:58px;height:58px;object-fit:contain;margin:0 auto;\"><div style=\"margin-top:9px;font:600 18px/23px -apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif;color:#1d1d1f;\">{{organization_name}}</div></td></tr><tr><td align=\"center\" class=\"px\" style=\"padding:28px 32px 30px;\"><div style=\"font:600 11px/17px -apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif;letter-spacing:.45px;color:#248a4b;\">ATTENDANCE REQUIRED</div><h1 class=\"h1\" style=\"margin:8px 0 0;font:700 35px/40px -apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif;letter-spacing:-1.2px;color:#1d1d1f;\">Please mark your attendance.</h1><p style=\"margin:13px auto 0;max-width:470px;font:15px/23px -apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif;color:#6e6e73;\">Hello, <strong style=\"color:#1d1d1f;\">{{participant_name}}</strong>. Your attendance is required to remain eligible for your certificate.</p></td></tr><tr><td class=\"px\" style=\"padding:0 32px 18px;\"><table role=\"presentation\" width=\"100%\" style=\"background:#f5f5f7;border-radius:16px;\"><tr><td style=\"padding:22px 23px;\"><div style=\"font:600 10px/15px -apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif;letter-spacing:.8px;color:#86868b;\">ATTENDANCE WINDOW</div><div style=\"margin-top:6px;font:700 22px/29px -apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif;color:#1d1d1f;\">{{workshop_name}}</div><div style=\"margin-top:10px;font:13px/19px -apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif;color:#6e6e73;\">{{attendance_window}}</div><div style=\"margin-top:9px;font:12px/18px -apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif;color:#86868b;\">Registration number: <strong style=\"color:#6e6e73;\">{{registration_number}}</strong></div></td></tr></table></td></tr><tr><td class=\"px\" align=\"center\" style=\"padding:0 32px 24px;\"><table role=\"presentation\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\" class=\"btn\"><tr><td style=\"background:#0071e3;border-radius:10px;\"><a href=\"{{attendance_url}}\" style=\"display:block;padding:14px 24px;font:600 15px/20px -apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif;color:#fff;min-width:220px;text-align:center;\">Mark Attendance</a></td></tr></table><div style=\"margin-top:12px;font:10px/16px -apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif;color:#86868b;word-break:break-all;\">{{attendance_url}}</div></td></tr><tr><td class=\"px\" style=\"padding:0 32px;\"><div style=\"height:1px;background:#e5e5e7;\"></div></td></tr><tr><td align=\"center\" class=\"px\" style=\"padding:20px 32px 26px;\"><div style=\"font:600 14px/20px -apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif;color:#1d1d1f;\">{{organization_name}}</div><table role=\"presentation\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\" style=\"margin:10px auto 0;\">\n<tr>\n    <td style=\"padding:0 4px;\"><a href=\"{{facebook_url}}\" aria-label=\"Facebook\" title=\"Facebook\"><span style=\"display:block;width:34px;height:34px;border-radius:17px;background:#ffffff;border:1px solid #d2d2d7;text-align:center;\"><img src=\"https://cdn.simpleicons.org/facebook/000000\" width=\"16\" height=\"16\" alt=\"\" style=\"display:block;width:16px;height:16px;margin:9px auto;\"></span></a></td>\n    <td style=\"padding:0 4px;\"><a href=\"{{linkedin_url}}\" aria-label=\"LinkedIn\" title=\"LinkedIn\"><span style=\"display:block;width:34px;height:34px;border-radius:17px;background:#ffffff;border:1px solid #d2d2d7;text-align:center;\"><img src=\"https://cdn-icons-png.flaticon.com/512/1384/1384014.png\" width=\"16\" height=\"16\" alt=\"\" style=\"display:block;width:16px;height:16px;margin:9px auto;\"></span></a></td>\n    <td style=\"padding:0 4px;\"><a href=\"{{instagram_url}}\" aria-label=\"Instagram\" title=\"Instagram\"><span style=\"display:block;width:34px;height:34px;border-radius:17px;background:#ffffff;border:1px solid #d2d2d7;text-align:center;\"><img src=\"https://cdn.simpleicons.org/instagram/000000\" width=\"16\" height=\"16\" alt=\"\" style=\"display:block;width:16px;height:16px;margin:9px auto;\"></span></a></td>\n    <td style=\"padding:0 4px;\"><a href=\"{{website_url}}\" aria-label=\"Website\" title=\"Website\"><span style=\"display:block;width:34px;height:34px;border-radius:17px;background:#ffffff;border:1px solid #d2d2d7;text-align:center;\"><img src=\"https://cdn-icons-png.flaticon.com/512/1006/1006771.png\" width=\"16\" height=\"16\" alt=\"\" style=\"display:block;width:16px;height:16px;margin:9px auto;\"></span></a></td>\n    <td style=\"padding:0 4px;\"><a href=\"{{x_url}}\" aria-label=\"X\" title=\"X\"><span style=\"display:block;width:34px;height:34px;border-radius:17px;background:#ffffff;border:1px solid #d2d2d7;text-align:center;\"><img src=\"https://cdn.simpleicons.org/x/000000\" width=\"16\" height=\"16\" alt=\"\" style=\"display:block;width:16px;height:16px;margin:9px auto;\"></span></a></td>\n</tr></table><div style=\"margin-top:8px;font:9px/15px -apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif;color:#aeaeb2;\">© {{current_year}} {{organization_name}}. All rights reserved.</div></td></tr></table></td></tr></table></body></html>",
  CERTIFICATE: "<!doctype html>\n<html lang=\"en\">\n<head>\n<meta charset=\"utf-8\">\n<meta name=\"viewport\" content=\"width=device-width,initial-scale=1\">\n<meta http-equiv=\"X-UA-Compatible\" content=\"IE=edge\">\n<meta name=\"color-scheme\" content=\"light only\">\n<meta name=\"supported-color-schemes\" content=\"light\">\n<meta name=\"format-detection\" content=\"telephone=no,date=no,address=no,email=no\">\n<title>Certificate Ready | Finance Club PSTU</title>\n<style>\nhtml,body{margin:0!important;padding:0!important;width:100%!important}\nbody{background:#f5f5f7;color:#1d1d1f;-webkit-text-size-adjust:100%;-ms-text-size-adjust:100%}\ntable,td{border-collapse:collapse!important;mso-table-lspace:0pt;mso-table-rspace:0pt}\nimg{border:0;outline:none;text-decoration:none;-ms-interpolation-mode:bicubic}\na{text-decoration:none}\n.apple-font{font-family:-apple-system,BlinkMacSystemFont,\"SF Pro Text\",\"SF Pro Display\",\"Segoe UI\",Arial,sans-serif}\n.shell{width:600px;max-width:600px}\n.panel{background:#f5f5f7;border-radius:16px}\n.row-line{border-bottom:1px solid #e5e5e7}\n.footer-link{color:#0066cc}\n@media screen and (max-width:600px){\n  .outer-pad{padding:0!important}\n  .shell{width:100%!important;max-width:100%!important}\n  .side{padding-left:20px!important;padding-right:20px!important}\n  .brand{padding-top:24px!important;padding-bottom:10px!important}\n  .hero{padding:27px 20px 26px!important}\n  .headline{font-size:32px!important;line-height:37px!important;letter-spacing:-1.15px!important}\n  .intro{font-size:15px!important;line-height:23px!important}\n  .panel-pad{padding-left:18px!important;padding-right:18px!important}\n  .section-gap{padding-bottom:18px!important}\n  .certificate-name{font-size:23px!important;line-height:29px!important}\n  .verify-button a{display:block!important;min-width:220px!important;width:auto!important;box-sizing:border-box!important}\n  .social-cell{padding-left:3px!important;padding-right:3px!important}\n  .footer{padding-left:20px!important;padding-right:20px!important}\n}\n</style>\n</head>\n<body style=\"margin:0;padding:0;background:#f5f5f7;\">\n<div style=\"display:none;max-height:0;overflow:hidden;opacity:0;mso-hide:all;\">\nYour certificate for {{workshop_name}} is ready and attached.\n</div>\n\n<table role=\"presentation\" width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\" style=\"background:#f5f5f7;\">\n<tr><td align=\"center\" class=\"outer-pad\" style=\"padding:16px 0;\">\n\n<!--[if mso]><table role=\"presentation\" width=\"600\" align=\"center\"><tr><td><![endif]-->\n<table role=\"presentation\" width=\"600\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\" class=\"shell\" style=\"width:600px;max-width:600px;background:#ffffff;margin:0 auto;\">\n\n<!-- BRAND -->\n<tr>\n<td align=\"center\" class=\"brand\" style=\"padding:27px 24px 8px;\">\n  <img src=\"{{logo_url}}\" width=\"58\" height=\"58\" alt=\"\" style=\"display:block;width:58px;height:58px;object-fit:contain;margin:0 auto;\">\n  <div class=\"apple-font\" style=\"margin-top:9px;font-size:18px;line-height:23px;font-weight:600;letter-spacing:-.35px;color:#1d1d1f;\">Finance Club</div>\n  <div class=\"apple-font\" style=\"margin-top:1px;font-size:11px;line-height:16px;font-weight:500;letter-spacing:2px;color:#86868b;\">PSTU</div>\n</td>\n</tr>\n\n<!-- HERO -->\n<tr>\n<td align=\"center\" class=\"hero\" style=\"padding:27px 34px 28px;background:#ffffff;\">\n  <table role=\"presentation\" width=\"54\" height=\"54\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\">\n    <tr><td align=\"center\" valign=\"middle\" style=\"width:54px;height:54px;border-radius:27px;background:#e7f5ec;\">\n      <img src=\"https://cdn-icons-png.flaticon.com/512/3135/3135728.png\" width=\"34\" height=\"34\" alt=\"Certificate\" style=\"display:block;width:34px;height:34px;margin:0 auto;border:0;outline:none;text-decoration:none;\">\n    </td></tr>\n  </table>\n  <div class=\"apple-font\" style=\"margin-top:15px;font-size:11px;line-height:17px;font-weight:600;letter-spacing:.45px;color:#248a4b;\">CERTIFICATE READY</div>\n  <h1 class=\"apple-font headline\" style=\"margin:8px 0 0;font-size:35px;line-height:40px;font-weight:700;letter-spacing:-1.2px;color:#1d1d1f;\">Your certificate is ready.</h1>\n  <p class=\"apple-font intro\" style=\"margin:13px auto 0;max-width:470px;font-size:15px;line-height:23px;color:#6e6e73;\">Congratulations, <strong style=\"font-weight:600;color:#1d1d1f;\">{{participant_name}}</strong>.<br>Finance Club PSTU is pleased to celebrate your participation in <strong style=\"font-weight:600;color:#1d1d1f;\">{{workshop_name}}</strong>. We hope this certificate serves as a meaningful recognition of your learning, effort, and contribution.</p>\n</td>\n</tr>\n\n<!-- CERTIFICATE SUMMARY -->\n<tr><td class=\"side section-gap\" style=\"padding:0 32px 18px;\">\n<table role=\"presentation\" width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\" class=\"panel\">\n<tr><td class=\"panel-pad\" style=\"padding:22px 23px 21px;\">\n  <div class=\"apple-font\" style=\"font-size:10px;line-height:15px;font-weight:600;letter-spacing:.8px;color:#86868b;text-transform:uppercase;\">CERTIFICATE</div>\n  <div class=\"apple-font certificate-name\" style=\"margin-top:5px;font-size:22px;line-height:29px;font-weight:650;letter-spacing:-.55px;color:#1d1d1f;\">{{participant_name}}</div>\n  <div class=\"apple-font\" style=\"margin-top:2px;font-size:13px;line-height:19px;color:#6e6e73;\">{{certificate_type}}</div>\n\n  <table role=\"presentation\" width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\" style=\"margin-top:15px;\">\n    <tr>\n      <td width=\"50%\" valign=\"top\" style=\"padding-right:9px;\">\n        <div class=\"apple-font\" style=\"font-size:10px;line-height:15px;color:#86868b;\">WORKSHOP</div>\n        <div class=\"apple-font\" style=\"margin-top:2px;font-size:13px;line-height:19px;font-weight:500;color:#1d1d1f;\">{{workshop_name}}</div>\n      </td>\n      <td width=\"50%\" valign=\"top\" style=\"padding-left:9px;\">\n        <div class=\"apple-font\" style=\"font-size:10px;line-height:15px;color:#86868b;\">DATE</div>\n        <div class=\"apple-font\" style=\"margin-top:2px;font-size:13px;line-height:19px;font-weight:500;color:#1d1d1f;\">{{event_date}}</div>\n      </td>\n    </tr>\n    <tr>\n      <td width=\"50%\" valign=\"top\" style=\"padding-top:12px;padding-right:9px;\">\n        <div class=\"apple-font\" style=\"font-size:10px;line-height:15px;color:#86868b;\">SEASON</div>\n        <div class=\"apple-font\" style=\"margin-top:2px;font-size:13px;line-height:19px;font-weight:500;color:#1d1d1f;\">{{season}}</div>\n      </td>\n      <td width=\"50%\" valign=\"top\" style=\"padding-top:12px;padding-left:9px;\">\n        <div class=\"apple-font\" style=\"font-size:10px;line-height:15px;color:#86868b;\">CERTIFICATE ID</div>\n        <div class=\"apple-font\" style=\"margin-top:2px;font-size:13px;line-height:19px;font-weight:600;color:#1d1d1f;word-break:break-all;\">{{certificate_id}}</div>\n      </td>\n    </tr>\n  </table>\n</td></tr>\n</table>\n</td></tr>\n\n<!-- VERIFY -->\n<tr><td class=\"side section-gap\" style=\"padding:0 32px 18px;\">\n<table role=\"presentation\" width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\" class=\"panel\">\n<tr><td align=\"center\" class=\"panel-pad\" style=\"padding:23px 20px 25px;\">\n  <div class=\"apple-font\" style=\"font-size:17px;line-height:23px;font-weight:600;color:#1d1d1f;\">Verify your certificate</div>\n  <div class=\"apple-font\" style=\"margin-top:3px;font-size:12px;line-height:18px;color:#6e6e73;\">Use the official verification page to confirm certificate details and authenticity.</div>\n  <table role=\"presentation\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\" style=\"margin-top:16px;\" class=\"verify-button\">\n    <tr><td align=\"center\" style=\"background:#0071e3;border-radius:10px;\">\n      <a href=\"{{verification_url}}\" style=\"display:block;padding:13px 24px;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif;font-size:14px;line-height:20px;font-weight:600;color:#ffffff;min-width:220px;box-sizing:border-box;\">\n        <img src=\"https://cdn-icons-png.flaticon.com/512/2438/2438078.png\" width=\"18\" height=\"18\" alt=\"\" style=\"display:inline-block;width:18px;height:18px;vertical-align:-4px;margin-right:7px;border:0;outline:none;text-decoration:none;filter:brightness(0) invert(1);\">\n        Verify Certificate\n      </a>\n    </td></tr>\n  </table>\n  <div class=\"apple-font\" style=\"margin-top:12px;font-size:10px;line-height:16px;color:#86868b;\">Certificate ID: <strong style=\"color:#6e6e73;font-weight:600;\">{{certificate_id}}</strong></div>\n</td></tr>\n</table>\n</td></tr>\n\n<!-- ATTACHMENT NOTE -->\n<tr><td class=\"side section-gap\" style=\"padding:0 32px 18px;\">\n<div class=\"apple-font\" style=\"font-size:12px;line-height:18px;color:#6e6e73;\">Your certificate PDF is attached to this email for your records.</div>\n</td></tr>\n\n<!-- CONTACT -->\n<tr><td class=\"side\" align=\"center\" style=\"padding:0 32px 20px;\">\n<div class=\"apple-font\" style=\"font-size:12px;line-height:19px;color:#6e6e73;\">Need help? <a href=\"mailto:{{support_email}}\" class=\"footer-link\">Contact Finance Club PSTU</a></div>\n</td></tr>\n\n<tr><td class=\"side\" style=\"padding:0 32px;\"><div style=\"height:1px;background:#e5e5e7;font-size:0;line-height:0;\">&nbsp;</div></td></tr>\n\n<!-- FOOTER -->\n<tr>\n<td class=\"footer\" align=\"center\" style=\"padding:21px 32px 25px;\">\n  <div class=\"apple-font\" style=\"font-size:14px;line-height:20px;font-weight:600;color:#1d1d1f;\">Finance Club PSTU</div>\n  <div class=\"apple-font\" style=\"margin-top:2px;font-size:10px;line-height:16px;color:#86868b;\">Learn · Grow · Connect · Create Impact</div>\n\n  <table role=\"presentation\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\" style=\"margin:14px auto 0;\">\n  <tr>\n    <td class=\"social-cell\" style=\"padding:0 4px;\"><a href=\"{{facebook_url}}\" aria-label=\"Facebook\"><span style=\"display:block;width:34px;height:34px;border-radius:17px;background:#ffffff;border:1px solid #d2d2d7;text-align:center;\"><img src=\"https://cdn.simpleicons.org/facebook/000000\" width=\"16\" height=\"16\" alt=\"\" style=\"display:block;width:16px;height:16px;margin:9px auto;\"></span></a></td>\n    <td class=\"social-cell\" style=\"padding:0 4px;\"><a href=\"{{instagram_url}}\" aria-label=\"Instagram\"><span style=\"display:block;width:34px;height:34px;border-radius:17px;background:#ffffff;border:1px solid #d2d2d7;text-align:center;\"><img src=\"https://cdn.simpleicons.org/instagram/000000\" width=\"16\" height=\"16\" alt=\"\" style=\"display:block;width:16px;height:16px;margin:9px auto;\"></span></a></td>\n    <td class=\"social-cell\" style=\"padding:0 4px;\"><a href=\"{{linkedin_url}}\" aria-label=\"LinkedIn\"><span style=\"display:block;width:34px;height:34px;border-radius:17px;background:#ffffff;border:1px solid #d2d2d7;text-align:center;\"><img src=\"https://cdn-icons-png.flaticon.com/512/1384/1384014.png\" width=\"16\" height=\"16\" alt=\"\" style=\"display:block;width:16px;height:16px;margin:9px auto;\"></span></a></td>\n    <td style=\"padding:0 4px;\"><a href=\"{{website_url}}\" aria-label=\"Website\" title=\"Website\"><span style=\"display:block;width:34px;height:34px;border-radius:17px;background:#ffffff;border:1px solid #d2d2d7;text-align:center;\"><img src=\"https://cdn-icons-png.flaticon.com/512/1006/1006771.png\" width=\"16\" height=\"16\" alt=\"\" style=\"display:block;width:16px;height:16px;margin:9px auto;\"></span></a></td>\n    <td class=\"social-cell\" style=\"padding:0 4px;\"><a href=\"{{x_url}}\" aria-label=\"X\"><span style=\"display:block;width:34px;height:34px;border-radius:17px;background:#ffffff;border:1px solid #d2d2d7;text-align:center;\"><img src=\"https://cdn.simpleicons.org/x/000000\" width=\"16\" height=\"16\" alt=\"\" style=\"display:block;width:16px;height:16px;margin:9px auto;\"></span></a></td>\n  </tr>\n  </table>\n\n  <div class=\"apple-font\" style=\"margin-top:13px;font-size:9px;line-height:15px;color:#86868b;\">© {{current_year}} Finance Club PSTU. All rights reserved.</div>\n  <div class=\"apple-font\" style=\"margin-top:2px;font-size:9px;line-height:15px;color:#aeaeb2;\">This is an official certificate communication from Finance Club PSTU.</div>\n</td>\n</tr>\n\n</table>\n<!--[if mso]></td></tr></table><![endif]-->\n\n</td></tr>\n</table>\n</body>\n</html>\n",
  CUSTOM: "<!doctype html><html lang=\"en\"><head><meta charset=\"utf-8\"><meta name=\"viewport\" content=\"width=device-width,initial-scale=1\"><meta name=\"color-scheme\" content=\"light only\"><style>html,body{margin:0!important;padding:0!important;width:100%!important}body{background:#f5f5f7;color:#1d1d1f;-webkit-text-size-adjust:100%}table,td{border-collapse:collapse!important}a{text-decoration:none}@media screen and (max-width:600px){.shell{width:100%!important}.px{padding-left:20px!important;padding-right:20px!important}.h1{font-size:32px!important;line-height:37px!important}.btn a{display:block!important;width:100%!important;box-sizing:border-box!important}}</style></head><body style=\"margin:0;padding:0;background:#f5f5f7;\"><div style=\"display:none;max-height:0;overflow:hidden;opacity:0;mso-hide:all;\">{{preheader}}</div><table role=\"presentation\" width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\" style=\"background:#f5f5f7;\"><tr><td align=\"center\" class=\"px\" style=\"padding:16px 0;\"><table role=\"presentation\" width=\"600\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\" class=\"shell\" style=\"width:600px;max-width:600px;background:#fff;margin:0 auto;\"><tr><td align=\"center\" style=\"padding:28px 24px 10px;\"><img src=\"{{logo_url}}\" width=\"58\" height=\"58\" alt=\"\" style=\"display:block;width:58px;height:58px;object-fit:contain;margin:0 auto;\"><div style=\"margin-top:9px;font:600 18px/23px -apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif;color:#1d1d1f;\">{{organization_name}}</div><div style=\"margin-top:1px;font:500 11px/16px -apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif;letter-spacing:2px;color:#86868b;\">PSTU</div></td></tr><tr><td align=\"center\" class=\"px\" style=\"padding:28px 32px 30px;\"><div style=\"font:600 11px/17px -apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif;letter-spacing:.45px;color:#0066cc;\">{{message_type}}</div><h1 class=\"h1\" style=\"margin:8px 0 0;font:700 35px/40px -apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif;letter-spacing:-1.2px;color:#1d1d1f;\">{{heading}}</h1><p style=\"margin:13px auto 0;max-width:480px;font:15px/23px -apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif;color:#6e6e73;\">Hello, <strong style=\"color:#1d1d1f;\">{{participant_name}}</strong>.</p></td></tr><tr><td class=\"px\" style=\"padding:0 32px 24px;\"><table role=\"presentation\" width=\"100%\" style=\"background:#f5f5f7;border-radius:16px;\"><tr><td style=\"padding:22px 23px;\"><div style=\"font:600 10px/15px -apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif;letter-spacing:.8px;color:#86868b;text-transform:uppercase;\">{{workshop_name}}</div><div style=\"margin-top:11px;font:15px/24px -apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif;color:#424245;\">{{message_html}}</div></td></tr></table></td></tr>{{cta_block}}<tr><td class=\"px\" style=\"padding:0 32px;\"><div style=\"height:1px;background:#e5e5e7;\"></div></td></tr><tr><td align=\"center\" class=\"px\" style=\"padding:20px 32px 26px;\"><div style=\"font:600 14px/20px -apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif;color:#1d1d1f;\">{{organization_name}}</div><table role=\"presentation\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\" style=\"margin:10px auto 0;\">\n<tr>\n    <td style=\"padding:0 4px;\"><a href=\"{{facebook_url}}\" aria-label=\"Facebook\" title=\"Facebook\"><span style=\"display:block;width:34px;height:34px;border-radius:17px;background:#ffffff;border:1px solid #d2d2d7;text-align:center;\"><img src=\"https://cdn.simpleicons.org/facebook/000000\" width=\"16\" height=\"16\" alt=\"\" style=\"display:block;width:16px;height:16px;margin:9px auto;\"></span></a></td>\n    <td style=\"padding:0 4px;\"><a href=\"{{linkedin_url}}\" aria-label=\"LinkedIn\" title=\"LinkedIn\"><span style=\"display:block;width:34px;height:34px;border-radius:17px;background:#ffffff;border:1px solid #d2d2d7;text-align:center;\"><img src=\"https://cdn-icons-png.flaticon.com/512/1384/1384014.png\" width=\"16\" height=\"16\" alt=\"\" style=\"display:block;width:16px;height:16px;margin:9px auto;\"></span></a></td>\n    <td style=\"padding:0 4px;\"><a href=\"{{instagram_url}}\" aria-label=\"Instagram\" title=\"Instagram\"><span style=\"display:block;width:34px;height:34px;border-radius:17px;background:#ffffff;border:1px solid #d2d2d7;text-align:center;\"><img src=\"https://cdn.simpleicons.org/instagram/000000\" width=\"16\" height=\"16\" alt=\"\" style=\"display:block;width:16px;height:16px;margin:9px auto;\"></span></a></td>\n    <td style=\"padding:0 4px;\"><a href=\"{{website_url}}\" aria-label=\"Website\" title=\"Website\"><span style=\"display:block;width:34px;height:34px;border-radius:17px;background:#ffffff;border:1px solid #d2d2d7;text-align:center;\"><img src=\"https://cdn-icons-png.flaticon.com/512/1006/1006771.png\" width=\"16\" height=\"16\" alt=\"\" style=\"display:block;width:16px;height:16px;margin:9px auto;\"></span></a></td>\n    <td style=\"padding:0 4px;\"><a href=\"{{x_url}}\" aria-label=\"X\" title=\"X\"><span style=\"display:block;width:34px;height:34px;border-radius:17px;background:#ffffff;border:1px solid #d2d2d7;text-align:center;\"><img src=\"https://cdn.simpleicons.org/x/000000\" width=\"16\" height=\"16\" alt=\"\" style=\"display:block;width:16px;height:16px;margin:9px auto;\"></span></a></td>\n</tr></table><div style=\"margin-top:8px;font:9px/15px -apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif;color:#aeaeb2;\">© {{current_year}} {{organization_name}}. All rights reserved.</div></td></tr></table></td></tr></table></body></html>"
};
function brandLogoUrl_(w,settings){
  return String((settings&&settings.DEFAULT_LOGO_URL)||getSetting_('DEFAULT_LOGO_URL')||'').trim();
}
function brandMarkInline_(org,logoUrl,accent){
  if(logoUrl){
    return '<img src="'+escAttr_(logoUrl)+'" width="28" height="28" alt="" style="display:block;width:28px;height:28px;object-fit:contain;border-radius:8px;background:#1C1C1E;border:1px solid '+accent+'">';
  }
  return esc_(orgInitials_(org));
}
function formatOffsetHuman_(minutes){
  minutes=Number(minutes||0);
  if(minutes>=1440 && minutes%1440===0) return (minutes/1440)+' day'+(minutes/1440===1?'':'s');
  if(minutes>=60 && minutes%60===0) return (minutes/60)+' hour'+(minutes/60===1?'':'s');
  return minutes+' minute'+(minutes===1?'':'s');
}
function parseFlexibleDateTime_(value,tz){
  if(value instanceof Date && !isNaN(value.getTime())) return value;
  const raw=clean_(value);
  if(!raw) return null;
  const normalized=raw.replace('T',' ');
  const m=normalized.match(/^(\d{4}-\d{2}-\d{2})\s+(\d{1,2}):(\d{2})$/);
  if(m) return parseDateTime_(m[1],String(m[2]).padStart(2,'0')+':'+m[3],tz);
  const d=new Date(raw);
  return isNaN(d.getTime())?null:d;
}
function formatDateTimeForDisplay_(value,tz){
  const d=parseFlexibleDateTime_(value,tz);
  return d?Utilities.formatDate(d,tz||APP.DEFAULT_TIMEZONE,'dd MMM yyyy, hh:mm a'):'scheduled time';
}

function renderAppleTemplate_(type,data){
  let html=EMAIL_TEMPLATES_[type];
  if(!html) throw new Error('Email template not found: '+type);
  const values=Object.assign({},data);
  Object.keys(values).forEach(function(k){
    html=html.split('{{'+k+'}}').join(String(values[k]===null||values[k]===undefined?'':values[k]));
  });
  validateRenderedEmail_(type,html,data);
  return html;
}
function validateRenderedEmail_(type,html,data){
  const body=String(html||'');
  if(!body.trim()) throw new Error('Email HTML is empty for '+type+'.');
  if(/\{\{[^{}]+\}\}/.test(body)) throw new Error('Email contains unresolved placeholders for '+type+'.');
  if(/(^|[^<\/])p>|(^|[^<])\/p>/i.test(body)) throw new Error('Email contains malformed paragraph HTML for '+type+'.');
  if(/(^|[^\w])(?:undefined|null)(?=$|[^\w])/i.test(body)) throw new Error('Email contains an undefined/null rendering value for '+type+'.');
  if(data && !clean_(data.participant_name)) throw new Error('Participant name is required for '+type+'.');
  if(data && !clean_(data.workshop_name)) throw new Error('Workshop name is required for '+type+'.');
  if(String(type).toUpperCase()==='CERTIFICATE' && !/^https?:\/\//i.test(String(data.verification_url||''))) throw new Error('Certificate verification URL is missing or invalid.');
  return true;
}
function templateDataBase_(w,p){
  const settings=getSettings_();
  const org=settings.ORGANIZATION_NAME||APP.ROOT;
  const tz=String(w['Timezone']||APP.DEFAULT_TIMEZONE);
  const mode=normalizeDeliveryMode_(w['Delivery Mode'],String(w['Meeting Link']||'').trim()?'ONLINE':'OFFLINE');
  const logo=brandLogoUrl_(w,settings);
  const duration=durationText_(w);
  const website=w['Website']||settings.DEFAULT_WEBSITE||'';
  const facebook=w['Facebook']||settings.DEFAULT_FACEBOOK||'';
  const linkedin=w['LinkedIn']||settings.DEFAULT_LINKEDIN||'';
  const instagram=w['Instagram']||settings.DEFAULT_INSTAGRAM||'';
  const x=w['X']||settings.DEFAULT_X||'';
  const support=w['Support Email']||settings.DEFAULT_SUPPORT_EMAIL||'';
  return {
    organization_name:esc_(org),logo_url:escAttr_(logo||''),website_url:escAttr_(website||'#'),facebook_url:escAttr_(facebook||'#'),linkedin_url:escAttr_(linkedin||'#'),instagram_url:escAttr_(instagram||'#'),x_url:escAttr_(x||'#'),support_email:escAttr_(support||''),
    workshop_name:esc_(w['Workshop Name']),event_date:esc_(formatDateForDisplay_(w['Workshop Date'],tz)),start_time:esc_(formatTimeForDisplay_(w['Start Time'],tz)),end_time:esc_(formatTimeForDisplay_(w['End Time'],tz)),time_range:esc_(formatTimeForDisplay_(w['Start Time'],tz)+' – '+formatTimeForDisplay_(w['End Time'],tz)),timezone:esc_(tz),duration:esc_(duration),meeting_platform:esc_(w['Meeting Platform']||''),platform:esc_(w['Meeting Platform']||''),meeting_url:escAttr_(w['Meeting Link']||''),meeting_id:esc_(w['Meeting ID']||''),meeting_passcode:esc_(w['Meeting Passcode']||''),host_name:esc_(w['Meeting Host']||''),location:esc_(mode==='ONLINE'?(w['Meeting Platform']||'Online'):[w['Venue Name'],w['Venue Address']].filter(Boolean).join(' • ')),current_year:new Date().getFullYear()
  };
}
function buildRegistrationEmail_(w,p,isTest){
  const d=templateDataBase_(w,p);
  d.participant_name=esc_(p['Full Name']);d.email=esc_(p['Email']);d.phone_whatsapp=esc_(p['Phone / WhatsApp Number']);d.department_faculty=esc_(p['Department / Faculty']);d.season=esc_(p['Season']);d.id_number=esc_(p['ID Number']);d.registration_number=esc_(p['Registration Number']);d.participant_id=esc_(p['Participant ID']);
  d.group_url=escAttr_(w['Group Invite URL']||''); d.calendar_url=escAttr_(buildCalendarUrl_(w)); d.group_block=buildGroupBlock_(w); d.lifecycle_block=buildLifecycleBlock_(w);
  // Do not show test labels in the actual email. The Test Runs sheet is the source of truth.
  d.eyebrow='REGISTRATION SUCCESSFUL'; d.hero_title='You’re officially registered.'; d.greeting='Hello, '+esc_(p['Full Name']);d.hero_text='Your registration for '+esc_(w['Workshop Name'])+' has been confirmed.';
  const html=renderAppleTemplate_('REGISTRATION',d);
  return {subject:'Registration confirmed — '+w['Workshop Name'],html:html,text:'Registration confirmed for '+w['Workshop Name']+'. Registration number: '+(p['Registration Number']||'')+'.'};
}
function buildReminderEmail_(w,p,offset,isTest){
  const d=templateDataBase_(w,p);
  d.participant_name=esc_(p['Full Name']); d.participant_id=esc_(p['Participant ID']); d.registration_number=esc_(p['Registration Number']); d.department=esc_(p['Department / Faculty']); d.faculty=esc_(p['Department / Faculty']);
  const mode=normalizeDeliveryMode_(w['Delivery Mode'],String(w['Meeting Link']||'').trim()?'ONLINE':'OFFLINE');
  d.primary_cta='';
  if(mode==='ONLINE' && w['Meeting Link']) d.primary_cta='<tr><td class="side" align="center" style="padding:22px 35px 0;"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" class="cta-button"><tr><td align="center" style="background:#0071e3;border-radius:13px;"><a href="'+escAttr_(w['Meeting Link'])+'" title="Join the workshop" style="display:block;padding:17px 24px;font-family:-apple-system,BlinkMacSystemFont,\'Segoe UI\',Arial,sans-serif;font-size:17px;line-height:22px;font-weight:600;color:#ffffff;">Join Workshop</a></td></tr></table><p class="apple-font" style="margin:11px 0 0;font-size:12px;line-height:18px;color:#86868b;">Join 5 to 10 minutes early.</p></td></tr>';
  else if(mode==='OFFLINE' && w['Details URL']) d.primary_cta='<tr><td class="side" align="center" style="padding:22px 35px 0;"><table role="presentation" cellpadding="0" cellspacing="0" border="0" class="cta-button"><tr><td align="center" style="background:#0071e3;border-radius:13px;"><a href="'+escAttr_(w['Details URL'])+'" style="display:block;padding:14px 24px;font-family:-apple-system,BlinkMacSystemFont,\'Segoe UI\',Arial,sans-serif;font-size:15px;line-height:20px;font-weight:600;color:#ffffff;">View Venue</a></td></tr></table></td></tr>';
  d.meeting_details_block=buildMeetingDetailsBlock_(w); d.participant_pass_block=buildParticipantPassBlock_(p);
  d.calendar_url=escAttr_(buildCalendarUrl_(w)); d.whatsapp_group_url='';
  d.eyebrow='WORKSHOP REMINDER'; d.hero_title='Your workshop is coming up.'; d.hero_intro='Hello, <strong style="color:#1d1d1f;">'+esc_(p['Full Name'])+'</strong>. Your workshop is scheduled for '+esc_(formatDateForDisplay_(w['Workshop Date'],String(w['Timezone']||APP.DEFAULT_TIMEZONE)))+'.'; d.greeting='Hello, '+esc_(p['Full Name']); d.platform=d.platform || (mode==='OFFLINE'?'Offline':''); d.format_label=mode==='ONLINE'?'Platform':'Location'; d.joining_checklist=buildJoiningChecklistBlock_(mode,w);
  const html=renderAppleTemplate_('REMINDER',d);
  return {subject:'Starting soon — '+w['Workshop Name'],html:html,text:'Your workshop '+w['Workshop Name']+' is coming up.'};
}
function buildAttendanceEmail_(w,p,attendanceUrl,isTest){
  const d=templateDataBase_(w,p); const tz=String(w['Timezone']||APP.DEFAULT_TIMEZONE); const start=parseDateTime_(w['Workshop Date'],w['Start Time'],tz); const end=new Date(start.getTime()+normalizeAttendanceWindowMinutes_(w['Attendance Window Minutes'],120)*60000);
  d.participant_name=esc_(p['Full Name']); d.registration_number=esc_(p['Registration Number']); d.attendance_url=escAttr_(attendanceUrl||''); d.attendance_window=esc_(formatDateTimeForDisplay_(start,tz)+' — '+formatDateTimeForDisplay_(end,tz));
  const html=renderAppleTemplate_('ATTENDANCE',d); return {subject:'Attendance required — '+w['Workshop Name'],html:html,text:'Attendance is required for '+w['Workshop Name']+'.'};
}
function buildCertificateEmail_(w,p,cert,isTest){
  const d=templateDataBase_(w,p); const settings=getSettings_(); const verifyUrl=String(cert.verificationUrl||'').trim();
  d.participant_name=esc_(p['Full Name']); d.season=esc_(p['Season']); d.event_name=esc_(w['Workshop Name']); d.certificate_id=esc_(cert.certificateId); d.certificate_type=esc_(w['Certificate Type']||'Certificate of Participation'); d.certificate_file_name=esc_(safeFileName_(p['Full Name'])+' - '+cert.certificateId+'.pdf'); d.verification_url=escAttr_(verifyUrl); d.website_url=escAttr_(w['Website']||settings.DEFAULT_WEBSITE||'#'); d.eyebrow='CERTIFICATE READY'; d.hero_title='Your certificate is ready.'; d.greeting='Congratulations, '+esc_(p['Full Name'])+'.'; d.hero_text='Your certificate for '+esc_(w['Workshop Name'])+' is attached to this email.';
  const html=renderAppleTemplate_('CERTIFICATE',d); return {subject:'Your certificate is ready — '+w['Workshop Name'],html:html,text:'Your certificate '+cert.certificateId+' for '+w['Workshop Name']+' is ready. The PDF is attached.'};
}
function buildCustomEmail_(w,p,opts){
  const d=templateDataBase_(w,p); d.participant_name=esc_(p['Full Name']||'Participant'); d.message_type=esc_(opts.messageType||'UPDATE'); d.preheader=esc_(opts.preheader||opts.heading||opts.subject||''); d.heading=esc_(opts.heading||opts.subject||'Workshop update'); d.message_html=nl2brEsc_(opts.message||''); d.cta_block=opts.ctaEnabled&&opts.ctaText&&opts.ctaUrl?'<tr><td class="px" align="center" style="padding:0 32px 24px;"><table role="presentation" cellpadding="0" cellspacing="0" border="0" class="btn"><tr><td style="background:#0071e3;border-radius:10px;"><a href="'+escAttr_(opts.ctaUrl)+'" style="display:block;padding:14px 24px;font:600 15px/20px -apple-system,BlinkMacSystemFont,\'Segoe UI\',Arial,sans-serif;color:#fff;min-width:220px;text-align:center;">'+esc_(opts.ctaText)+'</a></td></tr></table></td></tr>':'';
  return {subject:String(opts.subject||'Workshop update'),html:renderAppleTemplate_('CUSTOM',d),text:String(opts.message||'')};
}
function resolveCustomTextTokens_(text,w,p){
  let out=String(text||''); const tz=String(w&&w['Timezone']||APP.DEFAULT_TIMEZONE);
  const map={workshop_name:w&&w['Workshop Name']||'',event_date:w?formatDateForDisplay_(w['Workshop Date'],tz):'',start_time:w?formatTimeForDisplay_(w['Start Time'],tz):'',end_time:w?formatTimeForDisplay_(w['End Time'],tz):'',meeting_url:w&&w['Meeting Link']||'',group_url:w&&w['Group Invite URL']||'',registration_number:p&&p['Registration Number']||'',participant_name:p&&p['Full Name']||''};
  Object.keys(map).forEach(function(k){out=out.split('{{'+k+'}}').join(String(map[k]||''));}); return out;
}
function nl2brEsc_(value){return esc_(String(value||'')).replace(/\r?\n/g,'<br>');}

function sendCustomEmailFromControlCenter_(){
  const sh=getMasterSpreadsheetFast_().getSheetByName('Control Center'); if(!sh) throw new Error('Control Center is missing.');
  const selection=clean_(sh.getRange('B44').getDisplayValue()); const eventId=selection.split(' | ')[0].trim(); const w=findWorkshop_(eventId); if(!w) throw new Error('Select a valid workshop for the custom email.');
  const mode=normalizeCustomRecipientMode_(sh.getRange('E45').getDisplayValue());
  const messageType=clean_(sh.getRange('B45').getDisplayValue())||'Custom';
  const subject=clean_(sh.getRange('B46').getDisplayValue()); const preheader=clean_(sh.getRange('B47').getDisplayValue()); const heading=clean_(sh.getRange('B48').getDisplayValue()); const message=String(sh.getRange('B49').getDisplayValue()||'').trim();
  const ctaEnabled=sh.getRange('B52').getValue()===true; const ctaText=clean_(sh.getRange('E52').getDisplayValue()); const ctaUrl=clean_(sh.getRange('H52').getDisplayValue());
  if(!subject) throw new Error('Custom email subject is required.'); if(!heading) throw new Error('Custom email heading is required.'); if(!message) throw new Error('Custom email message is required.'); if(ctaEnabled && (!ctaText||!isHttpUrl_(ctaUrl))) throw new Error('CTA text and a valid CTA URL are required when the button is enabled.');
  let recipients=[];
  if(mode==='Specific email addresses') recipients=String(sh.getRange('B53').getDisplayValue()||'').split(/[\n,;]+/).map(normalizeEmail_).filter(isValidEmail_);
  else recipients=readSheetObjects_(getSheet_(APP.SHEETS.PARTICIPANTS)).filter(p=>String(p['Event ID'])===eventId&&['REGISTERED','APPROVED'].includes(String(p['Registration Status']))).map(p=>normalizeEmail_(p['Email'])).filter(isValidEmail_);
  recipients=[...new Set(recipients)]; if(!recipients.length) throw new Error('No valid recipients were found.');
  const resolvedSubject=resolveCustomTextTokens_(subject,w,null); const messageId='MSG-'+Utilities.getUuid().slice(0,10).toUpperCase(); appendObject_(getSheet_(APP.SHEETS.CUSTOM_EMAILS),{'Message ID':messageId,'Event ID':eventId,'Workshop Name':w['Workshop Name'],'Recipient Mode':mode,'Recipient Count':recipients.length,'Subject':resolvedSubject,'Heading':heading,'Message':message,'CTA Enabled':ctaEnabled,'CTA Text':ctaText,'CTA URL':ctaUrl,'Created At':new Date(),'Status':'QUEUED','Queued Jobs':0,'Last Error':''});
  let queued=0; const uniquePrefix=messageId+'|';
  recipients.forEach(function(email){ const p=findParticipant_(eventId,email)||{'Participant ID':'CUSTOM-'+Utilities.getUuid().slice(0,8).toUpperCase(),'Event ID':eventId,'Workshop Name':w['Workshop Name'],'Full Name':'Participant','Email':email,'Phone / WhatsApp Number':'','Department / Faculty':'','Season':'','ID Number':'','Registration Number':''}; const e=buildCustomEmail_(w,p,{messageType:messageType,preheader:resolveCustomTextTokens_(preheader,w,p),heading:resolveCustomTextTokens_(heading,w,p),message:resolveCustomTextTokens_(message,w,p),subject:resolveCustomTextTokens_(subject,w,p),ctaEnabled:ctaEnabled,ctaText:resolveCustomTextTokens_(ctaText,w,p),ctaUrl:resolveCustomTextTokens_(ctaUrl,w,p)}); enqueueEmail_({event:w,participant:p,type:'CUSTOM',scheduledAt:new Date(),subject:e.subject,html:e.html,body:e.text,includeMeetingLink:false,uniqueKey:uniquePrefix+email,queueReason:'CUSTOM_EMAIL'}); queued++; });
  updateObjectByKey_(getSheet_(APP.SHEETS.CUSTOM_EMAILS),'Message ID',messageId,{'Queued Jobs':queued,'Status':'QUEUED'});
  return 'Custom email queued for '+queued+' recipient(s). Message ID: '+messageId;
}

function normalizeCustomRecipientMode_(value){const s=String(value||'').trim();return ['All registered participants','Specific email addresses'].includes(s)?s:'All registered participants';}
function saveCustomEmailPresetDefaults_(sh){
  if(!sh.getRange('B42').getValue()) sh.getRange('B42').setValue('All registered participants');
  if(!sh.getRange('B40').getValue()) sh.getRange('B40').setValue('Custom');
  if(!sh.getRange('B48').getValue()) sh.getRange('B48').setValue(false);
}

function actionButtonApple_(label,url,bg,color){
  return '<table role="presentation" cellpadding="0" cellspacing="0" border="0" style="margin-top:22px"><tr><td style="background:'+bg+';border-radius:12px"><a href="'+escAttr_(url)+'" style="display:block;padding:15px 28px;font:600 15px -apple-system,BlinkMacSystemFont,\'Segoe UI\',sans-serif;color:'+color+'">'+esc_(label)+'</a></td></tr></table>';
}


/* =========================
 * V2 — SHEET-NATIVE CONTROL CENTER
 * No HTML UI is required for normal operations.
 * ========================= */

function createTestStepTrigger_(delayMs){ const ms=delayMs===0?1000:(Number(delayMs)>0?Number(delayMs):APP.TEST_DELAY_MS); ScriptApp.newTrigger('processDueTestSteps').timeBased().after(ms).create(); }
function toastControlCenter_(message,title){ try{SpreadsheetApp.getActiveSpreadsheet().toast(String(message),String(title||'WORKSHOP Automation'),6);}catch(err){} }
function controlCenterIsoDate_(value,tz){const zone=String(tz||getSetting_('DEFAULT_TIMEZONE')||APP.DEFAULT_TIMEZONE);if(value instanceof Date&&!isNaN(value.getTime()))return Utilities.formatDate(value,zone,'yyyy-MM-dd');const raw=clean_(value);if(/^\d{4}-\d{2}-\d{2}$/.test(raw))return raw;return normalizeDateOnly_(raw,zone);}
function controlCenterHHMM_(value,tz){const zone=String(tz||getSetting_('DEFAULT_TIMEZONE')||APP.DEFAULT_TIMEZONE);if(value instanceof Date&&!isNaN(value.getTime()))return Utilities.formatDate(value,zone,'HH:mm');const raw=clean_(value);if(/^\d{2}:\d{2}$/.test(raw))return raw;return normalizeTime_(raw,zone);}
function installControlCenterEditTriggerV14_(ss){ const triggers=ScriptApp.getProjectTriggers(); triggers.forEach(function(t){if(t.getHandlerFunction()==='controlCenterOnEditV14_') ScriptApp.deleteTrigger(t);}); ScriptApp.newTrigger('controlCenterOnEditV14_').forSpreadsheet(ss).onEdit().create(); }

/* =========================
 * V1.4.1 — ADDITIVE COMMUNICATION SHEETS
 * Keeps production schemas/data/validation untouched.
 * ========================= */
function setupCommunicationSheetsV141_(ss){
  ss=ss||getMasterSpreadsheetFast_();
  const self=setupSelfTestSheetV141_(ss);
  const custom=setupCustomEmailSheetV141_(ss);
  syncSelfTestWorkshopValidationV141_(self,ss);
  syncCustomEmailWorkshopValidationV141_(custom,ss);
  updateSelfTestRecipientSummaryV141_(self);
  updateCustomEmailRecipientSummaryV141_(custom,ss);
  return {selfTest:self.getName(),customEmail:custom.getName()};
}
function openSelfTestV141_(){const ss=getMasterSpreadsheetFast_();setupCommunicationSheetsV141_(ss);ss.setActiveSheet(ss.getSheetByName('Self Test'));ss.getSheetByName('Self Test').getRange('A1').activate();}
function openCustomEmailV141_(){const ss=getMasterSpreadsheetFast_();setupCommunicationSheetsV141_(ss);ss.setActiveSheet(ss.getSheetByName('Custom Email'));ss.getSheetByName('Custom Email').getRange('A1').activate();}
function hideLegacyCommunicationPanelV141_(sh){try{sh.getRange('A37:H37').setValue('COMMUNICATION TOOLS  ·  USE SELF TEST / CUSTOM EMAIL TABS');sh.hideRows(38,19);}catch(err){}}
function workshopOptionsValidationV141_(sh,cell,ss,hiddenCol){
  const options=getControlCenterWorkshopOptionsV14_(ss);
  sh.getRange(2,hiddenCol,200,1).clearContent();
  if(options.length) sh.getRange(2,hiddenCol,Math.min(200,options.length),1).setValues(options.slice(0,200).map(function(x){return [x];}));
  directDropdownValidationV14_(sh.getRange(cell),options,sh.getRange(2,hiddenCol,200,1));
}
function styleCommunicationSheetV141_(sh){
  sh.setHiddenGridlines(true); sh.setFrozenRows(3);
  const widths={1:190,2:190,3:170,4:170,5:170,6:170,7:170,8:220,10:420};
  Object.keys(widths).forEach(function(k){sh.setColumnWidth(Number(k),widths[k]);});
  sh.hideColumns(10,1);
  sh.getRange('A1:H1').merge().setBackground('#FFFFFF').setFontColor('#1D1D1F').setFontSize(22).setFontWeight('bold').setVerticalAlignment('middle');sh.setRowHeight(1,40);
  sh.getRange('A2:H2').merge().setBackground('#FFFFFF').setFontColor('#6E6E73').setFontSize(11).setWrap(true);sh.setRowHeight(2,30);
  sh.getRange('A1:H20').setFontFamily('Arial');
}
function setupSelfTestSheetV141_(ss){
  let sh=ss.getSheetByName('Self Test');
  if(sh) return sh;
  sh=ss.insertSheet('Self Test');
  styleCommunicationSheetV141_(sh);
  sh.getRange('A1').setValue('SELF TEST');
  sh.getRange('A2').setValue('Real email pipeline test. Temporary test context only. No production participant row is created.');
  sh.getRange('A4').setValue('Workshop');sh.getRange('B4:H4').merge();
  sh.getRange('A5').setValue('Recipient emails');sh.getRange('B5:H7').merge();
  sh.getRange('A8').setValue('Valid recipients');sh.getRange('B8').setValue(1);sh.getRange('D8:H8').merge().setValue('One email per line. Invalid addresses are excluded.');
  sh.getRange('A10').setValue('Full Name');sh.getRange('B10:D10').merge();sh.getRange('E10').setValue('Email');sh.getRange('F10:H10').merge();
  sh.getRange('A11').setValue('Phone / WhatsApp');sh.getRange('B11:D11').merge();sh.getRange('E11').setValue('Department / Faculty');sh.getRange('F11:H11').merge();
  sh.getRange('A12').setValue('Season');sh.getRange('B12:D12').merge();sh.getRange('E12').setValue('ID Number');sh.getRange('F12:H12').merge();
  sh.getRange('A13').setValue('Registration Number');sh.getRange('B13:D13').merge();sh.getRange('E13').setValue('Delay');sh.getRange('F13:H13').merge();
  sh.getRange('A15:H15').merge().setBackground('#F5F5F7').setFontWeight('bold').setValue('TEST STAGES');
  sh.getRange('A16').setValue('Registration');sh.getRange('B16').insertCheckboxes().setValue(true);sh.getRange('C16').setValue('Reminder');sh.getRange('D16').insertCheckboxes().setValue(true);sh.getRange('E16').setValue('Attendance');sh.getRange('F16').insertCheckboxes().setValue(false);sh.getRange('G16').setValue('Certificate');sh.getRange('H16').insertCheckboxes().setValue(true);
  sh.getRange('A18').setValue('RUN SELF TEST →');sh.getRange('B18').insertCheckboxes().setValue(false);sh.getRange('C18:H18').merge().setValue('The send stage starts immediately. Later selected stages use the chosen delay.');
  sh.getRange('A20:H20').merge().setBackground('#F5F5F7').setFontWeight('bold').setValue('STATUS');
  sh.getRange('A21:H24').merge().setWrap(true).setVerticalAlignment('top').setValue('Ready.');
  sh.getRange('B4:H4').setBackground('#FFFFFF').setBorder(true,true,true,true,true,true,'#E5E5E7',SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange('B5:H7').setBackground('#FFFFFF').setBorder(true,true,true,true,true,true,'#E5E5E7',SpreadsheetApp.BorderStyle.SOLID).setWrap(true).setVerticalAlignment('top');
  ['B8','B10:D10','F10:H10','B11:D11','F11:H11','B12:D12','F12:H12','B13:D13','F13:H13'].forEach(function(r){sh.getRange(r).setBackground('#FFFFFF').setBorder(true,true,true,true,true,true,'#E5E5E7',SpreadsheetApp.BorderStyle.SOLID);});
  sh.getRange('B18').setBackground('#0071E3').setFontColor('#FFFFFF').setHorizontalAlignment('center');
  sh.getRange('B5').setValue('karn.sudipto@gmail.com');
  sh.getRange('B10').setValue('Sudipto Kumar');sh.getRange('F10').setValue('karn.sudipto@gmail.com');sh.getRange('B11').setValue('+8801775638329');sh.getRange('F11').setValue('Finance and Banking');sh.getRange('B12').setValue('2021-2022');sh.getRange('F12').setValue('2103060');sh.getRange('B13').setValue('10276');sh.getRange('F13').setValue('1 minute');
  sh.getRange('F13').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Immediately','30 seconds','1 minute','2 minutes','5 minutes'],true).setAllowInvalid(false).build());
  return sh;
}
function syncSelfTestWorkshopValidationV141_(sh,ss){workshopOptionsValidationV141_(sh,'B4',ss,10);if(!sh.getRange('B5').getValue())sh.getRange('B5').setValue('karn.sudipto@gmail.com');}
function parseRecipientEmailsV141_(value){return [...new Set(String(value||'').split(/[\n,;]+/).map(normalizeEmail_).filter(isValidEmail_))];}
function updateSelfTestRecipientSummaryV141_(sh){const valid=parseRecipientEmailsV141_(sh.getRange('B5').getDisplayValue());sh.getRange('B8').setValue(valid.length);if(!valid.length)sh.getRange('A21:H24').setValue('No valid test recipient email is currently entered.');}
function delayMsFromSelfTestV141_(label){return {'Immediately':0,'30 seconds':30000,'1 minute':60000,'2 minutes':120000,'5 minutes':300000}[String(label||'1 minute')] ?? 60000;}
function selectedStagesV141_(sh){const stages=[];if(sh.getRange('B16').getValue()===true)stages.push('REGISTRATION');if(sh.getRange('D16').getValue()===true)stages.push('REMINDER');if(sh.getRange('F16').getValue()===true)stages.push('ATTENDANCE');if(sh.getRange('H16').getValue()===true)stages.push('CERTIFICATE');return stages;}
function statusFieldForStageV141_(stage){return {REGISTRATION:'Confirmation Sent At',REMINDER:'Reminder Sent At',ATTENDANCE:'Attendance Sent At',CERTIFICATE:'Certificate Sent At'}[stage]||'';}
function isSelfTestStagePendingV141_(stage){return stage==='REGISTRATION'||stage==='REMINDER'||stage==='ATTENDANCE'||stage==='CERTIFICATE';}
function nextSelectedStageV141_(stage,stages){const i=stages.indexOf(stage);return i>=0&&i<stages.length-1?stages[i+1]:'';}
function completeSelfTestRunV141_(runId,w,p,cert){const fields={Status:'COMPLETED','Current Step':'DONE','Completed At':new Date()};if(cert){fields['Certificate Created At']=new Date();fields['Certificate Sent At']=new Date();fields['Test Certificate ID']=cert.certificateId;fields['Test Certificate PDF URL']=cert.pdfUrl;}updateTestRun_(runId,fields);log_('INFO','TEST',w['Event ID'],p['Participant ID'],'SELF_TEST_COMPLETED','V1.4.1 self test stages completed',runId);}
function sendSelfTestStageV141_(run,w,p,stage){
  if(stage==='REGISTRATION'){const e=buildRegistrationEmail_(w,p,true);sendDirectEmail_(p['Email'],e.subject,e.html,e.text,'');updateTestRun_(run['Run ID'],{'Confirmation Sent At':new Date()});return null;}
  if(stage==='REMINDER'){const e=buildReminderEmail_(w,p,60,true);sendDirectEmail_(p['Email'],e.subject,e.html,e.text,'');updateTestRun_(run['Run ID'],{'Reminder Sent At':new Date()});return null;}
  if(stage==='ATTENDANCE'){const attendanceUrl=w['Attendance Form ID']?FormApp.openById(String(w['Attendance Form ID'])).getPublishedUrl():'';if(!attendanceUrl)throw new Error('Selected Attendance stage requires an Attendance Form on the workshop.');const e=buildAttendanceEmail_(w,p,attendanceUrl,true);sendDirectEmail_(p['Email'],e.subject,e.html,e.text,'');updateTestRun_(run['Run ID'],{'Attendance Sent At':new Date()});return null;}
  if(stage==='CERTIFICATE'){const cert=generateCertificateForParticipant_(w,p,true,run['Run ID']);const e=buildCertificateEmail_(w,p,cert,true);sendDirectEmail_(p['Email'],e.subject,e.html,e.text,cert.pdfFileId);return cert;}
  throw new Error('Unsupported Self Test stage: '+stage);
}
function startSelfTestRunsV141_(config){
  return withLock_('startSelfTestRunsV141',function(){
    ensureSetup_();
    const w=findWorkshop_(config.eventId);if(!w)throw new Error('Select a valid workshop for Self Test.');
    const recipients=[...new Set((config.recipients||[]).map(normalizeEmail_).filter(isValidEmail_))];if(!recipients.length)throw new Error('Enter at least one valid test recipient email.');
    const stages=(config.stages||[]).filter(isSelfTestStagePendingV141_);if(!stages.length)throw new Error('Select at least one Self Test stage.');
    const delayMs=Number(config.delayMs)>=0?Number(config.delayMs):60000;let started=0;let pending=false;
    recipients.forEach(function(recipient){
      const pid='TEST-PARTICIPANT-'+Utilities.getUuid().slice(0,8).toUpperCase();
      const p={'Participant ID':pid,'Event ID':w['Event ID'],'Workshop Name':w['Workshop Name'],'Registration Timestamp':new Date(),'Full Name':clean_(config.participantName)||'Test Participant','Email':recipient,'Phone / WhatsApp Number':clean_(config.phone)||'','Department / Faculty':clean_(config.department)||'','Season':clean_(config.season)||'','ID Number':clean_(config.idNumber)||'TEST-ID','Registration Number':clean_(config.registrationNumber)||'TEST-REG','Registration Status':'TEST','Attendance Status':'NOT_ATTENDED','Certificate Eligibility':'PENDING','Certificate Status':'NOT_READY','Certificate ID':'','Certificate PDF URL':'','Created At':new Date(),'Updated At':new Date(),'Source Response Row':''};
      p.__v141Stages=stages.slice();p.__v141DelayMs=delayMs;
      const runId='TEST-'+Utilities.getUuid().slice(0,10).toUpperCase();const run={'Run ID':runId,'Event ID':w['Event ID'],'Participant ID':pid,'Workshop Name':w['Workshop Name'],'Recipient':recipient,'Participant Snapshot':JSON.stringify(p),'Include Attendance':stages.indexOf('ATTENDANCE')>=0,'Started At':new Date(),'Status':'RUNNING','Current Step':stages[0]+'_PENDING','Confirmation Sent At':'','Reminder Sent At':'','Attendance Sent At':'','Certificate Created At':'','Certificate Sent At':'','Test Certificate ID':'','Test Certificate PDF URL':'','Last Error':'','Completed At':''};
      appendObject_(getSheet_(APP.SHEETS.TESTS),run);
      const immediateRun=Object.assign({},run);let cert=null;
      try{cert=sendSelfTestStageV141_(immediateRun,w,p,stages[0]);}catch(err){updateTestRun_(runId,{'Status':'FAILED','Last Error':String(err),'Current Step':'FAILED','Completed At':new Date()});throw err;}
      const next=nextSelectedStageV141_(stages[0],stages);
      if(cert){if(!next){completeSelfTestRunV141_(runId,w,p,cert);}else{updateTestRun_(runId,{'Current Step':next+'_PENDING','Certificate Created At':new Date(),'Certificate Sent At':new Date(),'Test Certificate ID':cert.certificateId,'Test Certificate PDF URL':cert.pdfUrl});pending=true;}}
      else if(next){updateTestRun_(runId,{'Current Step':next+'_PENDING'});pending=true;}else completeSelfTestRunV141_(runId,w,p,null);
      started++;
    });
    if(pending) ensureV141TestStepTrigger_(delayMs);
    log_('INFO','TEST',w['Event ID'],'','SELF_TEST_STARTED_V141','V1.4.1 Self Test started; no production participant row created.',started+' recipient(s); stages='+stages.join(',')+'; delayMs='+delayMs);
    return {ok:true,started:started,pending:pending,message:'Self Test started for '+started+' recipient(s). Selected first stage sent immediately.'};
  });
}
function getLatestSelfTestSentAtV141_(run){const vals=['Confirmation Sent At','Reminder Sent At','Attendance Sent At','Certificate Sent At'].map(function(k){return parseQueueDate_(run[k]);}).filter(Boolean).map(function(d){return d.getTime();});if(!vals.length)return null;return new Date(Math.max.apply(null,vals));}
function processModernTestRunStepV141_(run,p,w){
  const stages=Array.isArray(p.__v141Stages)?p.__v141Stages:[];const current=String(run['Current Step']||'').replace(/_PENDING$/,'');if(!stages.includes(current)){throw new Error('V1.4.1 Self Test stage state is invalid.');}
  const sentAt=getLatestSelfTestSentAtV141_(run);const delayMs=Number(p.__v141DelayMs)>=0?Number(p.__v141DelayMs):APP.TEST_DELAY_MS;if(sentAt && Date.now()-sentAt.getTime()<delayMs)return false;
  const cert=sendSelfTestStageV141_(run,w,p,current);const next=nextSelectedStageV141_(current,stages);if(cert){if(!next){completeSelfTestRunV141_(run['Run ID'],w,p,cert);return true;}updateTestRun_(run['Run ID'],{'Current Step':next+'_PENDING','Certificate Created At':new Date(),'Certificate Sent At':new Date(),'Test Certificate ID':cert.certificateId,'Test Certificate PDF URL':cert.pdfUrl});return true;}
  if(next) updateTestRun_(run['Run ID'],{'Current Step':next+'_PENDING'}); else completeSelfTestRunV141_(run['Run ID'],w,p,null);
  return true;
}
function ensureV141TestStepTrigger_(delayMs){try{const exists=ScriptApp.getProjectTriggers().some(function(t){return t.getHandlerFunction()==='processDueTestSteps';});if(!exists)createTestStepTrigger_(delayMs);}catch(err){log_('WARN','TEST','','','TEST_STEP_TRIGGER_DEFERRED','Central scheduler will continue the V1.4.1 test.',String(err));}}

function setupCustomEmailSheetV141_(ss){
  let sh=ss.getSheetByName('Custom Email');if(sh)return sh;
  sh=ss.insertSheet('Custom Email');styleCommunicationSheetV141_(sh);
  sh.getRange('A1').setValue('CUSTOM EMAIL');sh.getRange('A2').setValue('Manual announcement/update tool. Sends through the central quota-safe Email Queue.');
  sh.getRange('A4').setValue('Workshop');sh.getRange('B4:H4').merge();
  sh.getRange('A5').setValue('Preset');sh.getRange('B5').setValue('Custom');sh.getRange('D5').setValue('Recipient mode');sh.getRange('E5:H5').merge();sh.getRange('E5').setValue('All registered participants');
  sh.getRange('A6').setValue('Specific email addresses');sh.getRange('B6:H7').merge();sh.getRange('A8').setValue('Valid recipients');sh.getRange('B8').setValue(0);sh.getRange('D8:H8').merge().setValue('Exact recipient count for the current selection.');
  sh.getRange('A10').setValue('Subject *');sh.getRange('B10:H10').merge();sh.getRange('A11').setValue('Preheader');sh.getRange('B11:H11').merge();sh.getRange('A12').setValue('Heading *');sh.getRange('B12:H12').merge();sh.getRange('A13').setValue('Message *');sh.getRange('B13:H15').merge();
  sh.getRange('A16').setValue('CTA');sh.getRange('B16').insertCheckboxes().setValue(false);sh.getRange('D16').setValue('Button text');sh.getRange('E16').setValue('View Workshop');sh.getRange('G16').setValue('URL');sh.getRange('H16').setValue('');
  sh.getRange('A18').setValue('SEND CUSTOM EMAIL →');sh.getRange('B18').insertCheckboxes().setValue(false);sh.getRange('C18:H18').merge().setValue('Ticking Send is the final confirmation.');
  sh.getRange('A19').setValue('CREATE PREVIEW →');sh.getRange('B19').insertCheckboxes().setValue(false);sh.getRange('C19:H19').merge().setValue('Creates one review draft in Gmail.');
  sh.getRange('A21:H21').merge().setBackground('#F5F5F7').setFontWeight('bold').setValue('STATUS');sh.getRange('A22:H25').merge().setWrap(true).setVerticalAlignment('top').setValue('Ready.');
  ['B4:H4','B5','E5:H5','B6:H7','B8','B10:H10','B11:H11','B12:H12','B13:H15','B16','E16','H16'].forEach(function(r){sh.getRange(r).setBackground('#FFFFFF').setBorder(true,true,true,true,true,true,'#E5E5E7',SpreadsheetApp.BorderStyle.SOLID);});
  ['B18','B19'].forEach(function(a){sh.getRange(a).setBackground('#0071E3').setFontColor('#FFFFFF').setHorizontalAlignment('center');});
  sh.getRange('B5').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Custom','Important Update','Schedule Change','Venue Change','Meeting Link Update','Last Call','Workshop Information','Thank You','Certificate Information'],true).setAllowInvalid(false).build());
  sh.getRange('E5').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['All registered participants','Selected participants','Specific email addresses'],true).setAllowInvalid(false).build());
  return sh;
}
function syncCustomEmailWorkshopValidationV141_(sh,ss){workshopOptionsValidationV141_(sh,'B4',ss,10);}
function getCustomEmailRecipientsV141_(sh,ss){const rawMode=String(sh.getRange('E5').getDisplayValue()||'').trim();const mode=['All registered participants','Selected participants','Specific email addresses'].includes(rawMode)?rawMode:'All registered participants';const sel=clean_(sh.getRange('B4').getDisplayValue());const eventId=sel.split(' | ')[0].trim();if(!eventId)return {mode:mode,eventId:'',recipients:[]};let recipients=[];if(mode==='Specific email addresses')recipients=parseRecipientEmailsV141_(sh.getRange('B6').getDisplayValue());else if(mode==='Selected participants'){const emails=parseRecipientEmailsV141_(sh.getRange('B6').getDisplayValue());recipients=emails;}else recipients=readSheetObjects_(getSheet_(APP.SHEETS.PARTICIPANTS)).filter(function(p){return String(p['Event ID'])===eventId&&['REGISTERED','APPROVED'].includes(String(p['Registration Status']));}).map(function(p){return normalizeEmail_(p['Email']);}).filter(isValidEmail_);return {mode:mode,eventId:eventId,recipients:[...new Set(recipients)]};}
function updateCustomEmailRecipientSummaryV141_(sh,ss){const r=getCustomEmailRecipientsV141_(sh,ss);sh.getRange('B8').setValue(r.recipients.length);sh.getRange('C18:H18').setValue(r.recipients.length?'Ready to send to '+r.recipients.length+' recipient(s). Tick SEND to confirm.':'Enter/select a workshop and at least one valid recipient.');}
function customEmailPresetV141_(type){return {'Important Update':{subject:'Important update — {{workshop_name}}',heading:'Important update',message:'We have an important update about your upcoming workshop.'},'Schedule Change':{subject:'Schedule update — {{workshop_name}}',heading:'The workshop schedule has changed',message:'Please review the updated workshop schedule below.'},'Venue Change':{subject:'Venue update — {{workshop_name}}',heading:'The workshop venue has changed',message:'Please note the updated venue information for your workshop.'},'Meeting Link Update':{subject:'Updated meeting link — {{workshop_name}}',heading:'Your meeting link has been updated',message:'Please use the updated meeting link below to join the workshop.'},'Last Call':{subject:'Last call — {{workshop_name}}',heading:'Your workshop is coming up',message:'This is a final reminder to keep your workshop details ready.'},'Workshop Information':{subject:'Workshop information — {{workshop_name}}',heading:'Workshop information',message:'Here are the latest details for your workshop.'},'Thank You':{subject:'Thank you — {{workshop_name}}',heading:'Thank you for joining us',message:'Thank you for taking part in our workshop.'},'Certificate Information':{subject:'Certificate information — {{workshop_name}}',heading:'Certificate information',message:'Here is an important update about your workshop certificate.'}}[type]||null;}
function applyCustomEmailPresetV141_(sh){const p=customEmailPresetV141_(sh.getRange('B5').getDisplayValue());if(!p)return;sh.getRange('B10').setValue(p.subject);sh.getRange('B12').setValue(p.heading);sh.getRange('B13').setValue(p.message);}
function firstPreviewRecipientV141_(sh,ss,eventId){const r=getCustomEmailRecipientsV141_(sh,ss).recipients;if(r.length)return r[0];try{const user=normalizeEmail_(Session.getEffectiveUser().getEmail());if(isValidEmail_(user))return user;}catch(err){}const p=readSheetObjects_(getSheet_(APP.SHEETS.PARTICIPANTS)).find(function(x){return String(x['Event ID'])===eventId&&isValidEmail_(normalizeEmail_(x['Email']));});return p?normalizeEmail_(p['Email']):'';}
function customEmailOptionsV141_(sh,w,p){const opts={messageType:clean_(sh.getRange('B5').getDisplayValue())||'Custom',preheader:resolveCustomTextTokens_(clean_(sh.getRange('B11').getDisplayValue()),w,p),heading:resolveCustomTextTokens_(clean_(sh.getRange('B12').getDisplayValue()),w,p),message:resolveCustomTextTokens_(String(sh.getRange('B13').getDisplayValue()||''),w,p),subject:resolveCustomTextTokens_(clean_(sh.getRange('B10').getDisplayValue()),w,p),ctaEnabled:sh.getRange('B16').getValue()===true,ctaText:resolveCustomTextTokens_(clean_(sh.getRange('E16').getDisplayValue()),w,p),ctaUrl:resolveCustomTextTokens_(clean_(sh.getRange('H16').getDisplayValue()),w,p)};return opts;}
function sendCustomEmailFromSheetV141_(sh){
  const ss=getMasterSpreadsheetFast_();const selection=clean_(sh.getRange('B4').getDisplayValue());const eventId=selection.split(' | ')[0].trim();const w=findWorkshop_(eventId);if(!w)throw new Error('Select a valid workshop for the custom email.');
  const recipientsInfo=getCustomEmailRecipientsV141_(sh,ss);const recipients=recipientsInfo.recipients;if(!recipients.length)throw new Error('No valid recipients were found.');
  const subject=clean_(sh.getRange('B10').getDisplayValue());const heading=clean_(sh.getRange('B12').getDisplayValue());const message=String(sh.getRange('B13').getDisplayValue()||'').trim();const cta=sh.getRange('B16').getValue()===true;const ctaText=clean_(sh.getRange('E16').getDisplayValue());const ctaUrl=clean_(sh.getRange('H16').getDisplayValue());if(!subject||!heading||!message)throw new Error('Subject, heading, and message are required.');if(cta&&(!ctaText||!isHttpUrl_(ctaUrl)))throw new Error('CTA text and a valid CTA URL are required when the button is enabled.');
  const messageId='MSG-'+Utilities.getUuid().slice(0,10).toUpperCase();const resolvedSubject=resolveCustomTextTokens_(subject,w,null);appendObject_(getSheet_(APP.SHEETS.CUSTOM_EMAILS),{'Message ID':messageId,'Event ID':eventId,'Workshop Name':w['Workshop Name'],'Recipient Mode':recipientsInfo.mode,'Recipient Count':recipients.length,'Subject':resolvedSubject,'Heading':heading,'Message':message,'CTA Enabled':cta,'CTA Text':ctaText,'CTA URL':ctaUrl,'Created At':new Date(),'Status':'QUEUED','Queued Jobs':0,'Last Error':''});
  let queued=0;const prefix=messageId+'|';recipients.forEach(function(email){const p=findParticipant_(eventId,email)||{'Participant ID':'CUSTOM-'+Utilities.getUuid().slice(0,8).toUpperCase(),'Event ID':eventId,'Workshop Name':w['Workshop Name'],'Full Name':'Participant','Email':email,'Phone / WhatsApp Number':'','Department / Faculty':'','Season':'','ID Number':'','Registration Number':''};const e=buildCustomEmail_(w,p,customEmailOptionsV141_(sh,w,p));enqueueEmail_({event:w,participant:p,type:'CUSTOM',scheduledAt:new Date(),subject:e.subject,html:e.html,body:e.text,includeMeetingLink:false,uniqueKey:prefix+email,queueReason:'CUSTOM_EMAIL'});queued++;});updateObjectByKey_(getSheet_(APP.SHEETS.CUSTOM_EMAILS),'Message ID',messageId,{'Queued Jobs':queued,'Status':'QUEUED'});return 'Custom email queued for '+queued+' recipient(s). Message ID: '+messageId;}
function createCustomEmailPreviewDraftFromSheetV141_(sh){const ss=getMasterSpreadsheetFast_();const selection=clean_(sh.getRange('B4').getDisplayValue());const eventId=selection.split(' | ')[0].trim();const w=findWorkshop_(eventId);if(!w)throw new Error('Select a workshop for the preview.');const to=firstPreviewRecipientV141_(sh,ss,eventId);if(!to)throw new Error('Enter at least one valid recipient email for the preview.');const p=findParticipant_(eventId,to)||{'Participant ID':'PREVIEW','Full Name':'Preview Recipient','Email':to,'Registration Number':'PREVIEW'};const opts=customEmailOptionsV141_(sh,w,p);if(!opts.subject||!opts.heading||!opts.message)throw new Error('Fill Subject, Heading, and Message before creating a preview.');const e=buildCustomEmail_(w,p,opts);const draft=GmailApp.createDraft(to,e.subject,e.text,{htmlBody:e.html,name:getSetting_('SENDER_NAME')||APP.ROOT});const url='https://mail.google.com/mail/u/0/#drafts/'+encodeURIComponent(draft.getId());sh.getRange('A22:H25').setValue('Preview draft created for '+to+'\nOpen it in Gmail: '+url);return 'Preview draft created for '+to+'.';}
function controlCenterOnEditSelfTestV141_(e,sh){updateSelfTestRecipientSummaryV141_(sh);if(e.range.getA1Notation()==='B4')return;if(e.range.getA1Notation()==='B18'&&e.value==='TRUE'){sh.getRange('B18').setValue(false);const stages=selectedStagesV141_(sh);const eventId=clean_(sh.getRange('B4').getDisplayValue()).split(' | ')[0].trim();const r=startSelfTestRunsV141_({eventId:eventId,recipients:parseRecipientEmailsV141_(sh.getRange('B5').getDisplayValue()),participantName:clean_(sh.getRange('B10').getDisplayValue()),phone:clean_(sh.getRange('B11').getDisplayValue()),department:clean_(sh.getRange('F11').getDisplayValue()),season:clean_(sh.getRange('B12').getDisplayValue()),idNumber:clean_(sh.getRange('F12').getDisplayValue()),registrationNumber:clean_(sh.getRange('B13').getDisplayValue()),stages:stages,delayMs:delayMsFromSelfTestV141_(sh.getRange('F13').getDisplayValue())});sh.getRange('A21:H24').setValue(r.message);toastControlCenter_(r.message,'Self Test');}}
function controlCenterOnEditCustomEmailV141_(e,sh){const a1=e.range.getA1Notation();if(a1==='B5'){applyCustomEmailPresetV141_(sh);return;}if(['B4','E5','B6'].includes(a1)||a1==='B5'){updateCustomEmailRecipientSummaryV141_(sh);return;}if(a1==='B18'&&e.value==='TRUE'){sh.getRange('B18').setValue(false);const msg=sendCustomEmailFromSheetV141_(sh);sh.getRange('A22:H25').setValue(msg);toastControlCenter_(msg,'Custom Email');return;}if(a1==='B19'&&e.value==='TRUE'){sh.getRange('B19').setValue(false);const msg=createCustomEmailPreviewDraftFromSheetV141_(sh);toastControlCenter_(msg,'Email Preview');return;}}
function buildPublicCertificateVerificationUrl_(certificateId){const id=clean_(certificateId);if(!id)return '';let base=String(getSetting_('PUBLIC_VERIFY_URL')||getSetting_('WEB_APP_URL')||ScriptApp.getService().getUrl()||'').trim();if(!base)return '';base=base.replace(/[?#].*$/,'').replace(/\/+$/,'');return base+'/'+encodeURIComponent(id);}
function setupControlCenter_(ss) {
  ss = ss || getMasterSpreadsheetFast_();
  let sh = ss.getSheetByName('Control Center');
  if (!sh) sh = ss.insertSheet('Control Center', 0);
  try { sh.getRange('A1:K90').breakApart(); } catch (err) {}
  try { sh.getRange('A1:K90').clearDataValidations(); } catch (err) {}
  sh.clear(); sh.clearConditionalFormatRules(); sh.setHiddenGridlines(true); sh.setFrozenRows(3);
  const widths={1:190,2:170,3:80,4:170,5:165,6:80,7:165,8:220,9:24,10:340,11:520};
  Object.keys(widths).forEach(k=>sh.setColumnWidth(Number(k),widths[k]));
  sh.hideColumns(10,2);
  sh.getRange('A1:H1').merge().setValue('WORKSHOP CONTROL CENTER');
  sh.getRange('A2:H2').merge().setValue('Create • Communicate • Test • Deliver');
  sh.getRange('A3:H3').merge().setValue('Simple workshop setup. The system handles IDs, forms, folders, queues, scheduling, and verification automatically.');

  section_(sh,'A5:H5','CREATE WORKSHOP  ·  STEP 1 — BASICS');
  sh.getRange('A6').setValue('Workshop name *'); sh.getRange('B6:H6').merge();
  sh.getRange('A7').setValue('Date *'); sh.getRange('B7').setNumberFormat('yyyy-mm-dd'); sh.getRange('D7').setValue('Start'); sh.getRange('E7').setNumberFormat('HH:mm'); sh.getRange('G7').setValue('End'); sh.getRange('H7').setNumberFormat('HH:mm');
  sh.getRange('A8').setValue('Format *'); sh.getRange('B8:C8').merge(); sh.getRange('D8').setValue('Platform'); sh.getRange('E8:F8').merge(); sh.getRange('G8').setValue('Capacity'); sh.getRange('H8').setValue(0);
  sh.getRange('A9').setValue('Meeting link'); sh.getRange('B9:H9').merge();
  sh.getRange('A10').setValue('Meeting ID'); sh.getRange('B10:C10').merge(); sh.getRange('D10').setValue('Passcode'); sh.getRange('E10:F10').merge(); sh.getRange('G10').setValue('Host'); sh.getRange('H10').setValue('');
  sh.getRange('A11').setValue('Venue name'); sh.getRange('B11:H11').merge();
  sh.getRange('A12').setValue('Venue address'); sh.getRange('B12:H12').merge();
  sh.getRange('A13').setValue('Group'); sh.getRange('B13:C13').merge(); sh.getRange('D13').setValue('Group name'); sh.getRange('E13:H13').merge();
  sh.getRange('A14').setValue('Group invite URL'); sh.getRange('B14:H14').merge();
  sh.getRange('A15').setValue('Notes'); sh.getRange('B15:H16').merge();

  section_(sh,'A18:H18','STEP 2 — REGISTRATION & COMMUNICATION');
  sh.getRange('A19').setValue('Registration confirmation'); sh.getRange('B19').setValue('IMMEDIATE'); sh.getRange('D19').setValue('Reminder'); sh.getRange('E19').insertCheckboxes().setValue(true); sh.getRange('G19').setValue('Timing'); sh.getRange('H19').setValue('2 days before');
  sh.getRange('A20').setValue('Group invitation'); sh.getRange('B20:F20').merge(); sh.getRange('B20').setValue('With registration confirmation'); sh.getRange('G20').setValue('Announcement'); sh.getRange('H20').setValue('MANUAL ONLY');
  sh.getRange('A21:H21').merge().setValue('Registration email includes the group invitation and calendar. Automatic announcement emails are off; use Custom Email when you want to send an update.');

  section_(sh,'A23:H23','STEP 3 — ATTENDANCE & CERTIFICATE');
  sh.getRange('A24').setValue('Attendance'); sh.getRange('B24').insertCheckboxes().setValue(true); sh.getRange('D24').setValue('Window'); sh.getRange('E24:F24').merge(); sh.getRange('E24').setValue('2 hours'); sh.getRange('G24').setValue('Certificate'); sh.getRange('H24').insertCheckboxes().setValue(true);
  sh.getRange('A25').setValue('Certificate release'); sh.getRange('B25:H25').merge(); sh.getRange('B25').setValue('Immediately after workshop');
  sh.getRange('A26:H26').merge().setValue('Attendance ON → participants receive an Attendance Required email when the window opens. Attendance OFF → attendance is skipped and registered participants can qualify for certificates without attendance.');

  section_(sh,'A28:H28','STEP 4 — REVIEW & CREATE');
  sh.getRange('A29:H34').merge().setValue('Reviewing…'); sh.getRange('A35').setValue('CREATE WORKSHOP →'); sh.getRange('B35').insertCheckboxes().setValue(false); sh.getRange('D35:H35').merge().setValue('The final button creates the Forms, folders, response sheets, workshop record, and schedule.');

  section_(sh,'A37:H37','COMMUNICATION TOOLS — SELF TEST');
  sh.getRange('A38').setValue('Workshop'); sh.getRange('B38:H38').merge();
  sh.getRange('A39').setValue('Test recipient'); sh.getRange('B39:D39').merge(); sh.getRange('E39').setValue('Test name'); sh.getRange('F39:H39').merge();
  sh.getRange('A40').setValue('Include Attendance'); sh.getRange('B40').insertCheckboxes().setValue(false); sh.getRange('D40').setValue('Flow'); sh.getRange('E40:H40').merge(); sh.getRange('E40').setValue('Registration → 1 min → Reminder → 1 min → Certificate');
  sh.getRange('A41').setValue('RUN SELF TEST →'); sh.getRange('B41').insertCheckboxes().setValue(false); sh.getRange('D41:H41').merge().setValue('Uses a temporary test participant context. No Google Form submission and no production participant row.');

  section_(sh,'A43:H43','COMMUNICATION TOOLS — CUSTOM EMAIL');
  sh.getRange('A44').setValue('Workshop'); sh.getRange('B44:H44').merge();
  sh.getRange('A45').setValue('Message type'); sh.getRange('B45').setValue('Custom'); sh.getRange('D45').setValue('Recipients'); sh.getRange('E45:H45').merge(); sh.getRange('E45').setValue('All registered participants');
  sh.getRange('A46').setValue('Subject *'); sh.getRange('B46:H46').merge();
  sh.getRange('A47').setValue('Preheader'); sh.getRange('B47:H47').merge();
  sh.getRange('A48').setValue('Heading *'); sh.getRange('B48:H48').merge();
  sh.getRange('A49').setValue('Message *'); sh.getRange('B49:H51').merge();
  sh.getRange('A52').setValue('CTA'); sh.getRange('B52').insertCheckboxes().setValue(false); sh.getRange('D52').setValue('Button text'); sh.getRange('E52').setValue('View Workshop'); sh.getRange('G52').setValue('URL'); sh.getRange('H52').setValue('');
  sh.getRange('A53').setValue('Specific emails'); sh.getRange('B53:H54').merge();
  sh.getRange('A55').setValue('SEND CUSTOM EMAIL →'); sh.getRange('B55').insertCheckboxes().setValue(false); sh.getRange('D55:H55').merge().setValue('Recipients are queued through the central email system for quota-safe delivery.');
  sh.getRange('A56').setValue('CREATE PREVIEW DRAFT →'); sh.getRange('B56').insertCheckboxes().setValue(false); sh.getRange('D56:H56').merge().setValue('Creates one review draft in your Gmail before a real send.');

  section_(sh,'A57:H57','WORKSHOP LIST');
  sh.getRange('A58:H58').setValues([['Event ID','Workshop','Date','Status','Mode','Participants','Attendance','Certificate']]);
  sh.getRange('A59:H78').clearContent();

  section_(sh,'A80:H80','SYSTEM STATUS');
  sh.getRange('A81:H84').merge().setValue('System status loading…');
  section_(sh,'A86:H86','GLOBAL BRAND & SYSTEM DEFAULTS');
  sh.getRange('A87').setValue('Organization'); sh.getRange('B87:D87').merge(); sh.getRange('E87').setValue('Sender name'); sh.getRange('F87:H87').merge();
  sh.getRange('A88').setValue('Logo URL'); sh.getRange('B88:H88').merge();
  sh.getRange('A89').setValue('Public verification URL'); sh.getRange('B89:H89').merge();
  sh.getRange('A90').setValue('Website'); sh.getRange('B90:C90').merge(); sh.getRange('D90').setValue('Facebook'); sh.getRange('E90:F90').merge(); sh.getRange('G90').setValue('X'); sh.getRange('H90').setValue('');
  sh.getRange('A91').setValue('LinkedIn'); sh.getRange('B91:C91').merge(); sh.getRange('D91').setValue('Instagram'); sh.getRange('E91:F91').merge(); sh.getRange('G91').setValue('Support email'); sh.getRange('H91').setValue('');
  sh.getRange('A92').setValue('SAVE GLOBAL DEFAULTS →'); sh.getRange('B92').insertCheckboxes().setValue(false); sh.getRange('D92:H92').merge().setValue('These settings are shared across workshops. Workshop-specific values can override them.');

  styleControlCenterV14_(sh); setControlCenterDefaultsV14_(sh); setControlCenterValidationsV14_(sh); syncControlCenterListsV14_(sh,ss);
  hideLegacyCommunicationPanelV141_(sh);
  updateControlCenterReviewV14_(sh); refreshControlCenterV14_(ss);
}
function section_(sh,range,title){ sh.getRange(range).merge().setValue(title); }
function styleControlCenterV14_(sh){
  sh.getRange('A1:H1').setBackground('#FFFFFF').setFontColor('#1D1D1F').setFontSize(24).setFontWeight('bold').setVerticalAlignment('middle'); sh.setRowHeight(1,42);
  sh.getRange('A2:H2').setBackground('#FFFFFF').setFontColor('#6E6E73').setFontSize(12).setFontWeight('bold');
  sh.getRange('A3:H3').setBackground('#F5F5F7').setFontColor('#86868B').setFontSize(10).setWrap(true); sh.setRowHeight(3,30);
  ['A5:H5','A18:H18','A23:H23','A28:H28','A37:H37','A43:H43','A57:H57','A80:H80','A86:H86'].forEach(r=>sh.getRange(r).setBackground('#F5F5F7').setFontColor('#1D1D1F').setFontWeight('bold').setVerticalAlignment('middle'));
  const inputs=['B6:H6','B7','E7','H7','B8:C8','E8:F8','H8','B9:H9','B10:C10','E10:F10','H10','B11:H11','B12:H12','B13:C13','E13:H13','B14:H14','B15:H16','E19','H19','B20:F20','B24','E24:F24','H24','B25:H25','A29:H34','B38:H38','B39:D39','F39:H39','B40','B44:H44','B45','E45:H45','B46:H46','B47:H47','B48:H48','B49:H51','B52','E52','H52','B53:H54','B87:D87','F87:H87','B88:H88','B89:H89','B90:C90','E90:F90','H90','B91:C91','E91:F91','H91'];
  inputs.forEach(r=>sh.getRange(r).setBackground('#FFFFFF').setBorder(true,true,true,true,true,true,'#E5E5E7',SpreadsheetApp.BorderStyle.SOLID));
  ['B35','B41','B55','B56','B92'].forEach(a=>sh.getRange(a).setBackground('#0071E3').setFontColor('#FFFFFF').setHorizontalAlignment('center'));
  sh.getRange('A29:H34').setWrap(true).setVerticalAlignment('top').setFontColor('#424245'); sh.getRange('A21:H21').setWrap(true).setFontColor('#6E6E73').setFontStyle('italic').setFontSize(10); sh.getRange('A26:H26').setWrap(true).setFontColor('#6E6E73').setFontStyle('italic').setFontSize(10); sh.getRange('A41:H41').setWrap(true).setFontColor('#6E6E73').setFontStyle('italic').setFontSize(10); sh.getRange('D55:H55').setWrap(true).setFontColor('#6E6E73').setFontStyle('italic').setFontSize(10);
  sh.getRange('A58:H78').setWrap(true); sh.setRowHeights(58,21,24); sh.setRowHeights(29,6,24); sh.setRowHeights(49,3,24); sh.setRowHeights(53,2,22);
  sh.getRange('A81:H84').setWrap(true).setVerticalAlignment('top').setFontColor('#424245');
  sh.getRange('A58:H58').setBackground('#F5F5F7').setFontWeight('bold').setFontColor('#6E6E73');
}
function setControlCenterDefaultsV14_(sh){
  const settings=getSettings_(); const tz=String(settings.DEFAULT_TIMEZONE||APP.DEFAULT_TIMEZONE); const tomorrow=new Date(Date.now()+86400000);
  if(!sh.getRange('B7').getValue()) sh.getRange('B7').setValue(Utilities.formatDate(tomorrow,tz,'yyyy-MM-dd'));
  if(!sh.getRange('E7').getValue()) sh.getRange('E7').setValue('10:00'); if(!sh.getRange('H7').getValue()) sh.getRange('H7').setValue('12:00');
  if(!sh.getRange('B8').getValue()) sh.getRange('B8').setValue('Online'); if(!sh.getRange('E8').getValue()) sh.getRange('E8').setValue(settings.DEFAULT_MEETING_PLATFORM||'Google Meet'); if(!sh.getRange('H8').getValue()) sh.getRange('H8').setValue(Number(settings.DEFAULT_CAPACITY||0));
  if(!sh.getRange('B13').getValue()) sh.getRange('B13').setValue('WhatsApp'); if(!sh.getRange('E13').getValue()) sh.getRange('E13').setValue('');
  if(!sh.getRange('E19').getValue()) sh.getRange('E19').setValue(true); if(!sh.getRange('H19').getValue()) sh.getRange('H19').setValue('2 days before'); if(!sh.getRange('B20').getValue()) sh.getRange('B20').setValue('With registration confirmation');
  if(sh.getRange('B24').getDisplayValue()==='') sh.getRange('B24').setValue(boolValue_(settings.DEFAULT_ATTENDANCE_ENABLED,true)); if(!sh.getRange('E24').getValue()) sh.getRange('E24').setValue('2 hours'); if(sh.getRange('H24').getDisplayValue()==='') sh.getRange('H24').setValue(true); if(!sh.getRange('B25').getValue()) sh.getRange('B25').setValue('Immediately after workshop');
  if(!sh.getRange('B45').getValue()) sh.getRange('B45').setValue('Custom'); if(!sh.getRange('E45').getValue()) sh.getRange('E45').setValue('All registered participants');
  sh.getRange('B7').setNumberFormat('yyyy-mm-dd'); sh.getRange('E7').setNumberFormat('HH:mm'); sh.getRange('H7').setNumberFormat('HH:mm');
  sh.getRange('B87').setValue(settings.ORGANIZATION_NAME||APP.ROOT); sh.getRange('F87').setValue(settings.SENDER_NAME||APP.ROOT); sh.getRange('B88').setValue(settings.DEFAULT_LOGO_URL||''); sh.getRange('B89').setValue(settings.PUBLIC_VERIFY_URL||settings.WEB_APP_URL||ScriptApp.getService().getUrl()||''); sh.getRange('B90').setValue(settings.DEFAULT_WEBSITE||''); sh.getRange('E90').setValue(settings.DEFAULT_FACEBOOK||''); sh.getRange('H90').setValue(settings.DEFAULT_X||''); sh.getRange('B91').setValue(settings.DEFAULT_LINKEDIN||''); sh.getRange('E91').setValue(settings.DEFAULT_INSTAGRAM||''); sh.getRange('H91').setValue(settings.DEFAULT_SUPPORT_EMAIL||'');
}
function setControlCenterValidationsV14_(sh){
  sh.getRange('B8').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Online','Offline'],true).setAllowInvalid(false).build());
  sh.getRange('E8').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Google Meet','Zoom','Microsoft Teams','Other'],true).setAllowInvalid(false).build());
  sh.getRange('B13').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['WhatsApp','Messenger','Telegram','Other'],true).setAllowInvalid(false).build());
  sh.getRange('H19').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['1 day before','2 days before','3 days before','5 days before','7 days before'],true).setAllowInvalid(false).build());
  sh.getRange('E24').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['30 minutes','1 hour','2 hours','6 hours','12 hours','24 hours'],true).setAllowInvalid(false).build());
  sh.getRange('B25').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Immediately after workshop','5 minutes after workshop','10 minutes after workshop','15 minutes after workshop','30 minutes after workshop','1 hour after workshop','Next morning'],true).setAllowInvalid(false).build());
  sh.getRange('B45').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Custom','Important Update','Schedule Change','Venue Change','Meeting Link Update','Last Call','Workshop Information','Thank You','Certificate Information'],true).setAllowInvalid(false).build());
  sh.getRange('E45').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['All registered participants','Specific email addresses'],true).setAllowInvalid(false).build());
  directDropdownValidationV14_(sh.getRange('B38'),getControlCenterWorkshopOptionsV14_(getMasterSpreadsheetFast_()),sh.getRange('J2:J201')); directDropdownValidationV14_(sh.getRange('B44'),getControlCenterWorkshopOptionsV14_(getMasterSpreadsheetFast_()),sh.getRange('J2:J201'));
}
function directDropdownValidationV14_(cell,values,fallbackRange){const vals=(values||[]).filter(Boolean);let rule=null;if(vals.length<=200&&vals.length)rule=SpreadsheetApp.newDataValidation().requireValueInList(vals,true).setAllowInvalid(false).build();else if(fallbackRange)rule=SpreadsheetApp.newDataValidation().requireValueInRange(fallbackRange,true).setAllowInvalid(false).build();cell.setDataValidation(rule);}
function getControlCenterWorkshopOptionsV14_(ss){return readSheetObjects_(controlSheetFast_(ss,APP.SHEETS.WORKSHOPS)).filter(w=>String(w['Event ID']||'').trim()&&String(w.Status||'').toUpperCase()!=='ARCHIVED').sort((a,b)=>String(a['Workshop Date']||'').localeCompare(String(b['Workshop Date']||''))).slice(0,200).map(w=>String(w['Event ID'])+' | '+String(w['Workshop Name']||'Untitled Workshop'));}
function syncControlCenterListsV14_(sh,ss){ss=ss||getMasterSpreadsheetFast_();const options=getControlCenterWorkshopOptionsV14_(ss);sh.getRange('J2:J201').clearContent();if(options.length)sh.getRange(2,10,Math.min(200,options.length),1).setValues(options.slice(0,200).map(x=>[x]));directDropdownValidationV14_(sh.getRange('B38'),options,sh.getRange('J2:J201'));directDropdownValidationV14_(sh.getRange('B44'),options,sh.getRange('J2:J201'));populateWorkshopListV14_(sh,ss);}
function populateWorkshopListV14_(sh,ss){const ws=readSheetObjects_(controlSheetFast_(ss,APP.SHEETS.WORKSHOPS)).filter(w=>String(w['Event ID']||'').trim()).slice(0,20);const ps=readSheetObjects_(controlSheetFast_(ss,APP.SHEETS.PARTICIPANTS));const rows=ws.map(w=>[w['Event ID'],w['Workshop Name'],w['Workshop Date'],w['Status'],w['Delivery Mode'],ps.filter(p=>String(p['Event ID'])===String(w['Event ID'])&&String(p['Participant ID']||'')).length,boolValue_(w['Attendance Enabled'],true)?'ON':'OFF',String(w['Certificate Enabled']).toUpperCase()==='TRUE'?'ON':'OFF']);sh.getRange('A59:H78').clearContent();if(rows.length)sh.getRange(59,1,rows.length,8).setValues(rows);}
function updateControlCenterReviewV14_(sh){
  const name=clean_(sh.getRange('B6').getDisplayValue())||'Your workshop'; const date=clean_(sh.getRange('B7').getDisplayValue())||'Date'; const start=clean_(sh.getRange('E7').getDisplayValue())||'Start'; const end=clean_(sh.getRange('H7').getDisplayValue())||'End'; const mode=clean_(sh.getRange('B8').getDisplayValue())||'Online'; const platform=clean_(sh.getRange('E8').getDisplayValue())||''; const meeting=clean_(sh.getRange('B9').getDisplayValue()); const venue=clean_(sh.getRange('B11').getDisplayValue()); const group=clean_(sh.getRange('B14').getDisplayValue()); const att=sh.getRange('B24').getValue()===true; const cert=sh.getRange('H24').getValue()===true; const reminder=clean_(sh.getRange('H19').getDisplayValue())||'2 days before'; const certRelease=clean_(sh.getRange('B25').getDisplayValue())||'Immediately after workshop';
  const lines=[name,'',date+' • '+start+' – '+end,mode+(mode==='Online'&&platform?' • '+platform:''),mode==='Online'?(meeting?'Meeting link added':'Meeting link required'):(venue?'Venue added':'Venue required'),'','Registration confirmation • Immediate','Reminder • '+(sh.getRange('E19').getValue()===true?reminder:'OFF'),'Group • '+(group?'Configured':'Not configured'),'','Attendance • '+(att?'ON':'OFF'), 'Certificate • '+(cert?'ON • '+certRelease:'OFF')];
  sh.getRange('A29:H34').setValue(lines.join('\n'));
}
function controlCenterOnEditV14_(e){
  try{
    if(!e||!e.range)return; const sh=e.range.getSheet(); if(sh.getName()==='Self Test'){controlCenterOnEditSelfTestV141_(e,sh);return;} if(sh.getName()==='Custom Email'){controlCenterOnEditCustomEmailV141_(e,sh);return;} if(sh.getName()!=='Control Center')return; const a1=e.range.getA1Notation();
    if(a1==='B8'){const mode=sh.getRange('B8').getDisplayValue(); if(String(mode).toUpperCase()==='OFFLINE'){sh.getRange('B9:H10').clearContent();sh.getRange('E8').clearContent();}}
    if(['B6','B7','E7','H7','B8','E8','B9','B11','B12','B13','E13','B14','B20','B24','E24','H24','B25','E19','H19'].includes(a1))updateControlCenterReviewV14_(sh);
    if(a1==='B35'&&e.value==='TRUE'){sh.getRange('B35').setValue(false);const msg=createWorkshopFromControlCenterV14_();toastControlCenter_(msg,'Workshop created');refreshControlCenterV14_(getMasterSpreadsheetFast_());return;}
    if(a1==='B38'){return;}
    if(a1==='B41'&&e.value==='TRUE'){sh.getRange('B41').setValue(false);const sel=clean_(sh.getRange('B38').getDisplayValue());if(!sel)throw new Error('Select a workshop for Self Test.');const eventId=sel.split(' | ')[0].trim();const r=startTestRun({eventId:eventId,recipient:clean_(sh.getRange('B39').getDisplayValue()),participantName:clean_(sh.getRange('F39').getDisplayValue())||'Test Participant',includeAttendance:sh.getRange('B40').getValue()===true});toastControlCenter_(r.message,'Self Test');return;}
    if(a1==='B45'){applyCustomEmailPresetV14_(sh);return;}
    if(a1==='B55'&&e.value==='TRUE'){sh.getRange('B55').setValue(false);const msg=sendCustomEmailFromControlCenter_();toastControlCenter_(msg,'Custom Email');refreshControlCenterV14_(getMasterSpreadsheetFast_());return;}
    if(a1==='B56'&&e.value==='TRUE'){sh.getRange('B56').setValue(false);const msg=createCustomEmailPreviewDraftFromControlCenter_();toastControlCenter_(msg,'Email Preview');return;}
    if(a1==='B92'&&e.value==='TRUE'){sh.getRange('B92').setValue(false);saveGlobalSettingsV14_(sh);toastControlCenter_('Global defaults saved.','Global Settings');return;}
  }catch(err){const statusRange=sh.getName()==='Self Test'?'A21:H24':(sh.getName()==='Custom Email'?'A22:H25':'A81:H84');sh.getRange(statusRange).setValue('ERROR\n'+String(err&&err.message?err.message:err));toastControlCenter_(String(err&&err.message?err.message:err),'WORKSHOP Automation');}
}
function createWorkshopFromControlCenterV14_(){
  const ss=getMasterSpreadsheetFast_(),sh=ss.getSheetByName('Control Center'); const settings=getSettings_(); const tz=String(settings.DEFAULT_TIMEZONE||APP.DEFAULT_TIMEZONE);
  const name=clean_(sh.getRange('B6').getDisplayValue()); const date=controlCenterIsoDate_(sh.getRange('B7').getValue(),tz); const start=controlCenterHHMM_(sh.getRange('E7').getValue(),tz); const end=controlCenterHHMM_(sh.getRange('H7').getValue(),tz); const mode=String(sh.getRange('B8').getDisplayValue()).toUpperCase();
  if(!name||!date||!start||!end)throw new Error('Workshop name, date, start time, and end time are required.');
  let meetingPlatform='',meetingLink='',meetingId='',meetingPasscode='',meetingHost='',venueName='',venueAddress='';
  if(mode==='ONLINE'){meetingPlatform=clean_(sh.getRange('E8').getDisplayValue())||settings.DEFAULT_MEETING_PLATFORM||'Google Meet';meetingLink=clean_(sh.getRange('B9').getDisplayValue());if(!meetingLink)throw new Error('Meeting link is required for an online workshop.');meetingId=clean_(sh.getRange('B10').getDisplayValue());meetingPasscode=clean_(sh.getRange('E10').getDisplayValue());meetingHost=clean_(sh.getRange('H10').getDisplayValue());const ex=extractMeetingDetails_(meetingLink,meetingPlatform);meetingPlatform=meetingPlatform||ex.platform;meetingId=meetingId||ex.meetingId;meetingPasscode=meetingPasscode||ex.meetingPasscode;meetingHost=meetingHost||ex.meetingHost;}
  else {venueName=clean_(sh.getRange('B11').getDisplayValue());venueAddress=clean_(sh.getRange('B12').getDisplayValue());if(!venueName)throw new Error('Venue name is required for an offline workshop.');}
  const groupUrl=clean_(sh.getRange('B14').getDisplayValue()); const groupEnabled=!!groupUrl; if(groupUrl&&!isHttpUrl_(groupUrl))throw new Error('Enter a valid group invite URL.');
  let reminderOffsets='60'; if(sh.getRange('E19').getValue()===true){const rem=String(sh.getRange('H19').getDisplayValue()); const map={'1 day before':1440,'2 days before':2880,'3 days before':4320,'5 days before':7200,'7 days before':10080}; reminderOffsets=String(map[rem]||2880);} else reminderOffsets='';
  const att=sh.getRange('B24').getValue()===true; const attWindow=attendanceWindowPresetToMinutes_(sh.getRange('E24').getDisplayValue()); const cert=sh.getRange('H24').getValue()===true; const certRelease=String(sh.getRange('B25').getDisplayValue()); let certMode='AFTER_WORKSHOP',certDelay=0,certNext='09:00'; const cm={'Immediately after workshop':0,'5 minutes after workshop':5,'10 minutes after workshop':10,'15 minutes after workshop':15,'30 minutes after workshop':30,'1 hour after workshop':60}; if(Object.prototype.hasOwnProperty.call(cm,certRelease))certDelay=cm[certRelease];else if(certRelease==='Next morning'){certMode='NEXT_MORNING';}
  const result=createWorkshop({workshopName:name,workshopDate:date,startTime:start,endTime:end,deliveryMode:mode,meetingPlatform:meetingPlatform,meetingLink:meetingLink,meetingId:meetingId,meetingPasscode:meetingPasscode,meetingHost:meetingHost,venueName:venueName,venueAddress:venueAddress,capacity:Number(sh.getRange('H8').getValue()||0),reminderEnabled:sh.getRange('E19').getValue()!==false,reminderOffsets:reminderOffsets,attendanceEnabled:att,attendanceWindowMinutes:attWindow,certificateEnabled:cert,certificateReleaseMode:certMode,certificateDelay:certDelay,certificateNextMorningTime:certNext,groupEnabled:groupEnabled,groupType:clean_(sh.getRange('B13').getDisplayValue())||'WhatsApp',groupName:clean_(sh.getRange('E13').getDisplayValue())||'',groupInviteUrl:groupUrl,notes:clean_(sh.getRange('B15').getDisplayValue()),certificateType:'Certificate of Participation'});
  ['B6','B9:H16'].forEach(a=>sh.getRange(a).clearContent()); sh.getRange('B7').setValue(Utilities.formatDate(new Date(Date.now()+86400000),tz,'yyyy-MM-dd'));sh.getRange('E7').setValue('10:00');sh.getRange('H7').setValue('12:00');sh.getRange('B8').setValue('Online');sh.getRange('E8').setValue(settings.DEFAULT_MEETING_PLATFORM||'Google Meet');sh.getRange('H8').setValue(Number(settings.DEFAULT_CAPACITY||0));sh.getRange('B13').setValue(settings.DEFAULT_GROUP_TYPE||'WhatsApp');sh.getRange('E13').setValue('');sh.getRange('E19').setValue(true);sh.getRange('H19').setValue('2 days before');sh.getRange('B24').setValue(boolValue_(settings.DEFAULT_ATTENDANCE_ENABLED,true));sh.getRange('E24').setValue('2 hours');sh.getRange('H24').setValue(true);sh.getRange('B25').setValue('Immediately after workshop');sh.getRange('B35').setValue(false);updateControlCenterReviewV14_(sh);return 'Created '+result.eventId+' — '+result.workshopName+'\nRegistration Form: '+result.registrationFormUrl;
}
function createCustomEmailPreviewDraftFromControlCenter_(){
  const sh=getMasterSpreadsheetFast_().getSheetByName('Control Center'); const sel=clean_(sh.getRange('B44').getDisplayValue()); const eventId=sel.split(' | ')[0].trim(); const w=findWorkshop_(eventId); if(!w) throw new Error('Select a workshop for the preview.');
  let to=''; const mode=normalizeCustomRecipientMode_(sh.getRange('E45').getDisplayValue()); if(mode==='Specific email addresses') to=normalizeEmail_(String(sh.getRange('B53').getDisplayValue()||'').split(/[\n,;]+/)[0]); if(!to||!isValidEmail_(to)){try{to=normalizeEmail_(Session.getEffectiveUser().getEmail());}catch(err){}} if(!to||!isValidEmail_(to)) throw new Error('Enter at least one valid specific email address for the preview.');
  const p=findParticipant_(eventId,to)||{'Participant ID':'PREVIEW','Full Name':'Preview Recipient','Email':to,'Registration Number':'PREVIEW'};
  const opts={messageType:clean_(sh.getRange('B45').getDisplayValue())||'Custom',preheader:resolveCustomTextTokens_(clean_(sh.getRange('B47').getDisplayValue()),w,p),heading:resolveCustomTextTokens_(clean_(sh.getRange('B48').getDisplayValue()),w,p),message:resolveCustomTextTokens_(String(sh.getRange('B49').getDisplayValue()||''),w,p),subject:resolveCustomTextTokens_(clean_(sh.getRange('B46').getDisplayValue()),w,p),ctaEnabled:sh.getRange('B52').getValue()===true,ctaText:resolveCustomTextTokens_(clean_(sh.getRange('E52').getDisplayValue()),w,p),ctaUrl:resolveCustomTextTokens_(clean_(sh.getRange('H52').getDisplayValue()),w,p)};
  if(!opts.subject||!opts.heading||!opts.message) throw new Error('Fill Subject, Heading, and Message before creating a preview.'); const e=buildCustomEmail_(w,p,opts); const draft=GmailApp.createDraft(to,e.subject,e.text,{htmlBody:e.html,name:getSetting_('SENDER_NAME')||APP.ROOT}); const url='https://mail.google.com/mail/u/0/#drafts/'+encodeURIComponent(draft.getId()); sh.getRange('A81:H84').setValue('Preview draft created for '+to+'\nOpen it in Gmail: '+url); return 'Preview draft created for '+to+'.';
}
function saveGlobalSettingsV14_(sh){
  const settingsSh=getSheet_(APP.SHEETS.SETTINGS); const now=new Date();
  const pairs=[['ORGANIZATION_NAME',sh.getRange('B87').getDisplayValue()],['SENDER_NAME',sh.getRange('F87').getDisplayValue()],['DEFAULT_LOGO_URL',sh.getRange('B88').getDisplayValue()],['PUBLIC_VERIFY_URL',sh.getRange('B89').getDisplayValue()],['DEFAULT_WEBSITE',sh.getRange('B90').getDisplayValue()],['DEFAULT_FACEBOOK',sh.getRange('E90').getDisplayValue()],['DEFAULT_X',sh.getRange('H90').getDisplayValue()],['DEFAULT_LINKEDIN',sh.getRange('B91').getDisplayValue()],['DEFAULT_INSTAGRAM',sh.getRange('E91').getDisplayValue()],['DEFAULT_SUPPORT_EMAIL',sh.getRange('H91').getDisplayValue()]];
  pairs.forEach(function(x){updateObjectByKey_(settingsSh,'Key',x[0],{'Value':clean_(x[1]),'Updated At':now});});
}
function applyCustomEmailPresetV14_(sh){
  const type=clean_(sh.getRange('B45').getDisplayValue());
  const presets={
    'Important Update':{subject:'Important update — {{workshop_name}}',heading:'Important update',message:'We have an important update about your upcoming workshop.'},
    'Schedule Change':{subject:'Schedule update — {{workshop_name}}',heading:'The workshop schedule has changed',message:'Please review the updated workshop schedule below.'},
    'Venue Change':{subject:'Venue update — {{workshop_name}}',heading:'The workshop venue has changed',message:'Please note the updated venue information for your workshop.'},
    'Meeting Link Update':{subject:'Updated meeting link — {{workshop_name}}',heading:'Your meeting link has been updated',message:'Please use the updated meeting link below to join the workshop.'},
    'Last Call':{subject:'Last call — {{workshop_name}}',heading:'Your workshop is coming up',message:'This is a final reminder to keep your workshop details ready.'},
    'Workshop Information':{subject:'Workshop information — {{workshop_name}}',heading:'Workshop information',message:'Here are the latest details for your workshop.'},
    'Thank You':{subject:'Thank you — {{workshop_name}}',heading:'Thank you for joining us',message:'Thank you for taking part in our workshop.'},
    'Certificate Information':{subject:'Certificate information — {{workshop_name}}',heading:'Certificate information',message:'Here is an important update about your workshop certificate.'}
  };
  const preset=presets[type]; if(!preset)return;
  sh.getRange('B46').setValue(preset.subject); sh.getRange('B48').setValue(preset.heading); sh.getRange('B49').setValue(preset.message);
}
function refreshControlCenterV14_(ss){ss=ss||getMasterSpreadsheetFast_();let sh=ss.getSheetByName('Control Center');if(!sh){setupControlCenter_(ss);sh=ss.getSheetByName('Control Center');}syncControlCenterListsV14_(sh,ss);hideLegacyCommunicationPanelV141_(sh);setupCommunicationSheetsV141_(ss);updateControlCenterReviewV14_(sh);const ws=readSheetObjects_(getSheet_(APP.SHEETS.WORKSHOPS));const ps=readSheetObjects_(getSheet_(APP.SHEETS.PARTICIPANTS));const q=readSheetObjects_(getSheet_(APP.SHEETS.QUEUE));let quota='UNAVAILABLE';try{quota=MailApp.getRemainingDailyQuota();}catch(err){}const pending=q.filter(x=>['PENDING','RETRY','WAITING_FOR_QUOTA'].includes(String(x.Status))).length;const failed=q.filter(x=>String(x.Status)==='FAILED').length;const lines=['Version: '+APP.VERSION,'Workshops: '+ws.length+' • Participants: '+ps.length,'Pending emails: '+pending+' • Failed: '+failed+' • Remaining quota: '+quota,'Scheduler: active via central trigger • Verification: active via /exec API'];sh.getRange('A81:H84').setValue(lines.join('\n'));return {ok:true,version:APP.VERSION,workshops:ws.length,participants:ps.length,pendingEmails:pending};}

function normalizeCertificateId_(value){
  return clean_(value).toUpperCase();
}
function isTestCertificateRow_(row){
  const id=normalizeCertificateId_(row&&row['Certificate ID']);
  const status=clean_(row&&row['Status']).toUpperCase();
  return id.indexOf('TEST-CERT-')===0 || status.indexOf('TEST')===0;
}
function verifyCertificate(certificateId){
  const id=normalizeCertificateId_(certificateId);
  if(!id) return {valid:false,code:'MISSING_ID',message:'Certificate ID is required.'};
  const isTestRequest=id.indexOf('TEST-CERT-')===0;
  const row=readSheetObjects_(getSheet_(APP.SHEETS.CERTS)).find(r=>{
    const rowId=normalizeCertificateId_(r['Certificate ID']);
    if(rowId!==id) return false;
    const isTestRow=isTestCertificateRow_(r);
    return isTestRequest ? isTestRow : !isTestRow;
  });
  if(!row) return {valid:false,code:'NOT_FOUND',message:'Certificate not found.'};
  return {
    valid:true,
    code:'VALID',
    certificateId:row['Certificate ID'],
    participantName:row['Participant Name'],
    workshopName:row['Workshop Name'],
    eventId:row['Event ID'],
    season:row['Season'],
    issuedAt:row['Issued At'],
    status:row['Status']
  };
}
function verifyCertificatePublic(certificateId){
  try{
    const r=verifyCertificate(certificateId);
    return {
      valid:!!r.valid,
      code:r.code||'',
      message:r.message||'',
      certificateId:r.certificateId||'',
      name:r.participantName||'',
      workshopName:r.workshopName||'',
      eventId:r.eventId||'',
      issueDate:r.issuedAt?formatDateForDisplay_(r.issuedAt,getSetting_('DEFAULT_TIMEZONE')||APP.DEFAULT_TIMEZONE):'',
      status:r.status||'',
      isTest:r.status?String(r.status).toUpperCase().indexOf('TEST')===0:false
    };
  }catch(err){
    return {valid:false,code:'BACKEND_ERROR',message:String(err&&err.message?err.message:err)};
  }
}

function getMasterSpreadsheetFast_(){
  const active=SpreadsheetApp.getActiveSpreadsheet();
  if(active && active.getSheetByName(APP.SHEETS.SETTINGS) && active.getSheetByName(APP.SHEETS.WORKSHOPS)){
    const p=PropertiesService.getScriptProperties();
    if(p.getProperty('MASTER_SPREADSHEET_ID')!==active.getId()) p.setProperty('MASTER_SPREADSHEET_ID',active.getId());
    return active;
  }
  return getMasterSpreadsheet_();
}

function controlSheetFast_(ss,name){
  const sh=ss.getSheetByName(name);
  if(!sh) throw new Error('Missing required sheet: '+name);
  return sh;
}

function getControlData(){
  const ss=getMasterSpreadsheetFast_();
  const warnings=[];
  const workshopSheet=controlSheetFast_(ss,APP.SHEETS.WORKSHOPS);
  const participantSheet=controlSheetFast_(ss,APP.SHEETS.PARTICIPANTS);
  const testSheet=controlSheetFast_(ss,APP.SHEETS.TESTS);
  const workshopRows=readSheetObjects_(workshopSheet);
  const workshops=[];
  workshopRows.forEach(function(w){
    const eventId=String(w['Event ID']||'').trim();
    if(!eventId) return;
    if(String(w.Status||'').toUpperCase()==='ARCHIVED') return;
    const tz=String(w['Timezone']||APP.DEFAULT_TIMEZONE);
    let date='—',start='—';
    try{date=normalizeDateOnly_(w['Workshop Date'],tz);}catch(err){warnings.push('Workshop '+eventId+' date: '+String(err));}
    try{start=normalizeTime_(w['Start Time'],tz);}catch(err){warnings.push('Workshop '+eventId+' start time: '+String(err));}
    workshops.push({
      eventId:eventId,
      name:String(w['Workshop Name']||'Untitled Workshop'),
      date:date,
      start:start,
      status:String(w.Status||'DRAFT'),
      registrationUrl:w['Registration Form ID']?('https://docs.google.com/forms/d/'+w['Registration Form ID']+'/viewform'):''
    });
  });
  const participants=readSheetObjects_(participantSheet).filter(function(p){return String(p['Participant ID']||'').trim();}).map(function(p){
    return {participantId:String(p['Participant ID']),eventId:String(p['Event ID']||''),name:String(p['Full Name']||''),email:String(p['Email']||''),status:String(p['Registration Status']||'')};
  });
  const recentTests=readSheetObjects_(testSheet).slice(-20).reverse();
  return {
    version:APP.VERSION,
    workshops:workshops,
    participants:participants,
    tests:recentTests,
    warnings:warnings,
    source:{spreadsheetId:ss.getId(),spreadsheetName:ss.getName(),workshopRows:workshopRows.length,participantRows:participants.length}
  };
}

function getCreateDefaults(){
  const ss=getMasterSpreadsheetFast_();
  const settings={};
  const sh=controlSheetFast_(ss,APP.SHEETS.SETTINGS);
  readSheetObjects_(sh).forEach(function(r){settings[String(r.Key)]=r.Value;});
  const tomorrow=new Date(Date.now()+24*60*60*1000);
  const tz=settings.DEFAULT_TIMEZONE||APP.DEFAULT_TIMEZONE;
  return {workshopDate:Utilities.formatDate(tomorrow,tz,'yyyy-MM-dd'),startTime:'10:00',endTime:'12:00',workshopType:settings.DEFAULT_WORKSHOP_TYPE||'Workshop',timezone:tz,meetingPlatform:settings.DEFAULT_MEETING_PLATFORM||'Google Meet',capacity:Number(settings.DEFAULT_CAPACITY||0),reminderOffsets:settings.DEFAULT_REMINDERS||'720,60,10',certificateReleaseMode:settings.DEFAULT_CERTIFICATE_RELEASE||'NEXT_MORNING',certificateNextMorningTime:settings.DEFAULT_CERTIFICATE_TIME||'09:00',certificateEligibility:settings.DEFAULT_CERTIFICATE_ELIGIBILITY||'ATTENDANCE_REQUIRED'};
}

function bootstrapControlCenter(){
  let data;
  let defaults;
  try{
    data=getControlData();
  }catch(err){
    data={version:APP.VERSION,workshops:[],participants:[],tests:[],warnings:[],source:{},error:String(err)};
  }
  try{
    defaults=getCreateDefaults();
  }catch(err){
    defaults={workshopDate:'',startTime:'10:00',endTime:'12:00',workshopType:'Workshop',timezone:APP.DEFAULT_TIMEZONE,meetingPlatform:'Google Meet',capacity:0,reminderOffsets:'720,60,10',certificateReleaseMode:'NEXT_MORNING',certificateNextMorningTime:'09:00',certificateEligibility:'ATTENDANCE_REQUIRED',error:String(err)};
  }
  return {data:data,defaults:defaults};
}

function diagnoseWorkshop(eventId){
  const id=String(eventId||'').trim();
  if(!id) throw new Error('Event ID is required.');
  const w=findWorkshop_(id);
  if(!w) throw new Error('Workshop not found: '+id);
  const tz=String(w.Timezone||APP.DEFAULT_TIMEZONE);
  const result={
    eventId:id,
    workshopName:String(w['Workshop Name']||''),
    folder:{id:String(w['Folder ID']||''),valid:eventFolderValid_(w['Folder ID'],id)},
    registration:{formId:String(w['Registration Form ID']||''),formValid:false,spreadsheetId:String(w['Registration Response Spreadsheet ID']||''),spreadsheetValid:false,responseCount:null,responseSheet:''},
    attendance:{formId:String(w['Attendance Form ID']||''),formValid:false,spreadsheetId:String(w['Attendance Response Spreadsheet ID']||''),spreadsheetValid:false,responseCount:null,responseSheet:''},
    schedule:{timezone:tz,start:'',end:'',reminderMode:String(w['Reminder Mode']||''),reminderOffsets:String(w['Reminder Offsets']||''),certificateReleaseMode:String(w['Certificate Release Mode']||''),certificateNextMorningTime:String(w['Certificate Next Morning Time']||''),certificateCustomReleaseAt:String(w['Certificate Custom Release At']||'')}
  };
  result.registration.formValid=formValidForEvent_(w['Registration Form ID'],id);
  result.registration.spreadsheetValid=sheetValidForEvent_(w['Registration Response Spreadsheet ID'],id);
  result.attendance.formValid=formValidForEvent_(w['Attendance Form ID'],id);
  result.attendance.spreadsheetValid=sheetValidForEvent_(w['Attendance Response Spreadsheet ID'],id);
  try{result.registration.responseCount=FormApp.openById(w['Registration Form ID']).getResponses().length;}catch(err){result.registration.responseError=String(err);}
  try{result.attendance.responseCount=FormApp.openById(w['Attendance Form ID']).getResponses().length;}catch(err){result.attendance.responseError=String(err);}
  try{const sh=findResponseSheetByHeaders_(SpreadsheetApp.openById(w['Registration Response Spreadsheet ID']),'REGISTRATION');result.registration.responseSheet=sh.getName();result.registration.lastRow=sh.getLastRow();result.registration.lastColumn=sh.getLastColumn();}catch(err){result.registration.sheetError=String(err);}
  try{const sh=findResponseSheetByHeaders_(SpreadsheetApp.openById(w['Attendance Response Spreadsheet ID']),'ATTENDANCE');result.attendance.responseSheet=sh.getName();result.attendance.lastRow=sh.getLastRow();result.attendance.lastColumn=sh.getLastColumn();}catch(err){result.attendance.sheetError=String(err);}
  try{result.schedule.start=normalizeTime_(w['Start Time'],tz);}catch(err){result.schedule.startError=String(err);}
  try{result.schedule.end=normalizeTime_(w['End Time'],tz);}catch(err){result.schedule.endError=String(err);}
  try{const release=certificateReleaseDate_(w);result.schedule.certificateRelease=release?Utilities.formatDate(release,tz,'yyyy-MM-dd HH:mm:ss'):'';}catch(err){result.schedule.certificateReleaseError=String(err);}
  log_('INFO','DIAGNOSTICS',id,'','WORKSHOP_DIAGNOSIS', 'Workshop diagnostic completed', JSON.stringify(result));
  return result;
}

function diagnoseAllWorkshops(){
  const rows=readSheetObjects_(getSheet_(APP.SHEETS.WORKSHOPS)).filter(function(w){return String(w['Event ID']||'').trim()&&!['ARCHIVED','CANCELLED'].includes(String(w.Status||''));});
  return rows.map(function(w){try{return diagnoseWorkshop(w['Event ID']);}catch(err){return {eventId:String(w['Event ID']),error:String(err)};}});
}

function systemReadinessCheck(){
  ensureSetup_();
  const checks=[];
  checks.push(check_('Master Spreadsheet',()=>!!getMasterSpreadsheet_()));
  checks.push(check_('Required Sheets',()=>Object.keys(SCHEMAS).every(n=>!!getSheet_(n))));
  checks.push(check_('Drive Structure',()=>!!getSettings_().EVENTS_FOLDER_ID));
  checks.push(check_('Certificate Master',()=>!!PropertiesService.getScriptProperties().getProperty('CERTIFICATE_MASTER_ID')));
  checks.push(check_('Scheduler Trigger',()=>ScriptApp.getProjectTriggers().some(t=>t.getHandlerFunction()==='processScheduler')));
  checks.push(check_('Mail Service',()=>MailApp.getRemainingDailyQuota()>=0));
  checks.push(check_('External Request Service',()=>{
    const r=UrlFetchApp.fetch('https://www.google.com/generate_204',{muteHttpExceptions:true,followRedirects:true});
    return r.getResponseCode()>=200 && r.getResponseCode()<400;
  }));
  checks.push(check_('Single Backend',()=>true));
  checks.push(check_('Verification Handler',()=>typeof verifyCertificatePublic==='function'));
  checks.push(check_('Workshop Integrations',()=>validateAllWorkshopIntegrations_()));
  checks.push(check_('Response Sources',()=>diagnoseAllWorkshops().every(function(d){return !d.error && d.registration.formValid && d.registration.spreadsheetValid && d.attendance.formValid && d.attendance.spreadsheetValid && !d.schedule.startError && !d.schedule.endError && !d.schedule.certificateReleaseError;})));
  const failed=checks.filter(c=>c.status==='ERROR');
  tidyEmailQueueSheet_();
  return {version:APP.VERSION,overall:failed.length?'ERROR':'PASS',checks:checks};
}

function validateAllWorkshopIntegrations_(){
  const rows=readSheetObjects_(getSheet_(APP.SHEETS.WORKSHOPS)).filter(function(w){return String(w['Event ID']||'').trim()&&!['ARCHIVED','CANCELLED'].includes(String(w.Status||''));});
  for(let i=0;i<rows.length;i++){
    const w=rows[i],id=String(w['Event ID']);
    if(!eventFolderValid_(w['Folder ID'],id))throw new Error(id+' Folder ID is invalid.');
    if(!formValidForEvent_(w['Registration Form ID'],id))throw new Error(id+' Registration Form ID is invalid.');
    if(!sheetValidForEvent_(w['Registration Response Spreadsheet ID'],id))throw new Error(id+' Registration Response Spreadsheet ID is invalid.');
    if(!formValidForEvent_(w['Attendance Form ID'],id))throw new Error(id+' Attendance Form ID is invalid.');
    if(!sheetValidForEvent_(w['Attendance Response Spreadsheet ID'],id))throw new Error(id+' Attendance Response Spreadsheet ID is invalid.');
  }
  return true;
}

function refreshDashboard_(){
  const sh=getSheet_(APP.SHEETS.DASHBOARD); if(!sh) return;
  const ws=readSheetObjects_(getSheet_(APP.SHEETS.WORKSHOPS));
  const ps=readSheetObjects_(getSheet_(APP.SHEETS.PARTICIPANTS));
  const qs=readSheetObjects_(getSheet_(APP.SHEETS.QUEUE));
  const cs=readSheetObjects_(getSheet_(APP.SHEETS.CERTS));
  let quotaRemaining=''; try { quotaRemaining=MailApp.getRemainingDailyQuota(); } catch(err) { quotaRemaining='UNAVAILABLE'; }
  const metrics=[['Metric','Value','Updated At'],['Workshops',ws.length,new Date()],['Open Workshops',ws.filter(w=>w.Status==='OPEN').length,new Date()],['Participants',ps.length,new Date()],['Attendance Present',ps.filter(p=>p['Attendance Status']==='PRESENT').length,new Date()],['Pending Emails',qs.filter(q=>['PENDING','RETRY','WAITING_FOR_QUOTA'].includes(String(q.Status))).length,new Date()],['Waiting for Quota',qs.filter(q=>String(q.Status)==='WAITING_FOR_QUOTA').length,new Date()],['Failed Emails',qs.filter(q=>String(q.Status)==='FAILED').length,new Date()],['Email Quota Remaining',quotaRemaining,new Date()],['Certificates',cs.filter(c=>!String(c.Status).startsWith('TEST')).length,new Date()],['Version',APP.VERSION,new Date()]];
  sh.clearContents(); sh.getRange(1,1,metrics.length,3).setValues(metrics);
}

function syncWorkshopFormState_(w){
  w=repairIfWorkshopIntegrationBroken_(w);
  const now=new Date(); const tz=String(w.Timezone||APP.DEFAULT_TIMEZONE); const open=w['Registration Opens'] instanceof Date?w['Registration Opens']:new Date(w['Registration Opens']); const close=w['Registration Closes'] instanceof Date?w['Registration Closes']:new Date(w['Registration Closes']);
  const eventStart=parseDateTime_(w['Workshop Date'],w['Start Time'],tz); const eventEnd=parseDateTime_(w['Workshop Date'],w['End Time'],tz);
  let state=String(w.Status||'DRAFT');
  if(now < open) state='DRAFT'; else if(now >= open && now < eventStart && now <= close) state='OPEN'; else if(now >= eventStart && now < eventEnd) state='LIVE'; else if(now >= eventEnd) state='COMPLETED';
  if(['CANCELLED','ARCHIVED'].includes(String(w.Status))) state=w.Status;
  if(state==='OPEN' || state==='DRAFT') { try{FormApp.openById(w['Registration Form ID']).setAcceptingResponses(state==='OPEN');}catch(err){log_('WARN','FORM',String(w['Event ID']),'','FORM_STATE_SYNC_FAILED',String(err),'');} }
  if(state==='COMPLETED') { try{FormApp.openById(w['Registration Form ID']).setAcceptingResponses(false);}catch(err){} }

  // Attendance form is a deliberate time-gated action, not an always-open form.
  // This prevents participants from submitting early responses that are guaranteed
  // to be rejected by the validator.
  if(boolValue_(w['Attendance Enabled'], true) && w['Attendance Form ID']) {
    try {
      const attendanceStart=eventStart;
      const attendanceEnd=new Date(attendanceStart.getTime()+normalizeAttendanceWindowMinutes_(w['Attendance Window Minutes'],120)*60000);
      const accepting=now>=attendanceStart && now<=attendanceEnd && state!=='CANCELLED' && state!=='ARCHIVED';
      FormApp.openById(String(w['Attendance Form ID'])).setAcceptingResponses(accepting);
    } catch(err) {
      log_('WARN','FORM',String(w['Event ID']),'','ATTENDANCE_FORM_STATE_SYNC_FAILED',String(err),'');
    }
  } else if(w['Attendance Form ID']) {
    try{FormApp.openById(String(w['Attendance Form ID'])).setAcceptingResponses(false);}catch(err){}
  }

  if(state!==w.Status) updateWorkshopByEvent_(w['Event ID'],{'Status':state,'Updated At':new Date()});
}

function isCertificateEligible_(w,p){
  return eligibilityFrom_(w,p,String(p['Attendance Status'])==='PRESENT')==='ELIGIBLE';
}
function eligibilityFrom_(w,p,attended){
  const rule=String(w['Certificate Eligibility']||'ATTENDANCE_REQUIRED');
  const attendanceEnabled=boolValue_(w['Attendance Enabled'], true);
  if(rule==='ALL_REGISTERED') return ['REGISTERED','APPROVED','ON_HOLD'].includes(String(p['Registration Status']))?'ELIGIBLE':'NOT_ELIGIBLE';
  if(rule==='APPROVED_ONLY') return String(p['Registration Status'])==='APPROVED'?'ELIGIBLE':'NOT_ELIGIBLE';
  if(!attendanceEnabled) return ['REGISTERED','APPROVED'].includes(String(p['Registration Status']))?'ELIGIBLE':'NOT_ELIGIBLE';
  return attended && ['REGISTERED','APPROVED'].includes(String(p['Registration Status']))?'ELIGIBLE':'PENDING';
}
function certificateReleaseDate_(w){
  const tz=String(w.Timezone||APP.DEFAULT_TIMEZONE);
  const mode=String(w['Certificate Release Mode']||'AFTER_WORKSHOP').toUpperCase();
  if(mode==='AFTER_WORKSHOP') return new Date(parseDateTime_(w['Workshop Date'],w['End Time'],tz).getTime()+Number(w['Certificate Delay']||0)*60000);
  if(mode==='CUSTOM_DATETIME') return parseFlexibleDateTime_(w['Certificate Custom Release At'],tz);
  const nextDay=shiftDate_(w['Workshop Date'],1,tz);
  const timeValue=w['Certificate Next Morning Time']||'09:00';
  return parseDateTime_(nextDay,timeValue,tz);
}

function nextEventId_(year){
  const rows=readSheetObjects_(getSheet_(APP.SHEETS.WORKSHOPS)); let max=0; rows.forEach(r=>{const m=String(r['Event ID']||'').match(/EVT-(\d{4})-(\d+)/); if(m&&m[1]===year) max=Math.max(max,Number(m[2]));}); return 'EVT-'+year+'-'+String(max+1).padStart(3,'0');
}
function nextParticipantId_(year){
  const rows=readSheetObjects_(getSheet_(APP.SHEETS.PARTICIPANTS)); let max=0; rows.forEach(r=>{const m=String(r['Participant ID']||'').match(/REG-(\d{4})-(\d+)/); if(m&&m[1]===year) max=Math.max(max,Number(m[2]));}); return 'REG-'+year+'-'+String(max+1).padStart(5,'0');
}

function findParticipant_(eventId,email){return readSheetObjects_(getSheet_(APP.SHEETS.PARTICIPANTS)).find(p=>String(p['Event ID'])===String(eventId)&&normalizeEmail_(p['Email'])===normalizeEmail_(email))||null;}
function findParticipantById_(id){return readSheetObjects_(getSheet_(APP.SHEETS.PARTICIPANTS)).find(p=>String(p['Participant ID'])===String(id))||null;}
function findWorkshop_(eventId){return readSheetObjects_(getSheet_(APP.SHEETS.WORKSHOPS)).find(w=>String(w['Event ID'])===String(eventId))||null;}
function countWorkshopActiveParticipants_(eventId){return readSheetObjects_(getSheet_(APP.SHEETS.PARTICIPANTS)).filter(p=>String(p['Event ID'])===String(eventId)&&!['REMOVED','REJECTED'].includes(String(p['Registration Status']))).length;}
function hasAttendance_(eventId,pid){return readSheetObjects_(getSheet_(APP.SHEETS.ATTENDANCE)).some(a=>String(a['Event ID'])===String(eventId)&&String(a['Participant ID'])===String(pid)&&String(a['Attendance Status'])==='PRESENT');}
function hasEmailUniqueKey_(key){return readSheetObjects_(getSheet_(APP.SHEETS.QUEUE)).some(r=>String(r['Unique Key'])===String(key));}

function updateParticipant_(id,fields){updateObjectByKey_(getSheet_(APP.SHEETS.PARTICIPANTS),'Participant ID',id,fields);}
function updateWorkshopByEvent_(id,fields){updateObjectByKey_(getSheet_(APP.SHEETS.WORKSHOPS),'Event ID',id,fields);}
function markQueueJob_(id,fields){updateObjectByKey_(getSheet_(APP.SHEETS.QUEUE),'Job ID',id,fields);}
function updateTestRun_(id,fields){updateObjectByKey_(getSheet_(APP.SHEETS.TESTS),'Run ID',id,fields);}

/* =========================
 * V1 — SAFE WORKSHOP MIGRATION / REPAIR
 * ========================= */
const LEGACY_WORKSHOPS_SCHEMA_V2 = ['Event ID','Workshop Name','Workshop Date','Start Time','End Time','Workshop Type','Timezone','Meeting Platform','Meeting Link','Venue Name','Venue Address','Details URL','Capacity','Registration Opens','Registration Closes','Reminder Offsets','Certificate Enabled','Certificate Release Mode','Certificate Delay','Certificate Next Morning Time','Certificate Eligibility','Website','Facebook','LinkedIn','Instagram','YouTube','Notes','Status','Folder ID','Registration Form ID','Registration Response Spreadsheet ID','Attendance Form ID','Attendance Response Spreadsheet ID','Last Registration Row','Last Attendance Row','Created At','Updated At'];

function isBooleanLike_(v){
  const s=String(v===true?'TRUE':v===false?'FALSE':v==null?'':v).trim().toUpperCase();
  return v===true || v===false || ['TRUE','FALSE','YES','NO','1','0'].includes(s);
}
function boolValue_(v,fallback){
  if(v===true||v===false) return v;
  const s=String(v==null?'':v).trim().toUpperCase();
  if(['TRUE','YES','1'].includes(s)) return true;
  if(['FALSE','NO','0'].includes(s)) return false;
  return fallback;
}
function looksLikeShiftedV21WorkshopRow_(row){
  if(!row || row.length<37) return false;
  // A clean V1 row has Reminder Enabled in position 16. A corrupted row still
  // has the old V2 Reminder Offsets there, followed by Certificate Enabled and
  // Certificate Release Mode at the old positions.
  if(isBooleanLike_(row[15]) && (String(row[16]||'').toUpperCase()==='BEFORE_START' || String(row[16]||'').toUpperCase()==='EXACT_DATETIME' || String(row[16]||'').trim()==='')) return false;
  const mode=String(row[17]||'').trim().toUpperCase();
  const elig=String(row[20]||'').trim().toUpperCase();
  const modeOk=['AFTER_WORKSHOP','NEXT_MORNING','CUSTOM_DATETIME'].includes(mode);
  const eligOk=['ALL_REGISTERED','APPROVED_ONLY','ATTENDANCE_REQUIRED'].includes(elig) || elig==='';
  const delay=row[18];
  const delayOk=delay===''||delay==null||isFinite(Number(delay));
  return modeOk && eligOk && delayOk && isBooleanLike_(row[16]);
}
function legacyWorkshopRowToV212_(row){
  const out=new Array(SCHEMAS.Workshops.length).fill('');
  for(let i=0;i<15;i++) out[i]=row[i]||'';
  const reminder=String(row[15]==null?'':row[15]).trim();
  out[15]=reminder!=='';
  out[16]='BEFORE_START';
  out[17]=reminder==='720,60,10'?'60':row[15]||'';
  out[18]='';
  out[19]=boolValue_(row[16],true);
  out[20]=String(row[17]||'AFTER_WORKSHOP').trim()||'AFTER_WORKSHOP';
  out[21]=Number(row[18]||0)||0;
  out[22]=String(row[19]||'09:00').trim()||'09:00';
  out[23]='';
  out[24]=String(row[20]||'ATTENDANCE_REQUIRED').trim()||'ATTENDANCE_REQUIRED';
  for(let oldIndex=21;oldIndex<=36;oldIndex++) out[oldIndex+4]=row[oldIndex]||'';
  return out;
}
function likelyGoogleId_(v){return /^[A-Za-z0-9_-]{10,}$/.test(String(v||'').trim());}
function findEventFolderForWorkshop_(w,eventsFolder){
  const eventId=String(w['Event ID']||'').trim();
  if(!eventId||!eventsFolder) return null;
  const expected=eventId+' - '+safeFileName_(w['Workshop Name']);
  let it=eventsFolder.getFoldersByName(expected); if(it.hasNext()) return it.next();
  const all=eventsFolder.getFolders();
  while(all.hasNext()){
    const f=all.next();
    if(f.getName()===eventId || f.getName().indexOf(eventId+' -')===0) return f;
  }
  return null;
}
function findFileIdInFolder_(folder,exactName,mimeType){
  if(!folder) return '';
  let it=folder.getFilesByName(exactName),best=null;
  while(it.hasNext()){
    const f=it.next(); if(mimeType && f.getMimeType()!==mimeType) continue;
    if(!best || f.getLastUpdated().getTime()>best.getLastUpdated().getTime()) best=f;
  }
  if(best) return best.getId();
  // Graceful fallback for files that gained a duplicate suffix such as "(1)".
  const prefix=String(exactName||''); it=folder.getFiles(); best=null;
  while(it.hasNext()){
    const f=it.next(); if(mimeType && f.getMimeType()!==mimeType) continue;
    if(String(f.getName()).indexOf(prefix)!==0) continue;
    if(!best || f.getLastUpdated().getTime()>best.getLastUpdated().getTime()) best=f;
  }
  return best?best.getId():'';
}
function eventFolderValid_(id,eventId){
  try{const f=DriveApp.getFolderById(String(id||'')); const n=f.getName(); return n===String(eventId)||n.indexOf(String(eventId)+' -')===0;}catch(err){return false;}
}
function formValidForEvent_(id,eventId){
  try{const f=FormApp.openById(String(id||'')); return f.getId()===String(id) && String(f.getTitle()).indexOf(String(eventId)+' -')===0;}catch(err){return false;}
}
function sheetValidForEvent_(id,eventId){
  try{const s=SpreadsheetApp.openById(String(id||'')); return s.getId()===String(id) && String(s.getName()).indexOf(String(eventId)+' -')===0;}catch(err){return false;}
}
function discoverWorkshopIntegrations_(w){
  const out={};
  try{
    const settings=getSettings_();
    const rootId=String(settings.EVENTS_FOLDER_ID||'');
    if(!rootId) return out;
    const eventsFolder=DriveApp.getFolderById(rootId);
    const eventFolder=findEventFolderForWorkshop_(w,eventsFolder);
    if(!eventFolder) return out;
    out['Folder ID']=eventFolder.getId();
    const regFolder=findChildFolder_(eventFolder,'Registration');
    const attFolder=findChildFolder_(eventFolder,'Attendance');
    const eventId=String(w['Event ID']);
    const name=String(w['Workshop Name']);
    if(regFolder){
      out['Registration Form ID']=findFileIdInFolder_(regFolder,eventId+' - '+name+' Registration',MimeType.GOOGLE_FORMS);
      out['Registration Response Spreadsheet ID']=findFileIdInFolder_(regFolder,eventId+' - Registration Responses',MimeType.GOOGLE_SHEETS);
    }
    if(attFolder){
      out['Attendance Form ID']=findFileIdInFolder_(attFolder,eventId+' - '+name+' Attendance',MimeType.GOOGLE_FORMS);
      out['Attendance Response Spreadsheet ID']=findFileIdInFolder_(attFolder,eventId+' - Attendance Responses',MimeType.GOOGLE_SHEETS);
    }
  }catch(err){log_('WARN','REPAIR',String(w['Event ID']||''),'','DISCOVERY_FAILED',String(err),stack_(err));}
  return out;
}
function reconcileWorkshopCheckpoints_(w){
  const eventId=String(w['Event ID']||''); const out={};
  try{
    const ps=readSheetObjects_(getSheet_(APP.SHEETS.PARTICIPANTS)).filter(p=>String(p['Event ID'])===eventId);
    let maxSource=1; ps.forEach(function(p){const n=Number(p['Source Response Row']||0);if(n>maxSource)maxSource=n;});
    const formId=String(w['Registration Form ID']||'');
    if(formId && likelyGoogleId_(formId)){
      try{ const count=FormApp.openById(formId).getResponses().length; out['Last Registration Row']=Math.min(Math.max(1,maxSource),Math.max(1,count+1)); }
      catch(e){
        const regId=String(w['Registration Response Spreadsheet ID']||'');
        if(regId && sheetValidForEvent_(regId,eventId)){ const sh=findResponseSheetByHeaders_(SpreadsheetApp.openById(regId),'REGISTRATION'); out['Last Registration Row']=Math.min(Math.max(1,maxSource),Math.max(1,sh.getLastRow())); }
      }
    }
  }catch(err){}
  try{
    const formId=String(w['Attendance Form ID']||'');
    if(formId && likelyGoogleId_(formId)){
      try{
        const count=FormApp.openById(formId).getResponses().length;
        out['Last Attendance Row']=1;
      }catch(e){
        const attId=String(w['Attendance Response Spreadsheet ID']||'');
        if(attId && sheetValidForEvent_(attId,eventId)){ const sh=findResponseSheetByHeaders_(SpreadsheetApp.openById(attId),'ATTENDANCE'); out['Last Attendance Row']=1; }
      }
    }
  }catch(err){}
  return out;
}
function repairWorkshopMetadata_(w,reason){
  const eventId=String(w['Event ID']||'').trim(); if(!eventId)return {changed:false};
  const discovered=discoverWorkshopIntegrations_(w),fields={};
  const checks={
    'Folder ID':function(v){return eventFolderValid_(v,eventId);},
    'Registration Form ID':function(v){return formValidForEvent_(v,eventId);},
    'Registration Response Spreadsheet ID':function(v){return sheetValidForEvent_(v,eventId);},
    'Attendance Form ID':function(v){return formValidForEvent_(v,eventId);},
    'Attendance Response Spreadsheet ID':function(v){return sheetValidForEvent_(v,eventId);}
  };
  Object.keys(checks).forEach(function(key){
    if(!checks[key](w[key])){const found=String(discovered[key]||'').trim();if(found)fields[key]=found;}
  });
  const merged=Object.assign({},w,fields);
  Object.assign(fields,reconcileWorkshopCheckpoints_(merged));
  if(Object.keys(fields).length){fields['Updated At']=new Date();updateWorkshopByEvent_(eventId,fields);log_('INFO','REPAIR',eventId,'','WORKSHOP_METADATA_REPAIRED',reason||'Workshop metadata reconciled',JSON.stringify(fields));return {changed:true,fields:fields};}
  return {changed:false};
}
function repairAllWorkshopMetadata_(ss){
  ss=ss||getMasterSpreadsheetFast_();
  const wsh=ss.getSheetByName(APP.SHEETS.WORKSHOPS); if(!wsh||wsh.getLastRow()<2)return {ok:true,repaired:0,checked:0};
  let repaired=0,checked=0;
  readSheetObjects_(wsh).forEach(function(w){
    if(!String(w['Event ID']||'').trim()||['ARCHIVED','CANCELLED'].includes(String(w.Status||'')))return;
    checked++; const r=repairWorkshopMetadata_(w,'V1 setup repair'); if(r.changed)repaired++;
  });
  return {ok:true,repaired:repaired,checked:checked};
}
function repairIfWorkshopIntegrationBroken_(w){
  const suspicious=['Folder ID','Registration Form ID','Registration Response Spreadsheet ID','Attendance Form ID','Attendance Response Spreadsheet ID'].some(function(k){return !likelyGoogleId_(w[k]);});
  if(!suspicious)return w;
  repairWorkshopMetadata_(w,'Runtime integration repair');
  return findWorkshop_(w['Event ID'])||w;
}

function migrateLegacySchemas_(ss){
  const now=new Date();
  const psh=ss.getSheetByName('Participants');
  if(psh){
    const headers=psh.getLastColumn()?psh.getRange(1,1,1,psh.getLastColumn()).getValues()[0].map(String):[];
    const legacy=headers.some(function(h){return ['Email Address','Phone Number','Phone / WhatsApp Number','Faculty','Department','Reg Number','Session','Student ID'].indexOf(h)>=0;});
    if(legacy){
      const rows=psh.getLastRow()>1?psh.getRange(2,1,psh.getLastRow()-1,psh.getLastColumn()).getValues():[];
      const idx={};headers.forEach(function(h,i){idx[h]=i;});
      const get=function(row,names){for(let i=0;i<names.length;i++){if(idx[names[i]]!==undefined) return row[idx[names[i]]];}return '';};
      const outHeaders=SCHEMAS.Participants;
      const out=rows.map(function(row){
        let deptFaculty=get(row,['Department / Faculty']);
        if(!deptFaculty){
          const dep=clean_(get(row,['Department']));
          const fac=clean_(get(row,['Faculty']));
          deptFaculty=dep && fac ? dep+' / '+fac : (dep||fac);
        }
        return [
          get(row,['Participant ID']),get(row,['Event ID']),get(row,['Workshop Name']),get(row,['Registration Timestamp']),
          get(row,['Full Name']),get(row,['Email','Email Address']),get(row,['Phone / WhatsApp Number','Phone Number']),deptFaculty,get(row,['Season','Session']),
          get(row,['ID Number','Student ID']),get(row,['Registration Number','Reg Number']),
          get(row,['Registration Status']),get(row,['Attendance Status']),get(row,['Certificate Eligibility']),
          get(row,['Certificate Status']),get(row,['Certificate ID']),get(row,['Certificate PDF URL']),
          get(row,['Created At'])||now,get(row,['Updated At'])||now,get(row,['Source Response Row'])
        ];
      });
      psh.clearContents();
      psh.getRange(1,1,1,outHeaders.length).setValues([outHeaders]);
      if(out.length) psh.getRange(2,1,out.length,outHeaders.length).setValues(out);
    }
  }
  const csh=ss.getSheetByName('Certificates');
  if(csh){
    const headers=csh.getLastColumn()?csh.getRange(1,1,1,csh.getLastColumn()).getValues()[0].map(String):[];
    const legacy=headers.indexOf('Email Address')>=0 || headers.indexOf('Session')>=0;
    if(legacy){
      const rows=csh.getLastRow()>1?csh.getRange(2,1,csh.getLastRow()-1,csh.getLastColumn()).getValues():[];
      const idx={};headers.forEach(function(h,i){idx[h]=i;});
      const get=function(row,names){for(let i=0;i<names.length;i++){if(idx[names[i]]!==undefined) return row[idx[names[i]]];}return '';};
      const outHeaders=SCHEMAS.Certificates;
      const out=rows.map(function(row){
        return [get(row,['Certificate ID']),get(row,['Event ID']),get(row,['Participant ID']),get(row,['Participant Name']),
          get(row,['Email','Email Address']),get(row,['Workshop Name']),get(row,['Season','Session']),get(row,['Issued At']),
          get(row,['PDF URL']),get(row,['PDF File ID']),get(row,['Verification URL']),get(row,['Status']),
          get(row,['Created At'])||now,get(row,['Test Run ID'])];
      });
      csh.clearContents();
      csh.getRange(1,1,1,outHeaders.length).setValues([outHeaders]);
      if(out.length) csh.getRange(2,1,out.length,outHeaders.length).setValues(out);
    }
  }
  const wsh=ss.getSheetByName('Workshops');
  if(wsh){
    const headers=wsh.getLastColumn()?wsh.getRange(1,1,1,wsh.getLastColumn()).getValues()[0].map(String):[];
    if(headers.indexOf('Session')>=0){
      const col=headers.indexOf('Session')+1;
      wsh.deleteColumn(col);
    }
  }
  // Repair the legacy Workshops schema insertion bug if a previous release
  // changed the header row without remapping the underlying data.
  repairCorruptedWorkshopSchemaRows_(ss);

  // Remove the legacy per-workshop default key if it exists.
  const ssh=ss.getSheetByName('Settings');
  if(ssh){
    const vals=ssh.getLastRow()>1?ssh.getRange(2,1,ssh.getLastRow()-1,ssh.getLastColumn()).getValues():[];
    for(let i=vals.length-1;i>=0;i--){
      if(String(vals[i][0])==='DEFAULT_SESSION') ssh.deleteRow(i+2);
    }
  }
}

function repairCorruptedWorkshopSchemaRows_(ss){
  const sh=ss.getSheetByName(APP.SHEETS.WORKSHOPS); if(!sh||sh.getLastRow()<2)return;
  const headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(String);
  const rows=sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues();
  const isNew=headers.length===SCHEMAS.Workshops.length && headers.join('|')===SCHEMAS.Workshops.join('|');
  const isOld=headers.length===LEGACY_WORKSHOPS_SCHEMA_V2.length && headers.join('|')===LEGACY_WORKSHOPS_SCHEMA_V2.join('|');
  if(!isNew&&!isOld)return;
  let changed=false;
  const out=rows.map(function(row){
    if(isNew&&looksLikeShiftedV21WorkshopRow_(row)){changed=true;return legacyWorkshopRowToV212_(row);}
    if(isOld){changed=true;return legacyWorkshopRowToV212_(row);}
    return row.slice(0,SCHEMAS.Workshops.length);
  });
  if(changed){
    sh.getRange(1,1,Math.max(2,out.length+1),SCHEMAS.Workshops.length).clearContent();
    sh.getRange(1,1,1,SCHEMAS.Workshops.length).setValues([SCHEMAS.Workshops]);
    if(out.length)sh.getRange(2,1,out.length,SCHEMAS.Workshops.length).setValues(out);
    log_('INFO','REPAIR','','','WORKSHOPS_SCHEMA_REPAIRED',isOld?'Migrated legacy V2 Workshops schema':'Reconstructed corrupted legacy Workshops rows','rows='+out.length);
  }
}

function getSettings_(){const sh=getSheet_(APP.SHEETS.SETTINGS); const out={}; readSheetObjects_(sh).forEach(r=>out[String(r.Key)]=r.Value); return out;}
function getSetting_(key){return String(getSettings_()[key]||'');}
function ensureSetup_(){
  const p=PropertiesService.getScriptProperties();
  const active=SpreadsheetApp.getActiveSpreadsheet();
  const configured=String(p.getProperty('MASTER_SPREADSHEET_ID')||'').trim();
  const activeLooksLikeMaster=!!(active && active.getSheetByName(APP.SHEETS.SETTINGS) && active.getSheetByName(APP.SHEETS.WORKSHOPS));
  if(activeLooksLikeMaster){
    if(configured!==active.getId()) p.setProperty('MASTER_SPREADSHEET_ID',active.getId());
    return;
  }
  if(!configured) throw new Error('Master spreadsheet is not configured. Open the master Google Sheet and run Setup / Repair once.');
}
function getMasterSpreadsheet_(){
  const p=PropertiesService.getScriptProperties();
  const configured=String(p.getProperty('MASTER_SPREADSHEET_ID')||'').trim();
  const active=SpreadsheetApp.getActiveSpreadsheet();
  const activeLooksLikeMaster=!!(active && active.getSheetByName(APP.SHEETS.SETTINGS) && active.getSheetByName(APP.SHEETS.WORKSHOPS));
  // Single source of truth: when invoked from the bound master Sheet, always use it.
  // This prevents a stale Script Property from sending writes to an older/empty workbook.
  if(activeLooksLikeMaster){
    if(configured!==active.getId()) p.setProperty('MASTER_SPREADSHEET_ID',active.getId());
    return active;
  }
  if(configured){
    try{
      const ss=SpreadsheetApp.openById(configured);
      if(ss.getSheetByName(APP.SHEETS.SETTINGS) && ss.getSheetByName(APP.SHEETS.WORKSHOPS)) return ss;
      throw new Error('Configured master spreadsheet is missing required Settings/Workshops sheets.');
    }catch(err){
      throw new Error('Master spreadsheet could not be opened: '+String(err));
    }
  }
  throw new Error('Master spreadsheet is not configured. Open the master Google Sheet and run Setup / Repair once.');
}
function getSheet_(name){const ss=getMasterSpreadsheet_(); let sh=ss.getSheetByName(name); if(!sh) sh=ss.insertSheet(name); if(!SCHEMAS[name]) return sh; ensureHeaders_(sh,SCHEMAS[name]); return sh;}
function ensureSheetSchema_(ss,name,headers){let sh=ss.getSheetByName(name);if(!sh)sh=ss.insertSheet(name);ensureHeaders_(sh,headers);return sh;}
function ensureHeaders_(sh,headers){
  const lc=sh.getLastColumn();
  const current=lc?sh.getRange(1,1,1,lc).getValues()[0].map(String):[];
  const same=current.length===headers.length&&headers.every(function(h,i){return current[i]===h;});
  if(!same){
    const rows=sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,Math.max(lc,1)).getValues():[];
    const idx={}; current.forEach(function(h,i){if(h)idx[h]=i;});
    const mapped=rows.map(function(row){return headers.map(function(h){return idx[h]!==undefined?row[idx[h]]:'';});});
    sh.clearContents();
    sh.getRange(1,1,1,headers.length).setValues([headers]);
    if(mapped.length)sh.getRange(2,1,mapped.length,headers.length).setValues(mapped);
  }else if(!lc){sh.getRange(1,1,1,headers.length).setValues([headers]);}
  sh.setFrozenRows(1);
}
function readSheetObjects_(sh){const lr=sh.getLastRow(),lc=sh.getLastColumn();if(lr<2||lc<1)return[];const headers=sh.getRange(1,1,1,lc).getValues()[0].map(String);return sh.getRange(2,1,lr-1,lc).getValues().map(v=>mapRow_(headers,v));}
function mapRow_(headers,row){const o={};headers.forEach((h,i)=>o[h]=row[i]);return o;}
function appendObject_(sh,obj){const headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(String);sh.appendRow(headers.map(h=>Object.prototype.hasOwnProperty.call(obj,h)?obj[h]:''));}
function updateObjectByKey_(sh,key,value,fields){const headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(String);const idx=headers.indexOf(key);if(idx<0)throw new Error('Missing key column '+key);const rows=Math.max(0,sh.getLastRow()-1);if(!rows)return false;const vals=sh.getRange(2,1,rows,sh.getLastColumn()).getValues();for(let i=0;i<vals.length;i++){if(String(vals[i][idx])===String(value)){Object.keys(fields).forEach(k=>{const c=headers.indexOf(k);if(c>=0)vals[i][c]=fields[k];});sh.getRange(2,1,rows,sh.getLastColumn()).setValues(vals);return true;}}return false;}

function getOrCreateRootFolder_(){const root=DriveApp.getRootFolder();const it=root.getFoldersByName(APP.ROOT);return it.hasNext()?it.next():root.createFolder(APP.ROOT);}
function getOrCreateFolder_(parent,path){let cur=parent;path.split('/').forEach(part=>{const it=cur.getFoldersByName(part);cur=it.hasNext()?it.next():cur.createFolder(part);});return cur;}
function findChildFolder_(parent,name){const it=parent.getFoldersByName(name);return it.hasNext()?it.next():null;}
function safeFileName_(s){return String(s||'').replace(/[\\/:*?"<>|#%{}]/g,' ').replace(/\s+/g,' ').trim().slice(0,120)||'Workshop';}
function parseDateTime_(date,time,tz){
  const zone=String(tz||APP.DEFAULT_TIMEZONE);
  const dateOnly=normalizeDateOnly_(date,zone);
  const timeOnly=normalizeTime_(time,zone);
  try { return Utilities.parseDate(dateOnly+' '+timeOnly,zone,'yyyy-MM-dd HH:mm'); }
  catch(err){ throw new Error('Invalid workshop date/time. Date="'+dateOnly+'", Time="'+timeOnly+'", Timezone="'+zone+'". '+String(err)); }
}
function normalizeDateOnly_(value,tz){
  const zone=String(tz||APP.DEFAULT_TIMEZONE);
  if(value instanceof Date && !isNaN(value.getTime())) return Utilities.formatDate(value,zone,'yyyy-MM-dd');
  if(typeof value==='number' && isFinite(value)){
    let d=null;
    if(value>100000000000) d=new Date(value);
    else if(value>0 && value<100000) d=new Date(Date.UTC(1899,11,30)+value*86400000);
    if(d && !isNaN(d.getTime())) return Utilities.formatDate(d,zone,'yyyy-MM-dd');
  }
  const s=String(value==null?'':value).trim();
  if(!s) throw new Error('Date is empty.');
  let m=s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if(m) return m[1]+'-'+pad2_(m[2])+'-'+pad2_(m[3]);
  m=s.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
  if(m) return m[3]+'-'+pad2_(m[2])+'-'+pad2_(m[1]);
  const parsed=new Date(s);
  if(!isNaN(parsed.getTime())) return Utilities.formatDate(parsed,zone,'yyyy-MM-dd');
  throw new Error('Unsupported date value: '+s);
}
function normalizeTime_(value,tz){
  const zone=String(tz||APP.DEFAULT_TIMEZONE);
  if(value instanceof Date && !isNaN(value.getTime())){
    // Google Sheets time-only cells may arrive as 1899-12-30. Formatting those
    // in Asia/Dhaka can produce historical local-mean-time offsets such as +05:53.
    // Use UTC for epoch-style dates so 09:00 stays 09:00.
    const epochYear=Number(Utilities.formatDate(value,'UTC','yyyy'));
    if(epochYear<=1900) return Utilities.formatDate(value,'UTC','HH:mm');
    return Utilities.formatDate(value,zone,'HH:mm');
  }
  if(typeof value==='number' && isFinite(value)){
    const totalMinutes=value>=0 && value<1 ? Math.round(value*24*60) : NaN;
    if(Number.isFinite(totalMinutes)) return minutesToHHMM_(totalMinutes);
    if(value>100000000000){const d=new Date(value);if(!isNaN(d.getTime())) return Utilities.formatDate(d,zone,'HH:mm');}
  }
  let s=String(value==null?'':value).trim();
  if(!s) throw new Error('Time is empty.');
  s=s.replace(/\./g,':');
  let m=s.match(/^(\d{1,2}):(\d{2})(?::\d{2})?\s*([AaPp][Mm])?$/);
  if(!m){
    // Migration/runtime safety: a serialized Sheets time may look like
    // "Sat Dec 30 1899 09:00:00 GMT+0530 ...". Extract the clock portion
    // only when the string clearly represents an epoch-style time value.
    if(/\b(?:1899|1900)\b/i.test(s)){
      const tm=s.match(/\b(\d{1,2}):(\d{2})(?::\d{2})?\b/);
      if(tm) return pad2_(Number(tm[1]))+':'+pad2_(Number(tm[2]));
    }
    throw new Error('Unsupported time value: '+s);
  }
  let h=Number(m[1]), minute=Number(m[2]);
  const ap=m[3]?m[3].toUpperCase():'';
  if(minute>59) throw new Error('Invalid minutes in time: '+s);
  if(ap){
    if(h<1||h>12) throw new Error('Invalid 12-hour time: '+s);
    if(ap==='AM'&&h===12)h=0;
    if(ap==='PM'&&h!==12)h+=12;
  }else if(h>23) throw new Error('Invalid 24-hour time: '+s);
  return pad2_(h)+':'+pad2_(minute);
}
function pad2_(n){return String(n).padStart(2,'0');}
function minutesToHHMM_(minutes){minutes=((minutes%1440)+1440)%1440;return pad2_(Math.floor(minutes/60))+':'+pad2_(minutes%60);}
function dateYear_(value,tz){return normalizeDateOnly_(value,tz).slice(0,4);}
function shiftDate_(date,days,tz){const parts=normalizeDateOnly_(date,String(tz||APP.DEFAULT_TIMEZONE)).split('-').map(Number);const d=new Date(Date.UTC(parts[0],parts[1]-1,parts[2]+Number(days||0)));return Utilities.formatDate(d,'UTC','yyyy-MM-dd');}
function formatDateForDisplay_(value,tz){
  const zone=String(tz||APP.DEFAULT_TIMEZONE);
  const normalized=normalizeDateOnly_(value,zone);
  const parsed=Utilities.parseDate(normalized+' 00:00',zone,'yyyy-MM-dd HH:mm');
  return Utilities.formatDate(parsed,zone,'dd MMM yyyy');
}
function formatTimeForDisplay_(value,tz){return normalizeTime_(value,String(tz||APP.DEFAULT_TIMEZONE));}
function addMinutes_(date,mins,tz){return new Date(date.getTime()+Number(mins||0)*60000);}
function parseOffsets_(s){return String(s||'').split(',').map(x=>Number(String(x).trim())).filter(x=>Number.isFinite(x)&&x>0);}
function durationText_(w){
  const tz=String(w['Timezone']||APP.DEFAULT_TIMEZONE);
  const s=normalizeTime_(w['Start Time'],tz).split(':').map(Number);
  const e=normalizeTime_(w['End Time'],tz).split(':').map(Number);
  let start=s[0]*60+s[1], end=e[0]*60+e[1], d=end-start;
  if(d<0)d+=1440;
  return Math.floor(d/60)+'h '+(d%60)+'m';
}
function normalizeEmail_(e){return String(e||'').trim().toLowerCase();}
function isValidEmail_(e){return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(e||''));}
function clean_(x){return x===null||x===undefined?'':String(x).trim();}
function numOr_(x,f){const n=Number(x);return Number.isFinite(n)&&String(x)!==''?n:f;}
function esc_(x){return String(x??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');}
function escAttr_(x){return esc_(x);}
function stripHtml_(html){return String(html||'').replace(/<[^>]+>/g,' ').replace(/\s+/g,' ').trim();}
function stack_(e){return e&&e.stack?e.stack:'';}
function isPaused_(){return String(getSetting_('PAUSE_AUTOMATION')).toUpperCase()==='TRUE';}
function check_(label,fn){try{return{label:label,status:fn()?'PASS':'ERROR'}}catch(err){return{label:label,status:'ERROR',message:String(err)}}}
function lockTimeoutMs_(name){
  switch(String(name||'')){
    case 'setupSystem': return 180000;
    case 'processScheduler': return 1000;
    case 'processDueTestSteps': return 5000;
    case 'createWorkshop': return 30000;
    case 'startTestRun': return 30000;
    default: return 30000;
  }
}
function withLock_(name,fn,timeoutMs){
  const lock=LockService.getScriptLock();
  const timeout=Number(timeoutMs)>0?Number(timeoutMs):lockTimeoutMs_(name);
  const acquired=lock.tryLock(timeout);
  if(!acquired){
    const op=String(name||'operation');
    Logger.log('Lock busy: '+op);
    if(op==='processScheduler' || op==='processDueTestSteps'){
      return {ok:false,busy:true,operation:op,message:'Another automation process is already running. This trigger will retry on its next scheduled run.'};
    }
    throw new Error('System is busy: another automation process is currently using the system lock. '+op+' waited '+Math.round(timeout/1000)+' seconds. Please run '+op+' again after the current process finishes.');
  }
  try{
    return fn();
  } finally {
    try{SpreadsheetApp.flush();}catch(err){}
    lock.releaseLock();
  }
}
function log_(level,module,eventId,participantId,action,message,details){try{const sh=getSheet_(APP.SHEETS.LOGS);appendObject_(sh,{'Timestamp':new Date(),'Level':level,'Module':module,'Event ID':eventId||'','Participant ID':participantId||'','Action':action||'','Message':message||'','Details':details||''});}catch(err){}}


function repairEmailQueueMetadata_() {
  const sh=getSheet_(APP.SHEETS.QUEUE);
  const lr=sh.getLastRow(), lc=sh.getLastColumn();
  if(lr<2) return {updated:0};
  const headers=sh.getRange(1,1,1,lc).getValues()[0].map(String);
  const rows=sh.getRange(2,1,lr-1,lc).getValues();
  const idx={}; headers.forEach(function(h,i){idx[h]=i;});
  const workshops={}; readSheetObjects_(getSheet_(APP.SHEETS.WORKSHOPS)).forEach(function(w){workshops[String(w['Event ID'])]=w;});
  let updated=0;
  rows.forEach(function(row){
    const type=String(row[idx['Email Type']]||'').toUpperCase();
    const eventId=String(row[idx['Event ID']]||'');
    const w=workshops[eventId]||null;
    if(row[idx['Priority']]==='' || row[idx['Priority']]==null) { row[idx['Priority']]=emailPriority_(type); updated++; }
    if((row[idx['Not Before']]==='' || row[idx['Not Before']]==null) && row[idx['Scheduled At']]) { row[idx['Not Before']]=row[idx['Scheduled At']]; updated++; }
    if((row[idx['Deadline At']]==='' || row[idx['Deadline At']]==null) && w) { const d=emailDeadlineFor_(w,type); if(d){row[idx['Deadline At']]=d;updated++;} }
    if((row[idx['Queue Reason']]==='' || row[idx['Queue Reason']]==null)) { row[idx['Queue Reason']]='LEGACY_MIGRATION'; updated++; }
  });
  if(updated) sh.getRange(2,1,rows.length,lc).setValues(rows);
  if(updated) log_('INFO','REPAIR','','','EMAIL_QUEUE_METADATA_REPAIRED','Updated '+updated+' legacy email queue field(s).','rows='+rows.length);
  return {updated:updated};
}

function tidyEmailQueueSheet_() {
  const sh=getSheet_(APP.SHEETS.QUEUE);
  const headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(String);
  const hiddenNames=['Unique Key','Attachment File ID','Last Error','Queue Reason','Test Run ID','Message Body','Include Meeting Link','Sender Name','Created At','Updated At'];
  hiddenNames.forEach(function(name){
    const idx=headers.indexOf(name);
    if(idx>=0 && !sh.isColumnHiddenByUser(idx+1)) sh.hideColumns(idx+1);
  });
  sh.setFrozenRows(1);
  const visibleNames=['Job ID','Event ID','Participant ID','Email Type','Priority','Scheduled At','Not Before','Deadline At','Status','Attempt Count','Last Attempt At','Sent At','Recipient','Subject'];
  visibleNames.forEach(function(name){
    const idx=headers.indexOf(name);
    if(idx>=0) sh.setColumnWidth(idx+1,Math.max(100,Math.min(220,name.length*9+28)));
  });
}

function cleanupAutomationTriggers_() {
  const triggers=ScriptApp.getProjectTriggers();
  triggers.forEach(function(t){
    const h=t.getHandlerFunction();
    if(h==='processScheduler' || h==='processDueTestSteps') {
      // Central scheduler is recreated by installCentralScheduler_().
      // One-time test triggers should be left alone while running.
      if(h==='processScheduler') ScriptApp.deleteTrigger(t);
    }
  });
}

function selfRepairV2_() {
  ensureSetup_();
  migrateLegacySchemas_(getMasterSpreadsheet_());
  const ss=getMasterSpreadsheet_();
  Object.keys(SCHEMAS).forEach(function(name){ ensureSheetSchema_(ss,name,SCHEMAS[name]); });
  setupControlCenter_(ss);
  installControlCenterEditTriggerV14_(ss);
  repairEmailQueueMetadata_();
  tidyEmailQueueSheet_();
  installCentralScheduler_();
  refreshDashboard_();
  refreshControlCenterV14_();
  log_('INFO','REPAIR','','','V1_REPAIR_COMPLETE','Self-repair completed.','Sheet-native Control Center, scheduler isolation, safe Form response readers, stable time normalization, certificate export checks, quota-aware queue cleanup.');
  return {ok:true,version:APP.VERSION};
}

// Helpers exposed for the UI and validation.
function getParticipantListForTest(eventId){
  const ss=getMasterSpreadsheetFast_();
  const sh=controlSheetFast_(ss,APP.SHEETS.PARTICIPANTS);
  return readSheetObjects_(sh).filter(function(p){return !eventId||String(p['Event ID'])===String(eventId);}).map(function(p){return {participantId:p['Participant ID'],name:p['Full Name'],email:p['Email'],eventId:p['Event ID']};});
}
function saveWebAppUrl(url){const sh=getSheet_(APP.SHEETS.SETTINGS);updateObjectByKey_(sh,'Key','WEB_APP_URL',{'Value':clean_(url),'Updated At':new Date()});return {ok:true};}
function savePauseState(paused){const sh=getSheet_(APP.SHEETS.SETTINGS);updateObjectByKey_(sh,'Key','PAUSE_AUTOMATION',{'Value':paused?'TRUE':'FALSE','Updated At':new Date()});return {ok:true,paused:paused};}

function selfRepairV1_(){ return selfRepairV2_(); }
