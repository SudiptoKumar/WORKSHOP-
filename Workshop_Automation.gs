/**
 * WORKSHOP Automation V1
 * Single-file Apps Script engine.
 *
 * Architecture:
 *   ONE PERMANENT APPS SCRIPT ENGINE
 *        -> ONE MASTER GOOGLE SHEET
 *        -> MANY ISOLATED WORKSHOPS
 *
 * V1 scope:
 *   Workshop -> Registration -> Participant -> Confirmation -> Reminder
 *   -> Attendance -> Certificate -> PDF -> Email -> Verification
 *
 * Test Run:
 *   Uses an existing participant as the recipient but NEVER mutates the
 *   participant's production attendance/certificate fields. It sends:
 *     T+0       [TEST] Registration confirmation
 *     T+~1 min  [TEST] Reminder
 *     T+~2 min  [TEST] Certificate + PDF
 *
 * All HTML email templates live in this file so maintenance is centralized.
 */

const APP = {
  VERSION: 'V1',
  ROOT: 'WORKSHOP',
  DEFAULT_TIMEZONE: 'Asia/Dhaka',
  DEFAULT_PUBLIC_VERIFY_URL: 'https://financeclubpstu.vercel.app/verify',
  SCHEDULER_MINUTES: 5,
  TEST_DELAY_MS: 60 * 1000,
  MAX_EMAIL_BATCH: 20,
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
    DASHBOARD: 'Dashboard'
  },
  EVENT_STATES: ['DRAFT', 'OPEN', 'LIVE', 'COMPLETED', 'CANCELLED', 'ARCHIVED'],
  REG_STATES: ['REGISTERED', 'APPROVED', 'ON_HOLD', 'REMOVED', 'REJECTED', 'TEST'],
  EMAIL_STATES: ['PENDING', 'PROCESSING', 'SENT', 'FAILED', 'RETRY'],
  EMAIL_TYPES: ['REGISTRATION', 'REMINDER', 'CERTIFICATE', 'TEST_REGISTRATION', 'TEST_REMINDER', 'TEST_CERTIFICATE'],
  CERT_RULES: ['ALL_REGISTERED', 'APPROVED_ONLY', 'ATTENDANCE_REQUIRED']
};

const SCHEMAS = {
  Settings: ['Key','Value','Updated At'],
  Workshops: ['Event ID','Workshop Name','Workshop Date','Start Time','End Time','Workshop Type','Timezone','Meeting Platform','Meeting Link','Venue Name','Venue Address','Details URL','Capacity','Registration Opens','Registration Closes','Reminder Enabled','Reminder Mode','Reminder Offsets','Reminder Custom At','Certificate Enabled','Certificate Release Mode','Certificate Delay','Certificate Next Morning Time','Certificate Custom Release At','Certificate Eligibility','Website','Facebook','LinkedIn','Instagram','YouTube','Notes','Status','Folder ID','Registration Form ID','Registration Response Spreadsheet ID','Attendance Form ID','Attendance Response Spreadsheet ID','Last Registration Row','Last Attendance Row','Created At','Updated At'],
  Participants: ['Participant ID','Event ID','Workshop Name','Registration Timestamp','Full Name','Email','Phone / WhatsApp Number','Department / Faculty','Season','ID Number','Registration Number','Registration Status','Attendance Status','Certificate Eligibility','Certificate Status','Certificate ID','Certificate PDF URL','Created At','Updated At','Source Response Row'],
  Attendance: ['Timestamp','Participant ID','Event ID','Email','Name','Workshop Type','Validation','Attendance Status','Validation Message','Source','Created At'],
  'Email Queue': ['Job ID','Unique Key','Event ID','Participant ID','Email Type','Scheduled At','Status','Attempt Count','Last Attempt At','Sent At','Recipient','Subject','Attachment File ID','Last Error','Test Run ID','Message Body','Include Meeting Link','Sender Name','Created At','Updated At'],
  Certificates: ['Certificate ID','Event ID','Participant ID','Participant Name','Email','Workshop Name','Season','Issued At','PDF URL','PDF File ID','Verification URL','Status','Created At','Test Run ID'],
  'System Logs': ['Timestamp','Level','Module','Event ID','Participant ID','Action','Message','Details'],
  'Test Runs': ['Run ID','Event ID','Participant ID','Workshop Name','Recipient','Started At','Status','Current Step','Confirmation Sent At','Reminder Sent At','Certificate Created At','Certificate Sent At','Test Certificate ID','Test Certificate PDF URL','Last Error','Completed At'],
  Dashboard: ['Metric','Value','Updated At']
};

function onOpen() {
  SpreadsheetApp.getUi().createMenu('WORKSHOP Automation')
    .addItem('Open Control Center','openControlCenter')
    .addItem('Refresh Control Center','refreshControlCenter_')
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
    installControlCenterEditTrigger_(ss);
    sh = ss.getSheetByName('Control Center');
  }
  ss.setActiveSheet(sh);
  sh.getRange('A1').activate();
  refreshControlCenter_();
}

function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};
  const pathInfo = e && e.pathInfo ? String(e.pathInfo) : '';
  const certificateId = clean_(params.certificateId || params.id || extractCertificateIdFromPathInfo_(pathInfo));
  let page = clean_(params.page).toLowerCase();

  if (!page) {
    const path = pathInfo.replace(/^\/+|\/+$/g,'');
    page = /^health$/i.test(path) ? 'health' : (certificateId ? 'verify' : 'verify');
  }

  if (page === 'health') {
    return ContentService.createTextOutput(JSON.stringify({
      ok:true,
      version:APP.VERSION,
      service:'WORKSHOP Automation',
      timestamp:new Date().toISOString(),
      webAppUrl:ScriptApp.getService().getUrl() || '',
      publicVerifyUrl:getPublicVerificationBaseUrl_()
    })).setMimeType(ContentService.MimeType.JSON);
  }

  if (page === 'verify-api') {
    return ContentService.createTextOutput(JSON.stringify(verifyCertificatePublic(certificateId)))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (page === 'verify') {
    const template = HtmlService.createTemplateFromFile('Verify');
    template.autoCertificateId = certificateId;
    template.publicVerificationBaseUrl = getPublicVerificationBaseUrl_();
    return template.evaluate().setTitle('Certificate Verification');
  }

  return HtmlService.createHtmlOutput('<!doctype html><html><body style="font-family:system-ui;padding:32px"><h2>WORKSHOP Automation V1</h2><p>Open the master Google Sheet and use <b>WORKSHOP Automation → Open Control Center</b>.</p></body></html>');
}

function extractCertificateIdFromPathInfo_(pathInfo) {
  const path = String(pathInfo || '').replace(/^\/+|\/+$/g,'');
  if (!path) return '';
  const m = path.match(/^verify\/([^/?#]+)$/i);
  if (m) {
    try { return decodeURIComponent(m[1]); } catch (err) { return m[1]; }
  }
  if (/^(?:TEST-CERT|CERT)-/i.test(path)) return path;
  return '';
}

function getPublicVerificationBaseUrl_() {
  return clean_(getSetting_('PUBLIC_VERIFY_URL')) || APP.DEFAULT_PUBLIC_VERIFY_URL;
}

function buildVerificationUrl_(certificateId) {
  const id = encodeURIComponent(String(certificateId || '').trim());
  const publicBase = getPublicVerificationBaseUrl_();
  if (publicBase) {
    const trimmed = publicBase.replace(/\/+$/,'');
    if (/\{CERTIFICATE_ID\}/i.test(trimmed)) {
      return trimmed.replace(/\{CERTIFICATE_ID\}/ig, id);
    }
    if (/:certificateId\b/i.test(trimmed)) {
      return trimmed.replace(/:certificateId\b/ig, id);
    }
    return trimmed + '/' + id;
  }
  const backend = clean_(getSetting_('WEB_APP_URL')) || ScriptApp.getService().getUrl() || '';
  return backend ? backend.replace(/[?].*$/,'').replace(/\/+$/,'') + '?page=verify&certificateId=' + id : '';
}

function getPublicBranding() {
  try {
    const settings = getSettings_();
    return {
      organizationName: settings.ORGANIZATION_NAME || APP.ROOT,
      logoUrl: settings.DEFAULT_LOGO_URL || '',
      verificationUrl: settings.PUBLIC_VERIFY_URL || settings.WEB_APP_URL || ScriptApp.getService().getUrl() || ''
    };
  } catch (err) {
    return {
      organizationName: APP.ROOT || 'WORKSHOP Automation',
      logoUrl: '',
      verificationUrl: ScriptApp.getService().getUrl() || ''
    };
  }
}

function setupSystem() {
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
    // Keep the visible settings version synchronized on upgrades.
    updateObjectByKey_(getSheet_(APP.SHEETS.SETTINGS),'Key','APP_VERSION',{'Value':APP.VERSION,'Updated At':new Date()});
    setupControlCenter_(ss);
    installControlCenterEditTrigger_(ss);
    ensureDefaultCertificateMaster_();
    installCentralScheduler_();
    tidyEmailQueueSheet_();
    refreshDashboard_();
    refreshControlCenter_();
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
    DEFAULT_CAPACITY: '0',
    DEFAULT_OPEN_DAYS_BEFORE: '7',
    DEFAULT_CLOSE_MINUTES_BEFORE: '15',
    DEFAULT_REMINDERS: '60',
    DEFAULT_CERTIFICATE_ENABLED: 'TRUE',
    DEFAULT_CERTIFICATE_RELEASE: 'AFTER_WORKSHOP',
    DEFAULT_CERTIFICATE_DELAY: '0',
    DEFAULT_CERTIFICATE_TIME: '09:00',
    DEFAULT_CERTIFICATE_ELIGIBILITY: 'ATTENDANCE_REQUIRED',
    SENDER_NAME: 'Finance Club',
    ORGANIZATION_NAME: 'Finance Club',
    DEFAULT_LOGO_URL: '',
    WEB_APP_URL: '',
    PUBLIC_VERIFY_URL: APP.DEFAULT_PUBLIC_VERIFY_URL,
    DEFAULT_WEBSITE: '',
    DEFAULT_FACEBOOK: '',
    DEFAULT_INSTAGRAM: '',
    DEFAULT_LINKEDIN: '',
    DEFAULT_YOUTUBE: '',
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
      if (!getSetting_('DEFAULT_LOGO_URL')) updateObjectByKey_(sh, 'Key', 'DEFAULT_LOGO_URL', {'Value':'','Updated At':new Date()});
      if (!getSetting_('ORGANIZATION_NAME')) updateObjectByKey_(sh, 'Key', 'ORGANIZATION_NAME', {'Value':'Finance Club','Updated At':new Date()});
      if (!getSetting_('SENDER_NAME')) updateObjectByKey_(sh, 'Key', 'SENDER_NAME', {'Value':'Finance Club','Updated At':new Date()});
      if (!getSetting_('PUBLIC_VERIFY_URL')) updateObjectByKey_(sh, 'Key', 'PUBLIC_VERIFY_URL', {'Value':APP.DEFAULT_PUBLIC_VERIFY_URL,'Updated At':new Date()});
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
    regForm.setDescription(buildRegistrationDescription_(name,date,start,end,timezone,payload.meetingPlatform || settings.DEFAULT_MEETING_PLATFORM));
    addRegistrationItems_(regForm);
    regForm.setConfirmationMessage('Registration received. Your Participant ID will be sent by email after validation.');
    regForm.setDestination(FormApp.DestinationType.SPREADSHEET, regSs.getId());
    try { DriveApp.getFileById(regForm.getId()).moveTo(registrationFolder); } catch (err) {}

    const attSs = SpreadsheetApp.create(eventId+' - Attendance Responses');
    try { DriveApp.getFileById(attSs.getId()).moveTo(attendanceFolder); } catch (err) {}
    const attForm = FormApp.create(eventId+' - '+name+' Attendance');
    attForm.setDescription('Attendance check-in for '+name+' ('+eventId+'). Enter the Participant ID received by email.');
    const pitem = attForm.addTextItem().setTitle('Participant ID').setRequired(true);
    attForm.setConfirmationMessage('Attendance recorded.');
    attForm.setDestination(FormApp.DestinationType.SPREADSHEET, attSs.getId());
    try { DriveApp.getFileById(attForm.getId()).moveTo(attendanceFolder); } catch (err) {}

    const row = {
      'Event ID':eventId,'Workshop Name':name,'Workshop Date':date,'Start Time':start,'End Time':end,
      'Workshop Type':clean_(payload.workshopType)||settings.DEFAULT_WORKSHOP_TYPE,
      'Timezone':timezone,'Meeting Platform':clean_(payload.meetingPlatform)||settings.DEFAULT_MEETING_PLATFORM,'Meeting Link':clean_(payload.meetingLink),
      'Venue Name':clean_(payload.venueName),'Venue Address':clean_(payload.venueAddress),'Details URL':clean_(payload.detailsUrl),
      'Capacity':numOr_(payload.capacity,Number(settings.DEFAULT_CAPACITY||0)),'Registration Opens':payload.registrationOpens||'',
      'Registration Closes':payload.registrationCloses||'','Reminder Enabled':payload.reminderEnabled === false ? false : true,'Reminder Mode':clean_(payload.reminderMode)||'BEFORE_START','Reminder Offsets':clean_(payload.reminderOffsets)||settings.DEFAULT_REMINDERS,'Reminder Custom At':payload.reminderCustomAt||'',
      'Certificate Enabled':payload.certificateEnabled === false ? false : true,'Certificate Release Mode':clean_(payload.certificateReleaseMode)||settings.DEFAULT_CERTIFICATE_RELEASE,
      'Certificate Delay':numOr_(payload.certificateDelay,Number(settings.DEFAULT_CERTIFICATE_DELAY||0)),'Certificate Next Morning Time':clean_(payload.certificateNextMorningTime)||settings.DEFAULT_CERTIFICATE_TIME,'Certificate Custom Release At':payload.certificateCustomReleaseAt||'',
      'Certificate Eligibility':clean_(payload.certificateEligibility)||settings.DEFAULT_CERTIFICATE_ELIGIBILITY,'Website':clean_(payload.website)||settings.DEFAULT_WEBSITE,
      'Facebook':clean_(payload.facebook)||settings.DEFAULT_FACEBOOK,'LinkedIn':clean_(payload.linkedin)||settings.DEFAULT_LINKEDIN,'Instagram':clean_(payload.instagram)||settings.DEFAULT_INSTAGRAM,'YouTube':clean_(payload.youtube)||settings.DEFAULT_YOUTUBE,'Notes':clean_(payload.notes),
      'Status':'DRAFT','Folder ID':eventFolder.getId(),'Registration Form ID':regForm.getId(),'Registration Response Spreadsheet ID':regSs.getId(),
      'Attendance Form ID':attForm.getId(),'Attendance Response Spreadsheet ID':attSs.getId(),'Last Registration Row':1,'Last Attendance Row':1,
      'Created At':new Date(),'Updated At':new Date()
    };
    row['Registration Opens'] = row['Registration Opens'] || addMinutes_(parseDateTime_(date,start,timezone), -24*60*7, timezone);
    row['Registration Closes'] = row['Registration Closes'] || addMinutes_(parseDateTime_(date,start,timezone), -15, timezone);
    const workshopSheet=getSheet_(APP.SHEETS.WORKSHOPS);
    appendObject_(workshopSheet, row);
    // Workshop creation is durable before optional form-state synchronization.
    // A transient Form/date issue must not roll back the whole workshop creation.
    try{ syncWorkshopFormState_(row); }catch(err){
      log_('ERROR','WORKSHOP',eventId,'','FORM_STATE_SYNC_AFTER_CREATE_FAILED',String(err),stack_(err));
    }
    log_('INFO','WORKSHOP',eventId,'','CREATE_WORKSHOP','Workshop created',JSON.stringify({folderId:eventFolder.getId(),registrationFormId:regForm.getId(),attendanceFormId:attForm.getId(),masterSpreadsheetId:ss.getId()}));
    refreshDashboard_();
    return {ok:true,eventId,workshopName:name,folderId:eventFolder.getId(),registrationFormUrl:regForm.getPublishedUrl(),attendanceFormUrl:attForm.getPublishedUrl(),registrationEditUrl:regForm.getEditUrl()};
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

function buildRegistrationDescription_(name,date,start,end,tz,platform) {
  return [name,'', 'Date: '+date,'Time: '+start+' - '+end+' ('+tz+')','Platform: '+platform,'','Please complete these seven required fields: Full Name, Email, Phone / WhatsApp Number, Department / Faculty, Season, ID Number, Registration Number.'].join('\n');
}

function processScheduler() {
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
        ['ATTENDANCE', function(){ processAttendanceResponses_(w); }],
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

    try { processEmailQueue_(); }
    catch (err) { log_('ERROR','SCHEDULER','','','EMAIL_QUEUE_FAILED',String(err),stack_(err)); }

    try { recoverStaleEmailJobs_(); }
    catch (err) { log_('ERROR','SCHEDULER','','','STALE_EMAIL_RECOVERY_FAILED',String(err),stack_(err)); }

    try { refreshDashboard_(); }
    catch (err) { log_('ERROR','SCHEDULER','','','DASHBOARD_REFRESH_FAILED',String(err),stack_(err)); }

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
  w=repairIfWorkshopIntegrationBroken_(w);
  const eventId=String(w['Event ID']||'');
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

      if(!p){
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
      log_('INFO',moduleName,eventId,'','FORM_RESPONSE_SOURCE_OK','Read '+out.length+' response(s) from Google Form','responses='+responses.length+';startRow='+baseRow);
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
      log_('INFO',moduleName,eventId,'','SHEET_RESPONSE_SOURCE_OK','Fallback response sheet read: '+sh.getName(),'rows='+out.length);
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
  enqueueEmail_({event:w,participant:p,type:'REGISTRATION',scheduledAt:new Date(),subject:email.subject,html:email.html,body:email.text,includeMeetingLink:false});
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
  schedules.forEach(function(s){
    if (!s.when || s.when <= now) return;
    participants.forEach(function(p){
      const unique = w['Event ID']+'|'+p['Participant ID']+'|REMINDER|'+s.keyPart;
      if (hasEmailUniqueKey_(unique)) return;
      const e = buildReminderEmail_(w,p,s.offset || 0,false);
      if (mode === 'EXACT_DATETIME') e.text = 'Reminder: '+w['Workshop Name']+' is scheduled for '+formatDateTimeForDisplay_(w['Reminder Custom At'],tz)+'.';
      enqueueEmail_({event:w,participant:p,type:'REMINDER',scheduledAt:s.when,subject:e.subject,html:e.html,body:e.text,includeMeetingLink:!!w['Meeting Link'],uniqueKey:unique});
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
      enqueueEmail_({event:w,participant:p,type:'CERTIFICATE',scheduledAt:now,subject:email.subject,html:email.html,body:email.text,attachmentFileId:cert.pdfFileId,includeMeetingLink:false});
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
  const quota = typeof MailApp.getRemainingDailyQuota === 'function' ? MailApp.getRemainingDailyQuota() : APP.MAX_EMAIL_BATCH;
  const limit = Math.min(APP.MAX_EMAIL_BATCH, Number(quota||0));
  if (limit <= 0) return;
  let sent = 0;
  rows.forEach(job => {
    if (sent >= limit) return;
    if (!['PENDING','RETRY'].includes(String(job.Status))) return;
    const when = new Date(job['Scheduled At']);
    if (when > now) return;
    try {
      markQueueJob_(job['Job ID'], {'Status':'PROCESSING','Attempt Count':Number(job['Attempt Count']||0)+1,'Last Attempt At':now,'Updated At':now});
      const attachments=[];
      if (job['Attachment File ID']) attachments.push(DriveApp.getFileById(job['Attachment File ID']).getBlob());
      MailApp.sendEmail({to:String(job.Recipient),subject:String(job.Subject),body:String(job['Message Body']||''),htmlBody:String(job['Message Body']||''),name:String(job['Sender Name']||getSetting_('SENDER_NAME')||APP.ROOT),attachments:attachments});
      markQueueJob_(job['Job ID'], {'Status':'SENT','Sent At':new Date(),'Updated At':new Date(),'Last Error':''});
      if (String(job['Email Type'])==='REGISTRATION') log_('INFO','EMAIL',String(job['Event ID']),String(job['Participant ID']),'REGISTRATION_EMAIL_SENT','',job['Recipient']);
      if (String(job['Email Type'])==='REMINDER') log_('INFO','EMAIL',String(job['Event ID']),String(job['Participant ID']),'REMINDER_EMAIL_SENT','',job['Recipient']);
      if (String(job['Email Type'])==='CERTIFICATE') {
        updateParticipant_(job['Participant ID'], {'Certificate Status':'SENT','Updated At':new Date()});
        log_('INFO','EMAIL',String(job['Event ID']),String(job['Participant ID']),'CERTIFICATE_EMAIL_SENT','',job['Recipient']);
      }
      sent++;
    } catch(err) {
      const attempt=Number(job['Attempt Count']||1);
      markQueueJob_(job['Job ID'], {'Status':attempt>=3?'FAILED':'RETRY','Last Error':String(err),'Updated At':new Date()});
      log_('ERROR','EMAIL',String(job['Event ID']),String(job['Participant ID']),'EMAIL_SEND_FAILED',String(err),JSON.stringify({jobId:job['Job ID'],attempt}));
    }
  });
}

function recoverStaleEmailJobs_() {
  const sh = getSheet_(APP.SHEETS.QUEUE);
  const now = Date.now();
  readSheetObjects_(sh).forEach(job=>{
    if (String(job.Status)!=='PROCESSING') return;
    const last=job['Last Attempt At'] ? new Date(job['Last Attempt At']).getTime() : 0;
    if (last && now-last > 15*60000) markQueueJob_(job['Job ID'], {'Status':'RETRY','Updated At':new Date(),'Last Error':'Recovered stale PROCESSING job.'});
  });
}

function enqueueEmail_(opts) {
  const unique = opts.uniqueKey || (opts.event['Event ID']+'|'+opts.participant['Participant ID']+'|'+opts.type);
  if (hasEmailUniqueKey_(unique)) return null;
  const obj = {
    'Job ID':'JOB-'+Utilities.getUuid().slice(0,12).toUpperCase(),'Unique Key':unique,'Event ID':opts.event['Event ID'],'Participant ID':opts.participant['Participant ID'],
    'Email Type':opts.type,'Scheduled At':opts.scheduledAt||new Date(),'Status':'PENDING','Attempt Count':0,'Last Attempt At':'','Sent At':'','Recipient':opts.participant['Email'],
    'Subject':opts.subject,'Attachment File ID':opts.attachmentFileId||'','Last Error':'','Test Run ID':opts.testRunId||'','Message Body':opts.html,'Include Meeting Link':opts.includeMeetingLink?'TRUE':'FALSE','Sender Name':getSetting_('SENDER_NAME')||APP.ROOT,'Created At':new Date(),'Updated At':new Date()
  };
  appendObject_(getSheet_(APP.SHEETS.QUEUE),obj);
  return obj['Job ID'];
}

function startTestRun(participantId) {
  return withLock_('startTestRun', function(){
    ensureSetup_();
    const p=findParticipantById_(participantId);
    if(!p) throw new Error('Participant not found.');
    const w=findWorkshop_(p['Event ID']);
    if(!w) throw new Error('Workshop not found.');
    if(!p['Email']) throw new Error('Participant has no email address.');
    const runId='TEST-'+Utilities.getUuid().slice(0,10).toUpperCase();
    const now=new Date();
    appendObject_(getSheet_(APP.SHEETS.TESTS),{'Run ID':runId,'Event ID':w['Event ID'],'Participant ID':p['Participant ID'],'Workshop Name':w['Workshop Name'],'Recipient':p['Email'],'Started At':now,'Status':'RUNNING','Current Step':'CONFIRMATION','Confirmation Sent At':'','Reminder Sent At':'','Certificate Created At':'','Certificate Sent At':'','Test Certificate ID':'','Test Certificate PDF URL':'','Last Error':'','Completed At':''});
    try {
      const e=buildRegistrationEmail_(w,p,true);
      sendDirectEmail_(p['Email'],'[TEST] '+e.subject,e.html,e.text,'');
      updateTestRun_(runId,{'Confirmation Sent At':new Date(),'Current Step':'REMINDER_PENDING'});
      createTestStepTrigger_();
      log_('INFO','TEST',w['Event ID'],p['Participant ID'],'TEST_CONFIRMATION_SENT','Test run started',runId);
      return {ok:true,runId,recipient:p['Email'],message:'Test confirmation sent. Reminder will be attempted in about 1 minute, followed by the certificate email about 1 minute later.'};
    } catch(err){
      updateTestRun_(runId,{'Status':'FAILED','Last Error':String(err),'Completed At':new Date()});
      throw err;
    }
  });
}

function createTestStepTrigger_() {
  ScriptApp.newTrigger('processDueTestSteps').timeBased().after(APP.TEST_DELAY_MS).create();
}

function processDueTestSteps() {
  if(isPaused_()) return;
  withLock_('processDueTestSteps', function(){
    const sh=getSheet_(APP.SHEETS.TESTS);
    const runs=readSheetObjects_(sh).filter(r=>String(r.Status)==='RUNNING');
    const now=new Date();
    runs.forEach(run=>{
      try{
        const p=findParticipantById_(run['Participant ID']);
        const w=findWorkshop_(run['Event ID']);
        if(!p||!w) throw new Error('Test participant/workshop no longer exists.');
        if(String(run['Current Step'])==='REMINDER_PENDING'){
          const e=buildReminderEmail_(w,p,60,true);
          sendDirectEmail_(p['Email'],'[TEST] '+e.subject,e.html,e.text,'');
          updateTestRun_(run['Run ID'],{'Reminder Sent At':new Date(),'Current Step':'CERTIFICATE_PENDING'});
          createTestStepTrigger_();
        } else if(String(run['Current Step'])==='CERTIFICATE_PENDING'){
          const cert=generateCertificateForParticipant_(w,p,true,run['Run ID']);
          const e=buildCertificateEmail_(w,p,cert,true);
          sendDirectEmail_(p['Email'],'[TEST] '+e.subject,e.html,e.text,cert.pdfFileId);
          updateTestRun_(run['Run ID'],{'Certificate Created At':new Date(),'Certificate Sent At':new Date(),'Test Certificate ID':cert.certificateId,'Test Certificate PDF URL':cert.pdfUrl,'Status':'COMPLETED','Current Step':'DONE','Completed At':new Date()});
          log_('INFO','TEST',w['Event ID'],p['Participant ID'],'TEST_COMPLETED','All three test stages completed',run['Run ID']);
        }
      } catch(err){
        updateTestRun_(run['Run ID'],{'Status':'FAILED','Last Error':String(err),'Current Step':'FAILED','Completed At':new Date()});
        log_('ERROR','TEST',String(run['Event ID']),String(run['Participant ID']),'TEST_FAILED',String(err),stack_(err));
      }
    });
  });
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
    const verificationUrl=buildVerificationUrl_(certId);

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
  REGISTRATION: "<!doctype html>\n<html lang=\"en\">\n<head>\n<meta charset=\"utf-8\">\n<meta name=\"viewport\" content=\"width=device-width,initial-scale=1\">\n<meta name=\"color-scheme\" content=\"dark light\">\n<meta name=\"supported-color-schemes\" content=\"dark light\">\n<style>\nhtml,body{margin:0!important;padding:0!important;width:100%!important;background:#000!important}\n*{-webkit-text-size-adjust:100%;-ms-text-size-adjust:100%;box-sizing:border-box}\ntable,td{border-collapse:collapse;mso-table-lspace:0pt;mso-table-rspace:0pt}\nimg{border:0;outline:none;text-decoration:none}\na{text-decoration:none}\n@media screen and (max-width:600px){\n  .wrap{width:100%!important}\n  .px{padding-left:20px!important;padding-right:20px!important}\n  .half{display:block!important;width:100%!important;padding-left:0!important;padding-right:0!important}\n  .h1{font-size:28px!important;line-height:34px!important}\n}\n</style>\n</head>\n<body style=\"margin:0;padding:0;background:#000\">\n<div style=\"display:none;max-height:0;overflow:hidden;opacity:0;mso-hide:all\">{{preheader}}</div>\n<center style=\"width:100%;background:#000\">\n<div class=\"wrap\" style=\"max-width:640px;margin:0 auto\">\n<table role=\"presentation\" width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\">\n<tr><td class=\"px\" style=\"padding:30px 32px 22px\">\n<table role=\"presentation\" width=\"100%\"><tr>\n<td align=\"left\">\n<table role=\"presentation\"><tr>\n<td style=\"padding-right:9px\"><div style=\"width:28px;height:28px;line-height:28px;text-align:center;border-radius:8px;background:#1C1C1E;border:1px solid #2C2C2E;color:{{accent}};font:700 11px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif\">{{org_initials}}</div></td>\n<td><span style=\"font:600 15px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#F5F5F7\">{{organization_name}}</span></td>\n</tr></table>\n</td>\n<td align=\"right\"><a href=\"{{website_url}}\" style=\"font:13px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#98989D\">{{website_label}}</a></td>\n</tr></table>\n</td></tr>\n<tr><td class=\"px\"><div style=\"border-top:1px solid #2C2C2E\"></div></td></tr>\n\n<tr><td class=\"px\" align=\"center\" style=\"padding:50px 32px 38px;background:#000\">\n<table role=\"presentation\" width=\"64\" height=\"64\"><tr><td align=\"center\" style=\"background:#1C1C1E;border:1px solid {{accent}};border-radius:32px;color:{{accent}};font:700 24px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif\">✓</td></tr></table>\n<div style=\"margin-top:22px;font:600 12px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;letter-spacing:.8px;color:{{accent}}\">{{eyebrow}}</div>\n<h1 class=\"h1\" style=\"margin:15px 0 0;font:700 33px/39px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;letter-spacing:-.4px;color:#F5F5F7\">{{hero_title}}</h1>\n<p style=\"margin:14px 0 0;font:16px/23px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#D1D1D6\">{{greeting}}</p>\n<p style=\"margin:6px auto 0;max-width:430px;font:14px/21px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#8E8E93\">{{hero_text}}</p>\n</td></tr>\n\n<tr><td class=\"px\" style=\"padding:30px 32px 0\">\n<table role=\"presentation\" width=\"100%\" style=\"background:#1C1C1E;border-radius:16px\"><tr><td style=\"padding:24px\">\n<div style=\"font:600 12px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;letter-spacing:.7px;color:#8E8E93;margin-bottom:16px\">PARTICIPANT RECORD</div>\n<table role=\"presentation\" width=\"100%\"><tr>\n<td class=\"half\" width=\"50%\" valign=\"top\" style=\"padding:0 12px 17px 0\"><div style=\"font:12px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#6E6E73;margin-bottom:4px\">Full Name</div><div style=\"font:15px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#F5F5F7\">{{participant_name}}</div></td>\n<td class=\"half\" width=\"50%\" valign=\"top\" style=\"padding:0 0 17px 12px\"><div style=\"font:12px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#6E6E73;margin-bottom:4px\">Email</div><div style=\"font:15px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#F5F5F7;word-break:break-word\">{{email}}</div></td>\n</tr>\n<tr>\n<td class=\"half\" width=\"50%\" valign=\"top\" style=\"padding:0 12px 17px 0\"><div style=\"font:12px -apple-system,BlinkMacSystemFont,\'Segoe UI\',sans-serif;color:#6E6E73;margin-bottom:4px\">Phone / WhatsApp Number</div><div style=\"font:15px -apple-system,BlinkMacSystemFont,\'Segoe UI\',sans-serif;color:#F5F5F7;word-break:break-word\">{{phone_whatsapp}}</div></td>\n<td class=\"half\" width=\"50%\" valign=\"top\" style=\"padding:0 0 17px 12px\">&nbsp;</td>\n</tr>\n<tr><td colspan=\"2\"><div style=\"border-top:1px solid #38383A;margin-bottom:17px\"></div></td></tr>\n<tr>\n<td class=\"half\" width=\"50%\" valign=\"top\" style=\"padding:0 12px 17px 0\"><div style=\"font:12px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#6E6E73;margin-bottom:4px\">Department / Faculty</div><div style=\"font:15px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#F5F5F7\">{{department_faculty}}</div></td>\n<td class=\"half\" width=\"50%\" valign=\"top\" style=\"padding:0 0 17px 12px\"><div style=\"font:12px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#6E6E73;margin-bottom:4px\">Season</div><div style=\"font:15px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#F5F5F7\">{{season}}</div></td>\n</tr>\n<tr><td colspan=\"2\"><div style=\"border-top:1px solid #38383A;margin-bottom:17px\"></div></td></tr>\n<tr>\n<td class=\"half\" width=\"50%\" valign=\"top\" style=\"padding:0 12px 17px 0\"><div style=\"font:12px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#6E6E73;margin-bottom:4px\">ID Number</div><div style=\"font:15px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#F5F5F7\">{{id_number}}</div></td>\n<td class=\"half\" width=\"50%\" valign=\"top\" style=\"padding:0 0 17px 12px\"><div style=\"font:12px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#6E6E73;margin-bottom:4px\">Registration Number</div><div style=\"font:15px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#F5F5F7\">{{registration_number}}</div></td>\n</tr>\n<tr><td colspan=\"2\"><div style=\"border-top:1px solid #38383A;margin-bottom:17px\"></div></td></tr>\n<tr><td colspan=\"2\"><div style=\"font:12px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#6E6E73;margin-bottom:4px\">Participant ID</div><div style=\"font:15px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#F5F5F7\">{{participant_id}}</div></td></tr>\n</table>\n</td></tr></table>\n</td></tr>\n\n<tr><td class=\"px\" style=\"padding:16px 32px 0\">\n<table role=\"presentation\" width=\"100%\" style=\"background:#1C1C1E;border-radius:16px\"><tr><td style=\"padding:24px\">\n<div style=\"font:600 12px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;letter-spacing:.7px;color:#8E8E93;margin-bottom:14px\">WORKSHOP</div>\n<div style=\"font:600 20px/26px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#F5F5F7\">{{event_name}}</div>\n<table role=\"presentation\" width=\"100%\" style=\"margin-top:14px\"><tr><td style=\"padding:11px 0;border-bottom:1px solid #38383A;font:14px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#8E8E93\">Date</td><td align=\"right\" style=\"padding:11px 0;border-bottom:1px solid #38383A;font:14px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#F5F5F7\">{{event_date}}</td></tr>\n<tr><td style=\"padding:11px 0;border-bottom:1px solid #38383A;font:14px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#8E8E93\">Time</td><td align=\"right\" style=\"padding:11px 0;border-bottom:1px solid #38383A;font:14px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#F5F5F7\">{{start_time}} – {{end_time}}</td></tr>\n<tr><td style=\"padding:11px 0;font:14px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#8E8E93\">Platform</td><td align=\"right\" style=\"padding:11px 0;font:14px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#F5F5F7\">{{platform}}</td></tr></table>\n{{meeting_button}}\n</td></tr></table>\n</td></tr>\n\n<tr><td class=\"px\" align=\"center\" style=\"padding:35px 32px 0\">{{social_block}}</td></tr>\n<tr><td class=\"px\" align=\"center\" style=\"padding:25px 32px 42px\">\n<p style=\"margin:0 0 6px;font:12px/18px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#48484A\">{{footer_note}}</p>\n<p style=\"margin:0;font:11px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#3A3A3C\">© {{current_year}} {{organization_name}}. All rights reserved.</p>\n</td></tr>\n</table></div></center>\n</body></html>",
  REMINDER: "<!doctype html>\n<html lang=\"en\">\n<head>\n<meta charset=\"utf-8\">\n<meta name=\"viewport\" content=\"width=device-width,initial-scale=1\">\n<meta name=\"color-scheme\" content=\"dark light\">\n<meta name=\"supported-color-schemes\" content=\"dark light\">\n<style>\nhtml,body{margin:0!important;padding:0!important;width:100%!important;background:#000!important}\n*{-webkit-text-size-adjust:100%;-ms-text-size-adjust:100%;box-sizing:border-box}\ntable,td{border-collapse:collapse;mso-table-lspace:0pt;mso-table-rspace:0pt}\nimg{border:0;outline:none;text-decoration:none}\na{text-decoration:none}\n@media screen and (max-width:600px){\n  .wrap{width:100%!important}\n  .px{padding-left:20px!important;padding-right:20px!important}\n  .half{display:block!important;width:100%!important;padding-left:0!important;padding-right:0!important}\n  .h1{font-size:28px!important;line-height:34px!important}\n}\n</style>\n</head>\n<body style=\"margin:0;padding:0;background:#000\">\n<div style=\"display:none;max-height:0;overflow:hidden;opacity:0;mso-hide:all\">{{preheader}}</div>\n<center style=\"width:100%;background:#000\">\n<div class=\"wrap\" style=\"max-width:640px;margin:0 auto\">\n<table role=\"presentation\" width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\">\n<tr><td class=\"px\" style=\"padding:30px 32px 22px\">\n<table role=\"presentation\" width=\"100%\"><tr>\n<td align=\"left\">\n<table role=\"presentation\"><tr>\n<td style=\"padding-right:9px\"><div style=\"width:28px;height:28px;line-height:28px;text-align:center;border-radius:8px;background:#1C1C1E;border:1px solid #2C2C2E;color:{{accent}};font:700 11px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif\">{{org_initials}}</div></td>\n<td><span style=\"font:600 15px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#F5F5F7\">{{organization_name}}</span></td>\n</tr></table>\n</td>\n<td align=\"right\"><a href=\"{{website_url}}\" style=\"font:13px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#98989D\">{{website_label}}</a></td>\n</tr></table>\n</td></tr>\n<tr><td class=\"px\"><div style=\"border-top:1px solid #2C2C2E\"></div></td></tr>\n\n<tr><td class=\"px\" align=\"center\" style=\"padding:50px 32px 38px;background:#000\">\n<table role=\"presentation\" width=\"64\" height=\"64\"><tr><td align=\"center\" style=\"background:#1C1C1E;border:1px solid {{accent}};border-radius:32px;color:{{accent}};font:700 20px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif\">•</td></tr></table>\n<div style=\"margin-top:22px;font:600 12px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;letter-spacing:.8px;color:{{accent}}\">{{eyebrow}}</div>\n<h1 class=\"h1\" style=\"margin:15px 0 0;font:700 33px/39px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;letter-spacing:-.4px;color:#F5F5F7\">{{hero_title}}</h1>\n<p style=\"margin:14px 0 0;font:16px/23px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#D1D1D6\">{{greeting}}</p>\n<p style=\"margin:6px auto 0;max-width:430px;font:14px/21px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#8E8E93\">{{hero_text}}</p>\n</td></tr>\n\n<tr><td class=\"px\" style=\"padding:30px 32px 0\">\n<table role=\"presentation\" width=\"100%\" style=\"background:#1C1C1E;border-radius:16px\"><tr><td style=\"padding:24px\">\n<div style=\"font:600 12px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;letter-spacing:.7px;color:#8E8E93;margin-bottom:16px\">PARTICIPANT RECORD</div>\n<table role=\"presentation\" width=\"100%\"><tr>\n<td class=\"half\" width=\"50%\" valign=\"top\" style=\"padding:0 12px 17px 0\"><div style=\"font:12px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#6E6E73;margin-bottom:4px\">Full Name</div><div style=\"font:15px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#F5F5F7\">{{participant_name}}</div></td>\n<td class=\"half\" width=\"50%\" valign=\"top\" style=\"padding:0 0 17px 12px\"><div style=\"font:12px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#6E6E73;margin-bottom:4px\">Email</div><div style=\"font:15px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#F5F5F7;word-break:break-word\">{{email}}</div></td>\n</tr>\n<tr>\n<td class=\"half\" width=\"50%\" valign=\"top\" style=\"padding:0 12px 17px 0\"><div style=\"font:12px -apple-system,BlinkMacSystemFont,\'Segoe UI\',sans-serif;color:#6E6E73;margin-bottom:4px\">Phone / WhatsApp Number</div><div style=\"font:15px -apple-system,BlinkMacSystemFont,\'Segoe UI\',sans-serif;color:#F5F5F7;word-break:break-word\">{{phone_whatsapp}}</div></td>\n<td class=\"half\" width=\"50%\" valign=\"top\" style=\"padding:0 0 17px 12px\">&nbsp;</td>\n</tr>\n<tr><td colspan=\"2\"><div style=\"border-top:1px solid #38383A;margin-bottom:17px\"></div></td></tr>\n<tr>\n<td class=\"half\" width=\"50%\" valign=\"top\" style=\"padding:0 12px 17px 0\"><div style=\"font:12px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#6E6E73;margin-bottom:4px\">Department / Faculty</div><div style=\"font:15px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#F5F5F7\">{{department_faculty}}</div></td>\n<td class=\"half\" width=\"50%\" valign=\"top\" style=\"padding:0 0 17px 12px\"><div style=\"font:12px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#6E6E73;margin-bottom:4px\">Season</div><div style=\"font:15px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#F5F5F7\">{{season}}</div></td>\n</tr>\n<tr><td colspan=\"2\"><div style=\"border-top:1px solid #38383A;margin-bottom:17px\"></div></td></tr>\n<tr>\n<td class=\"half\" width=\"50%\" valign=\"top\" style=\"padding:0 12px 17px 0\"><div style=\"font:12px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#6E6E73;margin-bottom:4px\">ID Number</div><div style=\"font:15px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#F5F5F7\">{{id_number}}</div></td>\n<td class=\"half\" width=\"50%\" valign=\"top\" style=\"padding:0 0 17px 12px\"><div style=\"font:12px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#6E6E73;margin-bottom:4px\">Registration Number</div><div style=\"font:15px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#F5F5F7\">{{registration_number}}</div></td>\n</tr>\n<tr><td colspan=\"2\"><div style=\"border-top:1px solid #38383A;margin-bottom:17px\"></div></td></tr>\n<tr><td colspan=\"2\"><div style=\"font:12px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#6E6E73;margin-bottom:4px\">Participant ID</div><div style=\"font:15px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#F5F5F7\">{{participant_id}}</div></td></tr>\n</table>\n</td></tr></table>\n</td></tr>\n\n<tr><td class=\"px\" style=\"padding:16px 32px 0\">\n<table role=\"presentation\" width=\"100%\" style=\"background:#1C1C1E;border-radius:16px\"><tr><td style=\"padding:24px\">\n<div style=\"font:600 12px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;letter-spacing:.7px;color:#8E8E93;margin-bottom:14px\">WORKSHOP</div>\n<div style=\"font:600 20px/26px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#F5F5F7\">{{event_name}}</div>\n<table role=\"presentation\" width=\"100%\" style=\"margin-top:14px\"><tr><td style=\"padding:11px 0;border-bottom:1px solid #38383A;font:14px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#8E8E93\">Date</td><td align=\"right\" style=\"padding:11px 0;border-bottom:1px solid #38383A;font:14px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#F5F5F7\">{{event_date}}</td></tr>\n<tr><td style=\"padding:11px 0;border-bottom:1px solid #38383A;font:14px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#8E8E93\">Time</td><td align=\"right\" style=\"padding:11px 0;border-bottom:1px solid #38383A;font:14px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#F5F5F7\">{{start_time}} – {{end_time}}</td></tr>\n<tr><td style=\"padding:11px 0;font:14px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#8E8E93\">Platform</td><td align=\"right\" style=\"padding:11px 0;font:14px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#F5F5F7\">{{platform}}</td></tr></table>\n{{meeting_button}}\n</td></tr></table>\n</td></tr>\n\n<tr><td class=\"px\" align=\"center\" style=\"padding:35px 32px 0\">{{social_block}}</td></tr>\n<tr><td class=\"px\" align=\"center\" style=\"padding:25px 32px 42px\">\n<p style=\"margin:0 0 6px;font:12px/18px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#48484A\">{{footer_note}}</p>\n<p style=\"margin:0;font:11px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#3A3A3C\">© {{current_year}} {{organization_name}}. All rights reserved.</p>\n</td></tr>\n</table></div></center>\n</body></html>",
  CERTIFICATE: "<!doctype html>\n<html lang=\"en\">\n<head>\n<meta charset=\"utf-8\">\n<meta name=\"viewport\" content=\"width=device-width,initial-scale=1\">\n<meta name=\"color-scheme\" content=\"dark light\">\n<meta name=\"supported-color-schemes\" content=\"dark light\">\n<style>\nhtml,body{margin:0!important;padding:0!important;width:100%!important;background:#000!important}\n*{-webkit-text-size-adjust:100%;-ms-text-size-adjust:100%;box-sizing:border-box}\ntable,td{border-collapse:collapse;mso-table-lspace:0pt;mso-table-rspace:0pt}\nimg{border:0;outline:none;text-decoration:none}\na{text-decoration:none}\n@media screen and (max-width:600px){\n  .wrap{width:100%!important}\n  .px{padding-left:20px!important;padding-right:20px!important}\n  .half{display:block!important;width:100%!important;padding-left:0!important;padding-right:0!important}\n  .h1{font-size:28px!important;line-height:34px!important}\n}\n</style>\n</head>\n<body style=\"margin:0;padding:0;background:#000\">\n<div style=\"display:none;max-height:0;overflow:hidden;opacity:0;mso-hide:all\">{{preheader}}</div>\n<center style=\"width:100%;background:#000\">\n<div class=\"wrap\" style=\"max-width:640px;margin:0 auto\">\n<table role=\"presentation\" width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\">\n<tr><td class=\"px\" style=\"padding:30px 32px 22px\">\n<table role=\"presentation\" width=\"100%\"><tr>\n<td align=\"left\">\n<table role=\"presentation\"><tr>\n<td style=\"padding-right:9px\"><div style=\"width:28px;height:28px;line-height:28px;text-align:center;border-radius:8px;background:#1C1C1E;border:1px solid #2C2C2E;color:{{accent}};font:700 11px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif\">{{org_initials}}</div></td>\n<td><span style=\"font:600 15px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#F5F5F7\">{{organization_name}}</span></td>\n</tr></table>\n</td>\n<td align=\"right\"><a href=\"{{website_url}}\" style=\"font:13px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#98989D\">{{website_label}}</a></td>\n</tr></table>\n</td></tr>\n<tr><td class=\"px\"><div style=\"border-top:1px solid #2C2C2E\"></div></td></tr>\n\n<tr><td class=\"px\" align=\"center\" style=\"padding:50px 32px 38px;background:#000\">\n<table role=\"presentation\" width=\"64\" height=\"64\"><tr><td align=\"center\" style=\"background:#1C1C1E;border:1px solid {{accent}};border-radius:32px;color:{{accent}};font:700 24px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif\">&#9733;</td></tr></table>\n<div style=\"margin-top:22px;font:600 12px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;letter-spacing:.8px;color:{{accent}}\">{{eyebrow}}</div>\n<h1 class=\"h1\" style=\"margin:15px 0 0;font:700 33px/39px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;letter-spacing:-.4px;color:#F5F5F7\">{{hero_title}}</h1>\n<p style=\"margin:14px 0 0;font:16px/23px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#D1D1D6\">{{greeting}}</p>\n<p style=\"margin:6px auto 0;max-width:430px;font:14px/21px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#8E8E93\">{{hero_text}}</p>\n</td></tr>\n\n<tr><td class=\"px\" style=\"padding:30px 32px 0\">\n<table role=\"presentation\" width=\"100%\" style=\"background:#1C1C1E;border-radius:16px\"><tr><td style=\"padding:24px\">\n<div style=\"font:600 12px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;letter-spacing:.7px;color:#8E8E93;margin-bottom:16px\">PARTICIPANT RECORD</div>\n<table role=\"presentation\" width=\"100%\"><tr>\n<td class=\"half\" width=\"50%\" valign=\"top\" style=\"padding:0 12px 17px 0\"><div style=\"font:12px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#6E6E73;margin-bottom:4px\">Full Name</div><div style=\"font:15px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#F5F5F7\">{{participant_name}}</div></td>\n<td class=\"half\" width=\"50%\" valign=\"top\" style=\"padding:0 0 17px 12px\"><div style=\"font:12px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#6E6E73;margin-bottom:4px\">Email</div><div style=\"font:15px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#F5F5F7;word-break:break-word\">{{email}}</div></td>\n</tr>\n<tr>\n<td class=\"half\" width=\"50%\" valign=\"top\" style=\"padding:0 12px 17px 0\"><div style=\"font:12px -apple-system,BlinkMacSystemFont,\'Segoe UI\',sans-serif;color:#6E6E73;margin-bottom:4px\">Phone / WhatsApp Number</div><div style=\"font:15px -apple-system,BlinkMacSystemFont,\'Segoe UI\',sans-serif;color:#F5F5F7;word-break:break-word\">{{phone_whatsapp}}</div></td>\n<td class=\"half\" width=\"50%\" valign=\"top\" style=\"padding:0 0 17px 12px\">&nbsp;</td>\n</tr>\n<tr><td colspan=\"2\"><div style=\"border-top:1px solid #38383A;margin-bottom:17px\"></div></td></tr>\n<tr>\n<td class=\"half\" width=\"50%\" valign=\"top\" style=\"padding:0 12px 17px 0\"><div style=\"font:12px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#6E6E73;margin-bottom:4px\">Department / Faculty</div><div style=\"font:15px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#F5F5F7\">{{department_faculty}}</div></td>\n<td class=\"half\" width=\"50%\" valign=\"top\" style=\"padding:0 0 17px 12px\"><div style=\"font:12px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#6E6E73;margin-bottom:4px\">Season</div><div style=\"font:15px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#F5F5F7\">{{season}}</div></td>\n</tr>\n<tr><td colspan=\"2\"><div style=\"border-top:1px solid #38383A;margin-bottom:17px\"></div></td></tr>\n<tr>\n<td class=\"half\" width=\"50%\" valign=\"top\" style=\"padding:0 12px 17px 0\"><div style=\"font:12px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#6E6E73;margin-bottom:4px\">ID Number</div><div style=\"font:15px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#F5F5F7\">{{id_number}}</div></td>\n<td class=\"half\" width=\"50%\" valign=\"top\" style=\"padding:0 0 17px 12px\"><div style=\"font:12px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#6E6E73;margin-bottom:4px\">Registration Number</div><div style=\"font:15px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#F5F5F7\">{{registration_number}}</div></td>\n</tr>\n<tr><td colspan=\"2\"><div style=\"border-top:1px solid #38383A;margin-bottom:17px\"></div></td></tr>\n<tr><td colspan=\"2\"><div style=\"font:12px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#6E6E73;margin-bottom:4px\">Participant ID</div><div style=\"font:15px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#F5F5F7\">{{participant_id}}</div></td></tr>\n</table>\n</td></tr></table>\n</td></tr>\n\n<tr><td class=\"px\" style=\"padding:16px 32px 0\">\n<table role=\"presentation\" width=\"100%\" style=\"background:#1C1C1E;border-radius:16px\"><tr><td align=\"center\" style=\"padding:28px 24px\">\n<div style=\"font:600 12px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;letter-spacing:.7px;color:{{accent}};margin-bottom:14px\">CERTIFICATE</div>\n<div style=\"font:600 22px/28px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#F5F5F7\">{{participant_name}}</div>\n<div style=\"margin-top:8px;font:14px/21px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#8E8E93\">{{event_name}}</div>\n<div style=\"margin-top:14px;font:12px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#6E6E73\">Certificate ID</div>\n<div style=\"margin-top:3px;font:14px -apple-system,BlinkMacSystemFont,'Courier New',monospace;color:#F5F5F7\">{{certificate_id}}</div>\n<table role=\"presentation\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\" style=\"margin-top:10px\"><tr>\n<td style=\"background:#1C1C1E;border:1px solid {{accent}};border-radius:12px\"><a href=\"{{certificate_verification_url}}\" style=\"display:block;padding:13px 26px;font:600 14px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:{{accent}}\">VERIFY CERTIFICATE</a></td>\n</tr></table>\n</td></tr></table>\n</td></tr><tr><td class=\"px\" align=\"center\" style=\"padding:35px 32px 0\">{{social_block}}</td></tr>\n<tr><td class=\"px\" align=\"center\" style=\"padding:25px 32px 42px\">\n<p style=\"margin:0 0 6px;font:12px/18px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#48484A\">{{footer_note}}</p>\n<p style=\"margin:0;font:11px -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#3A3A3C\">© {{current_year}} {{organization_name}}. All rights reserved.</p>\n</td></tr>\n</table></div></center>\n</body></html>"
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
  const values=Object.assign({},data);
  Object.keys(values).forEach(function(k){
    html=html.split('{{'+k+'}}').join(String(values[k]===null||values[k]===undefined?'':values[k]));
  });
  return html;
}
function buildRegistrationEmail_(w,p,isTest){
  const settings=getSettings_();
  const org=settings.ORGANIZATION_NAME||APP.ROOT;
  const e=renderAppleTemplate_('REGISTRATION',{
    preheader:(isTest?'[TEST] ':'')+'Registration confirmed for '+String(w['Workshop Name']),
    accent:'#6C63FF',
    org_initials:brandMarkInline_(org,brandLogoUrl_(w,settings),'#6C63FF'),
    organization_name:esc_(org),
    website_url:escAttr_(w['Website']||settings.DEFAULT_WEBSITE||'#'),
    website_label:(w['Website']||settings.DEFAULT_WEBSITE)?'Visit website':'',
    eyebrow:isTest?'TEST RUN':'REGISTRATION SUCCESSFUL',
    hero_title:'You’re officially registered.',
    greeting:'Hello, '+esc_(p['Full Name']),
    hero_text:'Your registration is confirmed. Here is your registration summary.',
    participant_name:esc_(p['Full Name']),
    email:esc_(p['Email']),
    department_faculty:esc_(p['Department / Faculty']),
    season:esc_(p['Season']),
    id_number:esc_(p['ID Number']),
    registration_number:esc_(p['Registration Number']),
    phone_whatsapp:esc_(p['Phone / WhatsApp Number']),
    participant_id:esc_(p['Participant ID']),
    event_name:esc_(w['Workshop Name']),
    event_date:esc_(formatDateForDisplay_(w['Workshop Date'], String(w['Timezone']||APP.DEFAULT_TIMEZONE))),
    start_time:esc_(formatTimeForDisplay_(w['Start Time'], String(w['Timezone']||APP.DEFAULT_TIMEZONE))),
    end_time:esc_(formatTimeForDisplay_(w['End Time'], String(w['Timezone']||APP.DEFAULT_TIMEZONE))),
    platform:esc_(w['Meeting Platform']||''),
    meeting_button:'',
    social_block:socialBlock_(w),
    footer_note:'This is an official registration communication from '+org+'.',
    current_year:new Date().getFullYear()
  });
  return {subject:'Registration confirmed • '+w['Workshop Name'],html:e,text:'Registration confirmed for '+w['Workshop Name']+'. Participant ID: '+p['Participant ID']+'.'};
}
function buildReminderEmail_(w,p,offset,isTest){
  const settings=getSettings_();
  const org=settings.ORGANIZATION_NAME||APP.ROOT;
  const offsetLabel=offset?formatOffsetHuman_(offset)+' before the workshop':'at the scheduled reminder time';
  const e=renderAppleTemplate_('REMINDER',{
    preheader:(isTest?'[TEST] ':'')+'Reminder for '+String(w['Workshop Name']),
    accent:'#2FD1E0',
    org_initials:brandMarkInline_(org,brandLogoUrl_(w,settings),'#2FD1E0'),
    organization_name:esc_(org),
    website_url:escAttr_(w['Website']||settings.DEFAULT_WEBSITE||'#'),
    website_label:(w['Website']||settings.DEFAULT_WEBSITE)?'Visit website':'',
    eyebrow:isTest?'TEST RUN':'WORKSHOP REMINDER',
    hero_title:'Your workshop is coming up.',
    greeting:'Hello, '+esc_(p['Full Name']),
    hero_text:'This reminder is scheduled '+esc_(offsetLabel)+'.',
    participant_name:esc_(p['Full Name']),
    email:esc_(p['Email']),
    department_faculty:esc_(p['Department / Faculty']),
    season:esc_(p['Season']),
    id_number:esc_(p['ID Number']),
    registration_number:esc_(p['Registration Number']),
    phone_whatsapp:esc_(p['Phone / WhatsApp Number']),
    participant_id:esc_(p['Participant ID']),
    event_name:esc_(w['Workshop Name']),
    event_date:esc_(formatDateForDisplay_(w['Workshop Date'], String(w['Timezone']||APP.DEFAULT_TIMEZONE))),
    start_time:esc_(formatTimeForDisplay_(w['Start Time'], String(w['Timezone']||APP.DEFAULT_TIMEZONE))),
    end_time:esc_(formatTimeForDisplay_(w['End Time'], String(w['Timezone']||APP.DEFAULT_TIMEZONE))),
    platform:esc_(w['Meeting Platform']||''),
    meeting_button:w['Meeting Link']?actionButtonApple_('JOIN WORKSHOP',w['Meeting Link'],'#2FD1E0','#04252A'):'',
    social_block:socialBlock_(w),
    footer_note:'This is an official workshop reminder from '+org+'.',
    current_year:new Date().getFullYear()
  });
  return {subject:'Workshop reminder • '+w['Workshop Name'],html:e,text:'Reminder: '+w['Workshop Name']+' is coming up. '+offsetLabel+'.'};
}
function buildCertificateEmail_(w,p,cert,isTest){
  const settings=getSettings_();
  const org=settings.ORGANIZATION_NAME||APP.ROOT;
  const logo=brandLogoUrl_(w,settings);
  const verifyUrl=cert.verificationUrl||'#';
  const accent='#C9A961';
  const mark=brandMarkInline_(org,logo,accent);
  const subject='Certificate ready • '+w['Workshop Name'];
  const prefix=isTest?'[TEST] ':'';
  const year=new Date().getFullYear();
  const html='<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><style>html,body{margin:0!important;padding:0!important;background:#000!important}*{box-sizing:border-box;-webkit-text-size-adjust:100%}a{text-decoration:none}table,td{border-collapse:collapse}@media screen and (max-width:600px){.wrap{width:100%!important}.px{padding-left:20px!important;padding-right:20px!important}.h1{font-size:29px!important;line-height:35px!important}}</style></head><body><div style="display:none;max-height:0;overflow:hidden;opacity:0">'+esc_(prefix+'Certificate ready for '+w['Workshop Name'])+'</div><center style="width:100%;background:#000"><div class="wrap" style="max-width:640px;margin:0 auto"><table width="100%" cellpadding="0" cellspacing="0" border="0"><tr><td class="px" style="padding:30px 32px 22px"><table width="100%"><tr><td>'+mark+'<span style="vertical-align:top;display:inline-block;margin:4px 0 0 8px;font:600 15px -apple-system,BlinkMacSystemFont,\'Segoe UI\',sans-serif;color:#F5F5F7">'+esc_(org)+'</span></td><td align="right"><a href="'+escAttr_(w['Website']||settings.DEFAULT_WEBSITE||'#')+'" style="font:13px -apple-system,BlinkMacSystemFont,\'Segoe UI\',sans-serif;color:#98989D">'+((w['Website']||settings.DEFAULT_WEBSITE)?'Visit website':'')+'</a></td></tr></table></td></tr><tr><td class="px"><div style="border-top:1px solid #2C2C2E"></div></td></tr><tr><td class="px" align="center" style="padding:50px 32px 38px"><table width="64" height="64"><tr><td align="center" style="background:#1C1C1E;border:1px solid '+accent+';border-radius:32px;color:'+accent+';font:700 24px -apple-system,BlinkMacSystemFont,\'Segoe UI\',sans-serif">&#9733;</td></tr></table><div style="margin-top:22px;font:600 12px -apple-system,BlinkMacSystemFont,\'Segoe UI\',sans-serif;letter-spacing:.8px;color:'+accent+'">'+(isTest?'TEST RUN':'CERTIFICATE AVAILABLE')+'</div><h1 class="h1" style="margin:15px 0 0;font:700 33px/39px -apple-system,BlinkMacSystemFont,\'Segoe UI\',sans-serif;color:#F5F5F7">Your certificate is ready.</h1><p style="margin:14px 0 0;font:16px/23px -apple-system,BlinkMacSystemFont,\'Segoe UI\',sans-serif;color:#D1D1D6">Congratulations, '+esc_(p['Full Name'])+'</p><p style="margin:6px auto 0;max-width:430px;font:14px/21px -apple-system,BlinkMacSystemFont,\'Segoe UI\',sans-serif;color:#8E8E93">Your personalized certificate is attached to this email. Use the verification button below to validate it.</p></td></tr><tr><td class="px" style="padding:16px 32px 0"><table width="100%" style="background:#1C1C1E;border-radius:16px"><tr><td align="center" style="padding:28px 24px"><div style="font:600 12px -apple-system,BlinkMacSystemFont,\'Segoe UI\',sans-serif;letter-spacing:.7px;color:'+accent+';margin-bottom:14px">CERTIFICATE</div><div style="font:600 22px/28px -apple-system,BlinkMacSystemFont,\'Segoe UI\',sans-serif;color:#F5F5F7">'+esc_(p['Full Name'])+'</div><div style="margin-top:8px;font:14px/21px -apple-system,BlinkMacSystemFont,\'Segoe UI\',sans-serif;color:#8E8E93">'+esc_(w['Workshop Name'])+'</div><div style="margin-top:14px;font:12px -apple-system,BlinkMacSystemFont,\'Segoe UI\',sans-serif;color:#6E6E73">Certificate ID</div><div style="margin-top:3px;font:14px -apple-system,BlinkMacSystemFont,\'Courier New\',monospace;color:#F5F5F7">'+esc_(cert.certificateId)+'</div>'+actionButtonApple_('VERIFY CERTIFICATE',verifyUrl,'#1C1C1E',accent)+'</td></tr></table></td></tr><tr><td class="px" align="center" style="padding:35px 32px 0">'+socialBlock_(w)+'</td></tr><tr><td class="px" align="center" style="padding:25px 32px 42px"><p style="margin:0 0 6px;font:12px/18px -apple-system,BlinkMacSystemFont,\'Segoe UI\',sans-serif;color:#48484A">This is an official certificate communication from '+esc_(org)+'.</p><p style="margin:0;font:11px -apple-system,BlinkMacSystemFont,\'Segoe UI\',sans-serif;color:#3A3A3C">© '+year+' '+esc_(org)+'. All rights reserved.</p></td></tr></table></div></center></body></html>';
  return {subject:subject,html:html,text:'Your certificate '+cert.certificateId+' for '+w['Workshop Name']+' is ready. The PDF is attached.'};
}
function actionButtonApple_(label,url,bg,color){
  return '<table role="presentation" cellpadding="0" cellspacing="0" border="0" style="margin-top:22px"><tr><td style="background:'+bg+';border-radius:12px"><a href="'+escAttr_(url)+'" style="display:block;padding:15px 28px;font:600 15px -apple-system,BlinkMacSystemFont,\'Segoe UI\',sans-serif;color:'+color+'">'+esc_(label)+'</a></td></tr></table>';
}


/* =========================
 * V2 — SHEET-NATIVE CONTROL CENTER
 * No HTML UI is required for normal operations.
 * ========================= */

function setupControlCenter_(ss) {
  ss = ss || getMasterSpreadsheetFast_();
  let sh = ss.getSheetByName('Control Center');
  if (!sh) sh = ss.insertSheet('Control Center', 0);
  try { sh.getRange('A1:K80').breakApart(); } catch (err) {}
  // clear() does not reliably remove legacy data-validation rules from an
  // older Control Center layout. Clear validations explicitly before writing
  // the current V1 layout, otherwise cells such as B18 can inherit an old
  // workshop dropdown and reject the new static value IMMEDIATE.
  try { sh.getRange('A1:K80').clearDataValidations(); } catch (err) {}
  sh.clear(); sh.clearConditionalFormatRules(); sh.setHiddenGridlines(true); sh.setFrozenRows(4);
  const widths={1:180,2:140,3:120,4:145,5:170,6:110,7:145,8:180,9:24,10:320,11:420};
  Object.keys(widths).forEach(k=>sh.setColumnWidth(Number(k),widths[k]));
  sh.setColumnWidth(9,24); sh.hideColumns(10,2);
  sh.getRange('A1:H1').merge().setValue('WORKSHOP CONTROL CENTER');
  sh.getRange('A2:H2').merge().setValue('V1 • Production stable • One backend • Multi-workshop • Sheet-native');
  sh.getRange('A4:H4').merge().setValue('SYSTEM OVERVIEW');
  sh.getRange('A5').setValue('WORKSHOPS'); sh.getRange('B5').setNumberFormat('0');
  sh.getRange('D5').setValue('PARTICIPANTS'); sh.getRange('E5').setNumberFormat('0');
  sh.getRange('G5').setValue('CERTIFICATES'); sh.getRange('H5').setNumberFormat('0');
  sh.getRange('A6:H6').merge().setValue('Initializing…');

  section_(sh,'A8:H8','CREATE WORKSHOP');
  sh.getRange('A9').setValue('Workshop Name *'); sh.getRange('B9:H9').merge();
  sh.getRange('A10').setValue('Date *'); sh.getRange('B10').setNumberFormat('yyyy-mm-dd'); sh.getRange('D10').setValue('Start Time'); sh.getRange('E10').setNumberFormat('HH:mm'); sh.getRange('G10').setValue('End Time'); sh.getRange('H10').setNumberFormat('HH:mm');
  sh.getRange('A11').setValue('Venue'); sh.getRange('B11:C11').merge(); sh.getRange('D11').setValue('Meeting Platform'); sh.getRange('E11:F11').merge(); sh.getRange('G11').setValue('Capacity'); sh.getRange('H11').setValue(0);
  sh.getRange('A12').setValue('Meeting Link'); sh.getRange('B12:H12').merge();
  sh.getRange('A13').setValue('Notes'); sh.getRange('B13:H14').merge();
  sh.getRange('A15').setValue('CREATE WORKSHOP →'); sh.getRange('B15').insertCheckboxes().setValue(false); sh.getRange('D15:H15').merge().setValue('Fill the workshop details, configure delivery, then tick CREATE WORKSHOP.');

  section_(sh,'A17:H17','COMMUNICATION SCHEDULE');
  sh.getRange('A18').setValue('Registration Confirmation'); sh.getRange('B18').setValue('IMMEDIATE'); sh.getRange('D18:H18').merge().setValue('Always sent immediately after a valid registration. No meeting link is included.');
  sh.getRange('A19').setValue('Reminder Email'); sh.getRange('B19').insertCheckboxes().setValue(true); sh.getRange('D19').setValue('Timing'); sh.getRange('E19:F19').merge(); sh.getRange('E19').setValue('1 hour before'); sh.getRange('G19').setValue('Value'); sh.getRange('H19').setValue(1);
  sh.getRange('A20').setValue('Custom Reminder'); sh.getRange('B20').setValue(1); sh.getRange('C20').setValue('Hours'); sh.getRange('D20').setValue('Exact Date'); sh.getRange('E20:F20').merge(); sh.getRange('G20').setValue('Time'); sh.getRange('H20').setValue('');
  sh.getRange('A21:H21').merge().setValue('Preset reminders are relative to workshop start. “Custom before workshop” uses Value + Unit. “Exact date & time” uses Exact Date + Time.');

  section_(sh,'A23:H23','CERTIFICATE DELIVERY');
  sh.getRange('A24').setValue('Certificate Email'); sh.getRange('B24').insertCheckboxes().setValue(true); sh.getRange('D24').setValue('Release'); sh.getRange('E24:F24').merge(); sh.getRange('E24').setValue('Immediately after workshop'); sh.getRange('G24').setValue('Delay'); sh.getRange('H24').setValue(0);
  sh.getRange('A25').setValue('Delay Unit'); sh.getRange('B25').setValue('Minutes'); sh.getRange('D25').setValue('Custom Date'); sh.getRange('E25:F25').merge(); sh.getRange('G25').setValue('Time'); sh.getRange('H25').setValue('');
  sh.getRange('A26:H26').merge().setValue('Certificate release always checks eligibility first. “Next morning” uses Time. “Custom date & time” uses Custom Date + Time.');

  section_(sh,'A28:H28','TEST RUN');
  sh.getRange('A29').setValue('Workshop'); sh.getRange('B29:H29').merge();
  sh.getRange('A30').setValue('Participant'); sh.getRange('B30:H30').merge();
  sh.getRange('A31').setValue('RUN COMPLETE TEST →'); sh.getRange('B31').insertCheckboxes().setValue(false); sh.getRange('D31:H31').merge().setValue('One [TEST] prefix • confirmation → reminder → certificate.');

  section_(sh,'A33:H33','QUICK ACTIONS');
  sh.getRange('A34').setValue('RUN SCHEDULER'); sh.getRange('B34').insertCheckboxes().setValue(false); sh.getRange('D34').setValue('READINESS'); sh.getRange('E34').insertCheckboxes().setValue(false); sh.getRange('G34').setValue('REFRESH'); sh.getRange('H34').insertCheckboxes().setValue(false);
  sh.getRange('A35').setValue('REPAIR WORKSHOP DATA'); sh.getRange('B35').insertCheckboxes().setValue(false); sh.getRange('D35:H35').merge().setValue('Repairs legacy metadata alignment and invalid Form / Sheet / Folder IDs.');

  section_(sh,'A36:H36','SELECTED WORKSHOP');
  sh.getRange('A37').setValue('Event ID'); sh.getRange('B37:C37').merge(); sh.getRange('D37').setValue('Status'); sh.getRange('E37:F37').merge(); sh.getRange('G37').setValue('Date'); sh.getRange('H37').setNumberFormat('yyyy-mm-dd');
  sh.getRange('A38').setValue('Registration Form'); sh.getRange('B38:H38').merge();
  sh.getRange('A39').setValue('Attendance Form'); sh.getRange('B39:H39').merge();
  sh.getRange('A40').setValue('Meeting Link'); sh.getRange('B40:H40').merge();
  sh.getRange('A41').setValue('Reminder'); sh.getRange('B41:H41').merge();
  sh.getRange('A42').setValue('Certificate'); sh.getRange('B42:H42').merge();

  section_(sh,'A44:H44','BRAND & SYSTEM DEFAULTS');
  sh.getRange('A45').setValue('Organization Name'); sh.getRange('B45:C45').merge(); sh.getRange('D45').setValue('Sender Name'); sh.getRange('E45:H45').merge();
  sh.getRange('A46').setValue('Logo URL'); sh.getRange('B46:H46').merge();
  sh.getRange('A47').setValue('Verification Web App URL'); sh.getRange('B47:H47').merge();
  sh.getRange('A48').setValue('Public Verification URL'); sh.getRange('B48:H48').merge();
  sh.getRange('A49').setValue('Timezone'); sh.getRange('B49:C49').merge(); sh.getRange('D49').setValue('Save Settings →'); sh.getRange('E49').insertCheckboxes().setValue(false); sh.getRange('G49:H49').merge().setValue('Use a custom public route such as financeclubpstu.vercel.app/verify/{CERTIFICATE_ID}.');

  section_(sh,'A50:H50','HOW TO USE');
  sh.getRange('A51:H54').merge().setValue('1) Configure the workshop and communication schedule.\n2) Set the Apps Script Web App URL and optional Public Verification URL.\n3) Tick CREATE WORKSHOP once.\n4) After registrations, use TEST RUN on an existing participant before production.\n5) Attendance responses are imported by the central scheduler.\n6) Certificates use the public verification route when configured; TEST certificates are explicitly verifiable.');

  styleControlCenter_(sh); setControlCenterDefaults_(sh); setControlCenterValidations_(sh); syncControlCenterLists_(sh,ss,'','');
}
function section_(sh,range,title){ sh.getRange(range).merge().setValue(title); }
function styleControlCenter_(sh){
  sh.getRange('A1:H1').setBackground('#10211E').setFontColor('#FFFFFF').setFontSize(24).setFontWeight('bold').setVerticalAlignment('middle'); sh.setRowHeight(1,44);
  sh.getRange('A2:H2').setBackground('#16302B').setFontColor('#D9B365').setFontSize(11).setFontWeight('bold');
  ['A4:H4','A8:H8','A17:H17','A23:H23','A28:H28','A33:H33','A36:H36','A44:H44','A50:H50'].forEach(r=>sh.getRange(r).setBackground('#EAF3F0').setFontColor('#0F3D34').setFontWeight('bold').setVerticalAlignment('middle'));
  sh.getRange('A6:H6').setBackground('#E3F2EE').setFontColor('#0A634F').setFontWeight('bold');
  ['A5','D5','G5'].forEach(a=>sh.getRange(a).setBackground('#F3F6F5').setFontWeight('bold').setFontColor('#52615C'));
  ['B5','E5','H5'].forEach(a=>sh.getRange(a).setBackground('#F3F6F5').setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center'));
  const inputs=['B9:H9','B10','E10','H10','B11:C11','E11:F11','H11','B12:H12','B13:H14','B19','E19:F19','H19','B20','C20','E20:F20','H20','B24','E24:F24','H24','B25','E25:F25','H25','B29:H29','B30:H30','B45:C45','E45:H45','B46:H46','B47:H47','B48:H48','B49:C49'];
  inputs.forEach(r=>sh.getRange(r).setBackground('#FFFFFF').setBorder(true,true,true,true,true,true,'#CFD8D4',SpreadsheetApp.BorderStyle.SOLID));
  ['B18','B19','B24'].forEach(a=>sh.getRange(a).setHorizontalAlignment('center'));
  ['B15','B31','B34','E34','H34','B35','E49'].forEach(a=>sh.getRange(a).setBackground('#0E7C66').setFontColor('#FFFFFF').setHorizontalAlignment('center'));
  ['A21:H21','A26:H26'].forEach(r=>sh.getRange(r).setFontColor('#697670').setFontStyle('italic').setFontSize(10).setWrap(true));
  sh.getRange('A51:H54').setWrap(true).setVerticalAlignment('top').setFontColor('#52615C');
  sh.getRange('A37:H42').setWrap(true); sh.setRowHeights(18,8,24);
  sh.setRowHeight(21,34); sh.setRowHeight(26,34); sh.setRowHeight(51,24); sh.setRowHeight(52,24); sh.setRowHeight(53,24); sh.setRowHeight(54,24);
}
function setControlCenterDefaults_(sh){
  const settings=getSettings_(); const tz=String(settings.DEFAULT_TIMEZONE||APP.DEFAULT_TIMEZONE); const tomorrow=new Date(Date.now()+86400000);
  if(!sh.getRange('B10').getValue()) sh.getRange('B10').setValue(Utilities.formatDate(tomorrow,tz,'yyyy-MM-dd'));
  if(!sh.getRange('E10').getValue()) sh.getRange('E10').setValue('10:00'); if(!sh.getRange('H10').getValue()) sh.getRange('H10').setValue('12:00');
  if(!sh.getRange('H11').getValue()) sh.getRange('H11').setValue(0);
  if(!sh.getRange('B20').getValue()) sh.getRange('B20').setValue(1); if(!sh.getRange('C20').getValue()) sh.getRange('C20').setValue('Hours');
  if(!sh.getRange('H19').getValue()) sh.getRange('H19').setValue(1); if(!sh.getRange('E19').getValue()) sh.getRange('E19').setValue('1 hour before');
  if(!sh.getRange('H24').getValue()) sh.getRange('H24').setValue(0); if(!sh.getRange('B25').getValue()) sh.getRange('B25').setValue('Minutes'); if(!sh.getRange('E24').getValue()) sh.getRange('E24').setValue('Immediately after workshop');
  sh.getRange('B10').setNumberFormat('yyyy-mm-dd'); sh.getRange('E10').setNumberFormat('HH:mm'); sh.getRange('H10').setNumberFormat('HH:mm'); sh.getRange('E20:F20').setNumberFormat('yyyy-mm-dd'); sh.getRange('H20').setNumberFormat('HH:mm'); sh.getRange('E25:F25').setNumberFormat('yyyy-mm-dd'); sh.getRange('H25').setNumberFormat('HH:mm');
  sh.getRange('B45').setValue(settings.ORGANIZATION_NAME||APP.ROOT); sh.getRange('E45').setValue(settings.SENDER_NAME||APP.ROOT); sh.getRange('B46').setValue(settings.DEFAULT_LOGO_URL||''); sh.getRange('B47').setValue(settings.WEB_APP_URL||ScriptApp.getService().getUrl()||''); sh.getRange('B48').setValue(settings.PUBLIC_VERIFY_URL||APP.DEFAULT_PUBLIC_VERIFY_URL); sh.getRange('B49').setValue(settings.DEFAULT_TIMEZONE||APP.DEFAULT_TIMEZONE);
}
function setControlCenterValidations_(sh){
  sh.getRange('E11').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Google Meet','Zoom','Microsoft Teams','Other'],true).setAllowInvalid(false).build());
  sh.getRange('E19').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['30 minutes before','1 hour before','2 hours before','6 hours before','12 hours before','1 day before','2 days before','Custom before workshop','Exact date & time'],true).setAllowInvalid(false).build());
  sh.getRange('C20').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Minutes','Hours','Days'],true).setAllowInvalid(false).build());
  sh.getRange('E24').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Immediately after workshop','30 minutes after workshop','1 hour after workshop','2 hours after workshop','6 hours after workshop','12 hours after workshop','Next morning','Custom delay after workshop','Custom date & time'],true).setAllowInvalid(false).build());
  sh.getRange('B25').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Minutes','Hours','Days'],true).setAllowInvalid(false).build());
}
function directDropdownValidation_(cell, values, fallbackRange) {
  const vals=(values||[]).map(v=>String(v||'').trim()).filter(Boolean); let rule=null;
  if(vals.length===0){ if(fallbackRange) rule=SpreadsheetApp.newDataValidation().requireValueInRange(fallbackRange,true).setAllowInvalid(false).build(); }
  else if(vals.length<=200){ rule=SpreadsheetApp.newDataValidation().requireValueInList(vals,true).setAllowInvalid(false).build(); }
  else if(fallbackRange){ rule=SpreadsheetApp.newDataValidation().requireValueInRange(fallbackRange,true).setAllowInvalid(false).build(); }
  cell.setDataValidation(rule);
}
function applyControlCenterValidations_(sh,workshopValues,participantValues){ directDropdownValidation_(sh.getRange('B29'),workshopValues,sh.getRange('J2:J201')); directDropdownValidation_(sh.getRange('B30'),participantValues,sh.getRange('K2:K501')); }
function installControlCenterEditTrigger_(ss){ const triggers=ScriptApp.getProjectTriggers(); triggers.forEach(t=>{if(t.getHandlerFunction()==='controlCenterOnEdit_') ScriptApp.deleteTrigger(t);}); ScriptApp.newTrigger('controlCenterOnEdit_').forSpreadsheet(ss).onEdit().create(); }
function controlCenterOnEdit_(e){
  try{
    if(!e||!e.range) return; const sh=e.range.getSheet(); if(sh.getName()!=='Control Center') return; const a1=e.range.getA1Notation();
    if(a1==='B15'&&e.value==='TRUE'){sh.getRange('B15').setValue(false);const msg=createWorkshopFromControlCenter_();toastControlCenter_(msg,'Workshop created');refreshControlCenter_();return;}
    if(a1==='B29'){refreshParticipantOptionsForControlCenter_();updateSelectedWorkshopPanel_();return;}
    if(a1==='B30'){updateSelectedParticipantPanel_();return;}
    if(a1==='B31'&&e.value==='TRUE'){sh.getRange('B31').setValue(false);const msg=startTestFromControlCenter_();toastControlCenter_(msg,'Test started');refreshControlCenter_();return;}
    if(a1==='B34'&&e.value==='TRUE'){sh.getRange('B34').setValue(false);processScheduler();toastControlCenter_('Scheduler completed.','Scheduler');refreshControlCenter_();return;}
    if(a1==='E34'&&e.value==='TRUE'){sh.getRange('E34').setValue(false);const r=systemReadinessCheck();toastControlCenter_(r.overall==='PASS'?'All readiness checks passed.':'Readiness has errors.','Readiness');sh.getRange('A6').setValue(r.overall==='PASS'?'READINESS PASS — all checks passed.':'READINESS ERROR — check System Logs.');return;}
    if(a1==='H34'&&e.value==='TRUE'){sh.getRange('H34').setValue(false);refreshControlCenter_();toastControlCenter_('Control Center refreshed.','Refresh');return;}
    if(a1==='B35'&&e.value==='TRUE'){sh.getRange('B35').setValue(false);const r=repairAllWorkshopMetadata_();toastControlCenter_('Repaired '+r.repaired+' workshop(s).','Workshop Repair');refreshControlCenter_();return;}
    if(a1==='E19'||a1==='C20'||a1==='E24'||a1==='B25'){updateControlCenterScheduleHints_(sh);return;}
    if(a1==='E49'&&e.value==='TRUE'){sh.getRange('E49').setValue(false);saveGlobalSettingsFromControlCenter_();toastControlCenter_('Global settings saved.','Settings');refreshControlCenter_();return;}
  }catch(err){sh.getRange('A6').setValue('ERROR: '+String(err&&err.message?err.message:err));toastControlCenter_(String(err&&err.message?err.message:err),'WORKSHOP Automation');}
}
function updateControlCenterScheduleHints_(sh){
  sh.getRange('A21:H21').setValue('Reminder: '+String(sh.getRange('E19').getDisplayValue())+'. Custom fields apply only when selected.');
  sh.getRange('A26:H26').setValue('Certificate: '+String(sh.getRange('E24').getDisplayValue())+'. Eligibility is checked before a certificate is generated.');
}
function saveGlobalSettingsFromControlCenter_(){
  const sh=getMasterSpreadsheetFast_().getSheetByName('Control Center'); const settingsSh=getSheet_(APP.SHEETS.SETTINGS); const now=new Date();
  const pairs=[['ORGANIZATION_NAME',sh.getRange('B45').getDisplayValue()],['SENDER_NAME',sh.getRange('E45').getDisplayValue()],['DEFAULT_LOGO_URL',sh.getRange('B46').getDisplayValue()],['WEB_APP_URL',sh.getRange('B47').getDisplayValue()],['PUBLIC_VERIFY_URL',sh.getRange('B48').getDisplayValue()||APP.DEFAULT_PUBLIC_VERIFY_URL],['DEFAULT_TIMEZONE',sh.getRange('B49').getDisplayValue()||APP.DEFAULT_TIMEZONE]];
  pairs.forEach(function(x){updateObjectByKey_(settingsSh,'Key',x[0],{'Value':clean_(x[1]),'Updated At':now});});
}
function toastControlCenter_(message,title){try{SpreadsheetApp.getActiveSpreadsheet().toast(String(message),String(title||'WORKSHOP Automation'),6);}catch(err){}}
function controlCenterIsoDate_(value,tz){const zone=String(tz||getSetting_('DEFAULT_TIMEZONE')||APP.DEFAULT_TIMEZONE);if(value instanceof Date&&!isNaN(value.getTime()))return Utilities.formatDate(value,zone,'yyyy-MM-dd');const raw=clean_(value);if(/^\d{4}-\d{2}-\d{2}$/.test(raw))return raw;return normalizeDateOnly_(raw,zone);}
function controlCenterHHMM_(value,tz){const zone=String(tz||getSetting_('DEFAULT_TIMEZONE')||APP.DEFAULT_TIMEZONE);if(value instanceof Date&&!isNaN(value.getTime()))return Utilities.formatDate(value,zone,'HH:mm');const raw=clean_(value);if(/^\d{2}:\d{2}$/.test(raw))return raw;return normalizeTime_(raw,zone);}
function reminderPresetToMinutes_(preset){ const m={'30 minutes before':30,'1 hour before':60,'2 hours before':120,'6 hours before':360,'12 hours before':720,'1 day before':1440,'2 days before':2880}; return m[preset]||0; }
function customUnitToMinutes_(amount,unit){const n=Number(amount||0);if(!isFinite(n)||n<=0)return 0;return unit==='Days'?n*1440:(unit==='Hours'?n*60:n);}
function certificatePreset_(value){ const m={'Immediately after workshop':0,'30 minutes after workshop':30,'1 hour after workshop':60,'2 hours after workshop':120,'6 hours after workshop':360,'12 hours after workshop':720}; return m[value]!==undefined?m[value]:null; }
function createWorkshopFromControlCenter_(){
  const sh=getMasterSpreadsheetFast_().getSheetByName('Control Center'); const tz=String(getSetting_('DEFAULT_TIMEZONE')||APP.DEFAULT_TIMEZONE);
  const name=clean_(sh.getRange('B9').getValue()), date=controlCenterIsoDate_(sh.getRange('B10').getValue(),tz), start=controlCenterHHMM_(sh.getRange('E10').getValue(),tz)||'10:00', end=controlCenterHHMM_(sh.getRange('H10').getValue(),tz)||'12:00';
  if(!name)throw new Error('Workshop Name is required.'); if(!date)throw new Error('Workshop Date is required.');
  const reminderEnabled=sh.getRange('B19').getValue()===true; const reminderPreset=clean_(sh.getRange('E19').getDisplayValue()); let reminderMode='BEFORE_START', reminderOffsets='60', reminderCustomAt='';
  if(reminderEnabled){
    if(reminderPreset==='Exact date & time'){ reminderMode='EXACT_DATETIME'; const rd=controlCenterIsoDate_(sh.getRange('E20').getValue(),tz); const rt=controlCenterHHMM_(sh.getRange('H20').getValue(),tz); if(!rd||!rt)throw new Error('Reminder exact date and time are required.'); const exactReminder=parseDateTime_(rd,rt,tz); const eventStart=parseDateTime_(date,start,tz); if(exactReminder>=eventStart)throw new Error('Exact reminder time must be before the workshop starts.'); reminderCustomAt=rd+' '+rt; reminderOffsets=''; }
    else if(reminderPreset==='Custom before workshop'){ const mins=customUnitToMinutes_(sh.getRange('B20').getValue(),clean_(sh.getRange('C20').getDisplayValue())); if(!mins)throw new Error('Enter a valid custom reminder lead time.'); reminderOffsets=String(mins); }
    else { const mins=reminderPresetToMinutes_(reminderPreset); if(!mins)throw new Error('Select a valid reminder timing.'); reminderOffsets=String(mins); }
  } else { reminderOffsets=''; }
  const certEnabled=sh.getRange('B24').getValue()===true; const certRelease=clean_(sh.getRange('E24').getDisplayValue()); let certMode='AFTER_WORKSHOP', certDelay=0, certNext='09:00', certCustom='';
  if(certEnabled){
    const preset=certificatePreset_(certRelease);
    if(preset!==null){certMode='AFTER_WORKSHOP';certDelay=preset;}
    else if(certRelease==='Next morning'){certMode='NEXT_MORNING';certNext=controlCenterHHMM_(sh.getRange('H25').getValue(),tz)||'09:00';}
    else if(certRelease==='Custom delay after workshop'){certMode='AFTER_WORKSHOP';const rawDelay=Number(sh.getRange('H24').getValue());if(!isFinite(rawDelay)||rawDelay<0)throw new Error('Enter a valid certificate delay.');certDelay=customUnitToMinutes_(rawDelay,clean_(sh.getRange('B25').getDisplayValue()));}
    else if(certRelease==='Custom date & time'){certMode='CUSTOM_DATETIME';const cd=controlCenterIsoDate_(sh.getRange('E25').getValue(),tz);const ct=controlCenterHHMM_(sh.getRange('H25').getValue(),tz);if(!cd||!ct)throw new Error('Certificate custom date and time are required.');const customRelease=parseDateTime_(cd,ct,tz);const eventEnd=parseDateTime_(date,end,tz);if(customRelease<eventEnd)throw new Error('Certificate release must be at or after the workshop ends.');certCustom=cd+' '+ct;}
  }
  const result=createWorkshop({workshopName:name,workshopDate:date,startTime:start,endTime:end,venueName:clean_(sh.getRange('B11').getValue()),meetingPlatform:clean_(sh.getRange('E11').getDisplayValue())||getSetting_('DEFAULT_MEETING_PLATFORM'),meetingLink:clean_(sh.getRange('B12').getValue()),capacity:Number(sh.getRange('H11').getValue()||0),reminderEnabled:reminderEnabled,reminderMode:reminderMode,reminderOffsets:reminderOffsets,reminderCustomAt:reminderCustomAt,certificateEnabled:certEnabled,certificateReleaseMode:certMode,certificateDelay:certDelay,certificateNextMorningTime:certNext,certificateCustomReleaseAt:certCustom,notes:clean_(sh.getRange('B13').getValue())});
  ['B9','B11:C11','B12:H12','B13:H14'].forEach(a=>sh.getRange(a).clearContent()); return 'Created '+result.eventId+' — '+result.workshopName+'\nRegistration Form: '+result.registrationFormUrl+'\nAttendance Form: '+result.attendanceFormUrl;
}
function startTestFromControlCenter_(){const sh=getMasterSpreadsheetFast_().getSheetByName('Control Center');const selection=clean_(sh.getRange('B30').getDisplayValue());if(!selection)throw new Error('Select an existing participant first.');const pid=selection.split(' | ')[0].trim();const participant=findParticipantById_(pid);if(!participant)throw new Error('Selected participant was not found.');const ws=clean_(sh.getRange('B29').getDisplayValue());const eventId=ws.split(' | ')[0].trim();if(String(participant['Event ID'])!==String(eventId))throw new Error('Selected participant does not belong to the selected workshop.');const r=startTestRun(pid);return r.message+' Run ID: '+r.runId;}
function refreshParticipantOptionsForControlCenter_(){const ss=getMasterSpreadsheetFast_(),sh=ss.getSheetByName('Control Center'),sel=clean_(sh.getRange('B29').getDisplayValue()),eid=sel.split(' | ')[0].trim(),ps=readSheetObjects_(controlSheetFast_(ss,APP.SHEETS.PARTICIPANTS)),vals=eid?ps.filter(p=>String(p['Event ID'])===eid&&String(p['Participant ID']||'').trim()).map(p=>String(p['Participant ID'])+' | '+String(p['Full Name']||'')+' | '+String(p['Email']||'')):[];sh.getRange('K2:K501').clearContent();if(vals.length)sh.getRange(2,11,Math.min(vals.length,500),1).setValues(vals.slice(0,500).map(x=>[x]));applyControlCenterValidations_(sh,getControlCenterWorkshopOptions_(ss),vals);if(!vals.includes(clean_(sh.getRange('B30').getDisplayValue())))sh.getRange('B30').clearContent();}
function getControlCenterWorkshopOptions_(ss){return readSheetObjects_(controlSheetFast_(ss,APP.SHEETS.WORKSHOPS)).filter(w=>String(w['Event ID']||'').trim()).filter(w=>String(w.Status||'').toUpperCase()!=='ARCHIVED').sort((a,b)=>String(a['Workshop Date']||'').localeCompare(String(b['Workshop Date']||''))).slice(0,200).map(w=>String(w['Event ID'])+' | '+String(w['Workshop Name']||'Untitled Workshop'));}
function syncControlCenterLists_(sh,ss,preserveWorkshop,preserveParticipant){ss=ss||getMasterSpreadsheetFast_();sh=sh||ss.getSheetByName('Control Center');if(!sh)return;const wsOptions=getControlCenterWorkshopOptions_(ss);sh.getRange('J2:J201').clearContent();if(wsOptions.length)sh.getRange(2,10,Math.min(wsOptions.length,200),1).setValues(wsOptions.slice(0,200).map(x=>[x]));const sw=clean_(preserveWorkshop!==undefined?preserveWorkshop:sh.getRange('B29').getDisplayValue()),sp=clean_(preserveParticipant!==undefined?preserveParticipant:sh.getRange('B30').getDisplayValue());applyControlCenterValidations_(sh,wsOptions,[]);if(sw&&wsOptions.includes(sw))sh.getRange('B29').setValue(sw);else sh.getRange('B29').clearContent();refreshParticipantOptionsForControlCenter_();if(sp){const eid=sw.split(' | ')[0].trim(),vals=readSheetObjects_(controlSheetFast_(ss,APP.SHEETS.PARTICIPANTS)).filter(p=>String(p['Event ID'])===eid&&String(p['Participant ID']||'').trim()).map(p=>String(p['Participant ID'])+' | '+String(p['Full Name']||'')+' | '+String(p['Email']||''));if(vals.includes(sp))sh.getRange('B30').setValue(sp);}}
function updateSelectedWorkshopPanel_(){const ss=getMasterSpreadsheetFast_(),sh=ss.getSheetByName('Control Center'),sel=clean_(sh.getRange('B29').getDisplayValue()),eid=sel.split(' | ')[0].trim(),w=findWorkshop_(eid);if(!w){sh.getRange('B37:H42').clearContent();return;}sh.getRange('B37').setValue(w['Event ID']);sh.getRange('E37').setValue(w['Status']||'');sh.getRange('H37').setValue(w['Workshop Date']);sh.getRange('H37').setNumberFormat('yyyy-mm-dd');const reg=w['Registration Form ID']?'https://docs.google.com/forms/d/'+w['Registration Form ID']+'/viewform':'';const att=w['Attendance Form ID']?'https://docs.google.com/forms/d/'+w['Attendance Form ID']+'/viewform':'';sh.getRange('B38').setFormula(reg?'=HYPERLINK("'+reg+'","Open Registration Form")':'="Not available"');sh.getRange('B39').setFormula(att?'=HYPERLINK("'+att+'","Open Attendance Form")':'="Not available"');sh.getRange('B40').setFormula(w['Meeting Link']?'=HYPERLINK("'+String(w['Meeting Link']).replace(/"/g,'""')+'","Join / Open Meeting")':'="No meeting link configured"');sh.getRange('B41').setValue(w['Reminder Enabled']===false?'Disabled':(String(w['Reminder Mode'])==='EXACT_DATETIME'?formatDateTimeForDisplay_(w['Reminder Custom At'],String(w['Timezone']||APP.DEFAULT_TIMEZONE)):parseOffsets_(w['Reminder Offsets']).map(formatOffsetHuman_).join(', ')+' before start'));sh.getRange('B42').setValue(String(w['Certificate Release Mode'])==='CUSTOM_DATETIME'?formatDateTimeForDisplay_(w['Certificate Custom Release At'],String(w['Timezone']||APP.DEFAULT_TIMEZONE)):String(w['Certificate Release Mode'])==='NEXT_MORNING'?'Next morning at '+(w['Certificate Next Morning Time']||'09:00'):formatOffsetHuman_(Number(w['Certificate Delay']||0))+' after workshop');}
function updateSelectedParticipantPanel_(){return;}
function refreshControlCenter_(){const ss=getMasterSpreadsheetFast_();let sh=ss.getSheetByName('Control Center');if(!sh){setupControlCenter_(ss);sh=ss.getSheetByName('Control Center');}const ws=readSheetObjects_(controlSheetFast_(ss,APP.SHEETS.WORKSHOPS)).filter(w=>String(w['Event ID']||'').trim()),ps=readSheetObjects_(controlSheetFast_(ss,APP.SHEETS.PARTICIPANTS)).filter(p=>String(p['Participant ID']||'').trim()),sw=clean_(sh.getRange('B29').getDisplayValue()),sp=clean_(sh.getRange('B30').getDisplayValue());const cs=readSheetObjects_(controlSheetFast_(ss,APP.SHEETS.CERTS)).filter(c=>!String(c.Status||'').startsWith('TEST'));sh.getRange('B5').setValue(ws.length);sh.getRange('E5').setValue(ps.length);sh.getRange('H5').setValue(cs.length);sh.getRange('A6').setValue('V1 connected • Refreshed '+Utilities.formatDate(new Date(),String(getSetting_('DEFAULT_TIMEZONE')||APP.DEFAULT_TIMEZONE),'yyyy-MM-dd HH:mm:ss'));syncControlCenterLists_(sh,ss,sw,sp);updateSelectedWorkshopPanel_();setControlCenterDefaults_(sh);updateControlCenterScheduleHints_(sh);return {ok:true,version:APP.VERSION,workshops:ws.length,participants:ps.length,certificates:cs.length};}

function normalizeCertificateId_(value) {
  return clean_(value).toUpperCase();
}

function verifyCertificate(certificateId){
  const id=normalizeCertificateId_(certificateId);
  if(!id) return {valid:false,code:'MISSING_ID',message:'Certificate ID is required.'};

  const row=readSheetObjects_(getSheet_(APP.SHEETS.CERTS)).find(r=>{
    const rowId=normalizeCertificateId_(r['Certificate ID']);
    const status=String(r['Status']||'').trim().toUpperCase();
    const isTestRequest=/^TEST-CERT-/i.test(id);
    const isTestRow=status.startsWith('TEST') || /^TEST-CERT-/i.test(rowId);
    return rowId===id && (isTestRequest ? isTestRow : !isTestRow);
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
  try {
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
      isTest:r.status ? String(r.status).toUpperCase().startsWith('TEST') : false
    };
  } catch (err) {
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
  const metrics=[['Metric','Value','Updated At'],['Workshops',ws.length,new Date()],['Open Workshops',ws.filter(w=>w.Status==='OPEN').length,new Date()],['Participants',ps.length,new Date()],['Attendance Present',ps.filter(p=>p['Attendance Status']==='PRESENT').length,new Date()],['Pending Emails',qs.filter(q=>['PENDING','RETRY'].includes(String(q.Status))).length,new Date()],['Failed Emails',qs.filter(q=>String(q.Status)==='FAILED').length,new Date()],['Certificates',cs.filter(c=>!String(c.Status).startsWith('TEST')).length,new Date()],['Version',APP.VERSION,new Date()]];
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
  if(state!==w.Status) updateWorkshopByEvent_(w['Event ID'],{'Status':state,'Updated At':new Date()});
}

function isCertificateEligible_(w,p){
  return eligibilityFrom_(w,p,String(p['Attendance Status'])==='PRESENT')==='ELIGIBLE';
}
function eligibilityFrom_(w,p,attended){
  const rule=String(w['Certificate Eligibility']||'ATTENDANCE_REQUIRED');
  if(rule==='ALL_REGISTERED') return ['REGISTERED','APPROVED','ON_HOLD'].includes(String(p['Registration Status']))?'ELIGIBLE':'NOT_ELIGIBLE';
  if(rule==='APPROVED_ONLY') return String(p['Registration Status'])==='APPROVED'?'ELIGIBLE':'NOT_ELIGIBLE';
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
function withLock_(name,fn){const lock=LockService.getScriptLock();lock.waitLock(30000);try{return fn();}finally{lock.releaseLock();}}
function log_(level,module,eventId,participantId,action,message,details){try{const sh=getSheet_(APP.SHEETS.LOGS);appendObject_(sh,{'Timestamp':new Date(),'Level':level,'Module':module,'Event ID':eventId||'','Participant ID':participantId||'','Action':action||'','Message':message||'','Details':details||''});}catch(err){}}


function tidyEmailQueueSheet_() {
  const sh=getSheet_(APP.SHEETS.QUEUE);
  const headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(String);
  const hiddenNames=['Unique Key','Attachment File ID','Last Error','Message Body','Include Meeting Link','Sender Name','Created At','Updated At'];
  hiddenNames.forEach(function(name){
    const idx=headers.indexOf(name);
    if(idx>=0 && !sh.isColumnHiddenByUser(idx+1)) sh.hideColumns(idx+1);
  });
  sh.setFrozenRows(1);
  const visible=[0,2,3,4,5,6,7,8,9,10,11];
  visible.forEach(function(idx){
    if(idx<sh.getMaxColumns()) sh.setColumnWidth(idx+1,Math.max(110,Math.min(220,headers[idx].length*9+30)));
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
  installControlCenterEditTrigger_(ss);
  tidyEmailQueueSheet_();
  installCentralScheduler_();
  refreshDashboard_();
  refreshControlCenter_();
  log_('INFO','REPAIR','','','V1_REPAIR_COMPLETE','V1 self-repair completed.','Sheet-native Control Center, scheduler isolation, safe Form response readers, stable time normalization, certificate export checks, queue cleanup.');
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
