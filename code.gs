/***************************************************************
 * BETTEREDGE SOLUTIONS INC.
 * WORKFORCE MANAGEMENT PORTAL - V4.2
 ***************************************************************/

const APP = {
  NAME: 'BetterEdge Workforce Management System',
  COMPANY: 'BetterEdge Solutions Inc.',
  VERSION: 'V4.2-WORKFORCE',
  TIMEZONE: 'Asia/Manila',
  SHEETS: {
    USERS:'USERS',
    EMPLOYEES:'EMPLOYEES',
    SITES:'SITES',
    CLIENTS:'CLIENTS',
    DOCUMENTS:'DOCUMENTS',
    SITE_REPORTS:'SITE_REPORTS',
    HELP_TICKETS:'HELP_TICKETS',
    AUDIT_LOG:'AUDIT_LOG',
    SETTINGS:'SETTINGS'
  }
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(APP.NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function setupBesiWorkforce() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Open this Apps Script project from a Google Spreadsheet.');

  const defs = getDefinitions_();
  // An existing V4.1 sheet has a different column layout. Stop before writing.
  Object.keys(defs).forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh && sh.getLastRow()) {
      const actual = sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0];
      if (actual.join('\u001f') !== defs[name].join('\u001f'))
        throw new Error('Existing '+name+' layout differs. Use a new spreadsheet for this clean build; preserve your original spreadsheet.');
    }
  });

  Object.keys(defs).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) writeHeader_(sh, defs[name]);
  });

  if (getSheet_(APP.SHEETS.USERS).getLastRow() < 2) createDefaultAdmin_();
  if (getSheet_(APP.SHEETS.SETTINGS).getLastRow() < 2) createDefaultSettings_();


  return {success:true,message:'BESI workforce portal initialized.'};
}



function writeHeader_(sh, headers) {
  sh.getRange(1,1,1,headers.length).setValues([headers]);
  sh.getRange(1,1,1,headers.length)
    .setFontWeight('bold')
    .setBackground('#58BDEB')
    .setFontColor('#FFFFFF');
  sh.setFrozenRows(1);
}

function getDefinitions_() {
  const S = APP.SHEETS;
  const d = {};

  d[S.USERS] = [
    'User ID','Username','Password Hash','Full Name','Role',
    'Department','Assigned Site','Status','Last Login'
  ];

  d[S.EMPLOYEES] = [
    'Employee ID','Last Name','First Name','Middle Name','Full Name',
    'Birthday','Gender','Contact Number','Email','Address',
    'Emergency Contact','Emergency Number','Date Hired',
    'Client','Site','Department','Position','Employment Type',
    'Status','Last Updated'
  ];

  d[S.SITES] = [
    'Site ID','Site Name','Client','Location','Region','Required HC',
    'Coordinator','Team Leader','Required Documents','Status','Notes'
  ];

  d[S.CLIENTS] = [
    'Client ID','Client Name','Contact Person','Contact Number',
    'Email','Address','Contract Start','Contract End','Status'
  ];

  d[S.DOCUMENTS] = [
    'Document ID','Employee ID','Employee Name','Site','Document Type',
    'File Name','Drive File ID','Drive URL','Expiration Date',
    'Status','Remarks','Uploaded By','Uploaded Date'
  ];

  d[S.SITE_REPORTS] = [
    'Report ID','Site','Department','Reported By','Role','Category',
    'Subject','Description','Employee Involved','Priority',
    'Status','Management Response','Reported Date','Last Updated'
  ];

  d[S.HELP_TICKETS] = [
    'Ticket ID','Site','Department','Reported By','Role','Category',
    'Subject','Description','Priority','Status','Management Response',
    'Created Date','Last Updated'
  ];

  d[S.AUDIT_LOG] = [
    'Timestamp','User','Role','Department','Module','Action',
    'Record ID','Old Value','New Value'
  ];

  d[S.SETTINGS] = ['Setting','Value','Description'];

  return d;
}

function createDefaultAdmin_() {
  const sh = getSheet_(APP.SHEETS.USERS);
  if (getUsersRaw_().some(u => u.username.toLowerCase() === 'admin')) return;

  sh.appendRow([
    'ADM-0001',
    'admin',
    hashPassword_('BetterEdge2026!'),
    'Head Administrator',
    'HEAD_ADMIN',
    'MANAGEMENT',
    'ALL',
    'ACTIVE',
    ''
  ]);
}

function createDefaultSettings_() {
  const sh = getSheet_(APP.SHEETS.SETTINGS);
  if (sh.getLastRow() > 1) return;

  sh.getRange(2,1,4,3).setValues([
    ['COMPANY_NAME',APP.COMPANY,'Company name'],
    ['TIMEZONE',APP.TIMEZONE,'System timezone'],
    ['DOCUMENT_ROOT_FOLDER','BetterEdge Employee Documents','Root Drive folder'],
    ['SITE_REPORT_PRIVACY','STRICT','Reports are private to the selected/assigned site']
  ]);
}





/***************************************************************
 * LOGIN / ACCESS
 ***************************************************************/

function loginUser(username,password) {
  username = clean_(username).toLowerCase();
  password = String(password || '');

  if (!username || !password) {
    return {success:false,message:'Enter username and password.'};
  }

  const hash = hashPassword_(password);

  const user = getUsersRaw_().find(
    u => u.username.toLowerCase() === username && u.passwordHash === hash
  );

  if (!user) return {success:false,message:'Invalid username or password.'};
  if (user.status !== 'ACTIVE') return {success:false,message:'This account is inactive.'};
  if (!['HEAD_ADMIN','HEAD_COORDINATOR','HR_HEAD','HR','COORDINATOR','TL','VIEWER'].includes(user.role)) return {success:false,message:'This role is no longer available.'};

  const row = findRowById_(APP.SHEETS.USERS,1,user.userId);
  getSheet_(APP.SHEETS.USERS).getRange(row,9).setValue(new Date());

  const session = {
    userId:user.userId,
    username:user.username,
    fullName:user.fullName,
    role:user.role,
    department:user.department,
    site:user.site
  };

  const token = Utilities.getUuid();

  CacheService.getScriptCache().put(
    'SESSION_' + token,
    JSON.stringify(session),
    21600
  );

  logAudit_(session,'AUTH','LOGIN',session.userId,'','Successful login');

  return {
    success:true,
    token:token,
    user:session,
    permissions:getPermissions_(session)
  };
}

function logoutUser(token) {
  if (token) CacheService.getScriptCache().remove('SESSION_' + token);
  return {success:true};
}

function getSession_(token) {
  if (!token) throw new Error('Session expired. Please log in again.');

  const raw = CacheService.getScriptCache().get('SESSION_' + token);
  if (!raw) throw new Error('Session expired. Please log in again.');

  return JSON.parse(raw);
}

function getPermissions_(user) {
  const role = clean_(user.role).toUpperCase();

  const p = {
    dashboard:true,
    sites:false,
    siteManage:false,
    employees:false,
    employeeEdit:false,
    employeeBulk:false,
    documents:false,
    documentUpload:false,
    reports:true,
    help:true,
    users:false,
    audit:false,
    showReportsIssues:false,
    showHelpIssues:false,
    showDocumentIssues:false
  };

  if (role === 'HEAD_ADMIN') {
    Object.keys(p).forEach(k => p[k] = true);
    return p;
  }

  if (role === 'HR_HEAD') {
    Object.assign(p,{
      sites:true,siteManage:true,employees:true,employeeEdit:true,employeeBulk:true,
      documents:true,documentUpload:true,showDocumentIssues:true,
      showReportsIssues:true,showHelpIssues:true
    });
  }

  if (role === 'HR') {
    Object.assign(p,{
      sites:true,employees:true,employeeEdit:true,employeeBulk:true,
      documents:true,documentUpload:true,showDocumentIssues:true
    });
  }

  if (role === 'HEAD_COORDINATOR') {
    Object.assign(p,{sites:true,employees:true,documents:true,reports:true,help:true,showReportsIssues:true,showHelpIssues:true,showDocumentIssues:true});
  }

  if (role === 'COORDINATOR') {
    Object.assign(p,{
      sites:true,employees:true,employeeEdit:true,employeeBulk:true,
      documents:true,documentUpload:true,reports:true,help:true
    });
  }

  if (role === 'TL') {
    Object.assign(p,{sites:true,employees:true,reports:true,help:true});
  }

  if (role === 'VIEWER') {
    Object.assign(p,{sites:true,employees:true});
  }

  return p;
}

function assertRole_(user,roles) {
  if (!roles.includes(clean_(user.role).toUpperCase())) {
    throw new Error('You do not have permission for this action.');
  }
}

function canViewSite_(user,site) {
  if (['HEAD_ADMIN','HEAD_COORDINATOR'].includes(clean_(user.role).toUpperCase())) return true;

  const assigned = clean_(user.site);
  if (!assigned || assigned.toUpperCase() === 'ALL') return true;

  return assigned.toUpperCase() === clean_(site).toUpperCase();
}

function strictSite_(user,requestedSite) {
  const assigned = clean_(user.site);

  // Head/global accounts may select any specific site. When no site is supplied,
  // an empty string means "all accessible sites" for read-only functions.
  if (['HEAD_ADMIN','HEAD_COORDINATOR'].includes(clean_(user.role).toUpperCase()) || assigned.toUpperCase() === 'ALL') {
    const requested = clean_(requestedSite || '');
    return requested.toUpperCase() === 'ALL' ? '' : requested;
  }

  if (!assigned) {
    throw new Error('This account does not have a site assignment.');
  }

  return assigned;
}

/***************************************************************
 * DASHBOARD
 ***************************************************************/

function getDashboardData(token) {
  const user = getSession_(token);
  const permissions = getPermissions_(user);

  const employees = getEmployeesRaw_().filter(e => canViewSite_(user,e.site));
  const sites = getSitesRaw_().filter(s => canViewSite_(user,s.siteName));

  const counts = {
    employees:employees.length,
    activeEmployees:employees.filter(e => e.status === 'ACTIVE').length,
    sites:sites.length
  };



  if (permissions.showReportsIssues) {
    counts.openReports = getReportsRaw_()
      .filter(r => canViewSite_(user,r.site))
      .filter(r => !['RESOLVED','CLOSED'].includes(r.status))
      .length;
  }

  if (permissions.showHelpIssues) {
    counts.openHelpTickets = getHelpRaw_()
      .filter(r => canViewSite_(user,r.site))
      .filter(r => !['RESOLVED','CLOSED'].includes(r.status))
      .length;
  }

  if (permissions.showDocumentIssues) {
    counts.documentIssues = getDocumentsRaw_()
      .filter(d => canViewSite_(user,d.site))
      .filter(d => d.status !== 'COMPLETE')
      .length;
  }

  return {
    company:APP.COMPANY,
    version:APP.VERSION,
    user:user,
    permissions:permissions,
    counts:counts,
    siteSummary:getSiteSummaries_(user)
  };
}

/***************************************************************
 * SITES / CLIENTS
 ***************************************************************/

function getSites(token) {
  const user = getSession_(token);
  return getSiteSummaries_(user);
}

function getSiteSummaries_(user) {
  const employees = getEmployeesRaw_();
  const users = getUsersRaw_();

  return getSitesRaw_()
    .filter(s => canViewSite_(user,s.siteName))
    .map(s => {
      const activeHC = employees.filter(
        e => e.site.toUpperCase() === s.siteName.toUpperCase() && e.status === 'ACTIVE'
      ).length;

      return {
        siteId:s.siteId,
        siteName:s.siteName,
        client:s.client,
        location:s.location,
        region:s.region,
        requiredHC:Number(s.requiredHC)||0,
        activeHC:activeHC,
        shortage:Math.max(0,(Number(s.requiredHC)||0)-activeHC),
        coordinator:s.coordinator,
        teamLeader:s.teamLeader,
        requiredDocuments:s.requiredDocuments,
        status:s.status,
        notes:s.notes,
        coordinatorCount:users.filter(
          u => u.site.toUpperCase() === s.siteName.toUpperCase() &&
               u.role === 'COORDINATOR' &&
               u.status === 'ACTIVE'
        ).length
      };
    });
}

function addSite(token,data) {
  const user = getSession_(token);
  assertRole_(user,['HEAD_ADMIN','HR_HEAD']);

  if (!clean_(data.siteName)) throw new Error('Site name is required.');
  if (!clean_(data.location)) throw new Error('Location is required.');

  if (getSitesRaw_().some(
    s => s.siteName.toUpperCase() === clean_(data.siteName).toUpperCase()
  )) {
    throw new Error('Site already exists.');
  }

  const sh = getSheet_(APP.SHEETS.SITES);
  const id = nextId_(sh,1,'SITE-',3);

  sh.appendRow([
    id,
    clean_(data.siteName),
    clean_(data.client),
    clean_(data.location),
    clean_(data.region),
    numberOrBlank_(data.requiredHC),
    clean_(data.coordinator),
    clean_(data.teamLeader),
    clean_(data.requiredDocuments),
    clean_(data.status || 'ACTIVE').toUpperCase(),
    clean_(data.notes)
  ]);

  logAudit_(user,'SITES','ADD SITE',id,'',data.siteName);

  return {success:true,message:'Site added successfully.',siteId:id};
}

function updateSite(token,data) {
  const user = getSession_(token);
  assertRole_(user,['HEAD_ADMIN','HR_HEAD']);

  const row = findRowById_(APP.SHEETS.SITES,1,data.siteId);
  if (row < 2) throw new Error('Site not found.');

  getSheet_(APP.SHEETS.SITES).getRange(row,1,1,11).setValues([[
    data.siteId,
    clean_(data.siteName),
    clean_(data.client),
    clean_(data.location),
    clean_(data.region),
    numberOrBlank_(data.requiredHC),
    clean_(data.coordinator),
    clean_(data.teamLeader),
    clean_(data.requiredDocuments),
    clean_(data.status || 'ACTIVE').toUpperCase(),
    clean_(data.notes)
  ]]);

  logAudit_(user,'SITES','UPDATE SITE',data.siteId,'','Updated');

  return {success:true,message:'Site updated.'};
}

function getClients(token) {
  getSession_(token);

  return readRows_(getSheet_(APP.SHEETS.CLIENTS)).map(r => ({
    clientId:r[0],
    clientName:r[1],
    contactPerson:r[2],
    contactNumber:r[3],
    email:r[4],
    address:r[5],
    contractStart:r[6],
    contractEnd:r[7],
    status:r[8]
  }));
}

function addClient(token,data) {
  const user = getSession_(token);
  assertRole_(user,['HEAD_ADMIN','HR_HEAD']);

  if (!clean_(data.clientName)) throw new Error('Client name is required.');

  const sh = getSheet_(APP.SHEETS.CLIENTS);
  const id = nextId_(sh,1,'CLIENT-',3);

  sh.appendRow([
    id,
    clean_(data.clientName),
    clean_(data.contactPerson),
    clean_(data.contactNumber),
    clean_(data.email),
    clean_(data.address),
    clean_(data.contractStart),
    clean_(data.contractEnd),
    'ACTIVE'
  ]);

  logAudit_(user,'CLIENTS','ADD CLIENT',id,'',data.clientName);

  return {success:true,message:'Client added.'};
}

/***************************************************************
 * EMPLOYEES
 ***************************************************************/

function getEmployeesBySite(token,site) {
  const user = getSession_(token);
  const actualSite = strictSite_(user,site);

  return getEmployeesRaw_()
    .filter(e => !actualSite || e.site.toUpperCase() === actualSite.toUpperCase())
    .filter(e => canViewSite_(user,e.site))
    .map(e => sanitizeEmployee_(e,user));
}

function getEmployee(token,employeeId) {
  const user = getSession_(token);
  const e = getEmployeesRaw_().find(x => x.employeeId === employeeId);

  if (!e) throw new Error('Employee not found.');
  if (!canViewSite_(user,e.site)) throw new Error('Access denied.');

  return sanitizeEmployee_(e,user);
}

function sanitizeEmployee_(e,user) {
  return e;
}

function addEmployee(token,data) {
  const user = getSession_(token);

  if (!getPermissions_(user).employeeEdit) {
    throw new Error('You do not have permission to add employees.');
  }

  if (!['HEAD_ADMIN','HR_HEAD'].includes(user.role)) data.site = user.site;

  validateEmployee_(data);

  const id = generateEmployeeId_();
  const fullName = buildFullName_(data);

  getSheet_(APP.SHEETS.EMPLOYEES).appendRow(employeeToRow_(id,data,fullName));

  logAudit_(user,'EMPLOYEES','ADD EMPLOYEE',id,'',fullName);

  return {success:true,message:fullName+' added successfully.',employeeId:id};
}

function updateEmployee(token,data) {
  const user = getSession_(token);

  if (!getPermissions_(user).employeeEdit) {
    throw new Error('You do not have permission to edit employees.');
  }

  const existing = getEmployeesRaw_().find(e => e.employeeId === data.employeeId);
  if (!existing) throw new Error('Employee not found.');
  if (!canViewSite_(user,existing.site)) throw new Error('Access denied.');

  if (!['HEAD_ADMIN','HR_HEAD'].includes(user.role)) data.site = user.site;

  validateEmployee_(data);

  const row = findRowById_(APP.SHEETS.EMPLOYEES,1,data.employeeId);
  const fullName = buildFullName_(data);

  getSheet_(APP.SHEETS.EMPLOYEES)
    .getRange(row,1,1,20)
    .setValues([employeeToRow_(data.employeeId,data,fullName)]);

  logAudit_(user,'EMPLOYEES','UPDATE EMPLOYEE',data.employeeId,existing.fullName,fullName);

  return {success:true,message:'Employee updated.'};
}

function bulkImportEmployees(token,rows,defaultSite) {
  const user = getSession_(token);

  if (!getPermissions_(user).employeeBulk) {
    throw new Error('You do not have permission for bulk employee import.');
  }

  if (!Array.isArray(rows) || !rows.length) {
    throw new Error('No employee rows received.');
  }

  const forcedSite = ['HEAD_ADMIN','HR_HEAD'].includes(user.role)
    ? clean_(defaultSite)
    : clean_(user.site);

  const existing = new Set(getEmployeesRaw_().map(e => e.employeeId));
  const output = [];
  let added = 0;
  let skipped = 0;

  rows.forEach(raw => {
    const data = normalizeBulkEmployee_(raw,forcedSite);

    if (!data.firstName || !data.lastName || !data.site) {
      skipped++;
      return;
    }

    if (!['HEAD_ADMIN','HR_HEAD'].includes(user.role)) data.site = user.site;

    let id = clean_(data.employeeId);

    if (!id || existing.has(id)) {
      id = generateEmployeeIdFromSet_(existing);
    }

    existing.add(id);
    output.push(employeeToRow_(id,data,buildFullName_(data)));
    added++;
  });

  if (output.length) {
    const sh = getSheet_(APP.SHEETS.EMPLOYEES);
    sh.getRange(sh.getLastRow()+1,1,output.length,20).setValues(output);
  }

  logAudit_(user,'EMPLOYEES','BULK IMPORT','BULK','',added+' added / '+skipped+' skipped');

  return {success:true,message:'Bulk import completed.',added:added,skipped:skipped};
}

/***************************************************************
 * DOCUMENTS
 ***************************************************************/

function getDocumentsBySite(token,site) {
  const user = getSession_(token);
  const actualSite = strictSite_(user,site);

  return getDocumentsRaw_()
    .filter(d => !actualSite || d.site.toUpperCase() === actualSite.toUpperCase())
    .filter(d => canViewSite_(user,d.site));
}

function uploadEmployeeDocument(token,payload) {
  const user = getSession_(token);

  if (!getPermissions_(user).documentUpload) {
    throw new Error('You do not have permission to upload documents.');
  }

  const employee = getEmployeesRaw_().find(e => e.employeeId === clean_(payload.employeeId));

  if (!employee) throw new Error('Employee not found.');
  if (!canViewSite_(user,employee.site)) throw new Error('Access denied.');

  if (!payload.base64Data || !payload.fileName) {
    throw new Error('File data is missing.');
  }

  const root = getOrCreateFolder_('BetterEdge Employee Documents');
  const siteFolder = getOrCreateChildFolder_(root,employee.site);
  const employeeFolder = getOrCreateChildFolder_(
    siteFolder,
    employee.employeeId+' - '+employee.fullName
  );

  const bytes = Utilities.base64Decode(payload.base64Data);
  const blob = Utilities.newBlob(
    bytes,
    payload.mimeType || 'application/octet-stream',
    payload.fileName
  );

  const file = employeeFolder.createFile(blob);

  const sh = getSheet_(APP.SHEETS.DOCUMENTS);
  const id = nextId_(sh,1,'DOC-',6);

  sh.appendRow([
    id,
    employee.employeeId,
    employee.fullName,
    employee.site,
    clean_(payload.documentType || 'OTHER'),
    payload.fileName,
    file.getId(),
    file.getUrl(),
    clean_(payload.expirationDate),
    clean_(payload.status || 'COMPLETE').toUpperCase(),
    clean_(payload.remarks),
    user.fullName,
    new Date()
  ]);

  logAudit_(user,'DOCUMENTS','UPLOAD',id,'',payload.fileName);

  return {success:true,message:'Document uploaded.',documentId:id};
}

/***************************************************************
 * REPORTS / HELP - SITE PRIVATE
 ***************************************************************/

function submitSiteReport(token,data) {
  const user = getSession_(token);

  const site = ['HEAD_ADMIN','HEAD_COORDINATOR'].includes(user.role) ? clean_(data.site) : clean_(user.site);
  if (!site || site === 'ALL') throw new Error('Select a valid site.');

  const sh = getSheet_(APP.SHEETS.SITE_REPORTS);
  const id = nextId_(sh,1,'RPT-',6);
  const now = new Date();

  sh.appendRow([
    id,
    site,
    clean_(data.department || user.department),
    user.fullName,
    user.role,
    clean_(data.category),
    clean_(data.subject),
    clean_(data.description),
    clean_(data.employeeInvolved),
    clean_(data.priority || 'NORMAL').toUpperCase(),
    'PENDING',
    '',
    now,
    now
  ]);

  logAudit_(user,'SITE_REPORTS','ADD',id,'',data.subject);

  return {success:true,message:'Site report submitted.'};
}

function getSiteReports(token,site) {
  const user = getSession_(token);
  const actualSite = strictSite_(user,site);


  return getReportsRaw_()
    .filter(r => !actualSite || r.site.toUpperCase() === actualSite.toUpperCase());
}

function updateSiteReportStatus(token,reportId,status,response) {
  const user = getSession_(token);

  const row = findRowById_(APP.SHEETS.SITE_REPORTS,1,reportId);
  if (row < 2) throw new Error('Report not found.');

  const sh = getSheet_(APP.SHEETS.SITE_REPORTS);
  const current = sh.getRange(row,1,1,14).getDisplayValues()[0];
  const reportSite = clean_(current[1]);

  if (
    !['HEAD_ADMIN','HEAD_COORDINATOR'].includes(user.role) &&
    clean_(user.site).toUpperCase() !== reportSite.toUpperCase()
  ) {
    throw new Error('You can only update reports for your assigned site.');
  }

  sh.getRange(row,11).setValue(clean_(status).toUpperCase());
  sh.getRange(row,12).setValue(clean_(response));
  sh.getRange(row,14).setValue(new Date());

  logAudit_(user,'SITE_REPORTS','STATUS UPDATE',reportId,current[10],status);

  return {success:true,message:'Report status updated.'};
}

function submitHelpTicket(token,data) {
  const user = getSession_(token);

  const site = ['HEAD_ADMIN','HEAD_COORDINATOR'].includes(user.role) ? clean_(data.site) : clean_(user.site);
  if (!site || site === 'ALL') throw new Error('Select a valid site.');

  const sh = getSheet_(APP.SHEETS.HELP_TICKETS);
  const id = nextId_(sh,1,'HD-',6);
  const now = new Date();

  sh.appendRow([
    id,
    site,
    clean_(data.department || user.department),
    user.fullName,
    user.role,
    clean_(data.category),
    clean_(data.subject),
    clean_(data.description),
    clean_(data.priority || 'NORMAL').toUpperCase(),
    'PENDING',
    '',
    now,
    now
  ]);

  logAudit_(user,'HELP','ADD',id,'',data.subject);

  return {success:true,message:'Help ticket created.'};
}

function getHelpTickets(token,site) {
  const user = getSession_(token);
  const actualSite = strictSite_(user,site);


  return getHelpRaw_()
    .filter(r => !actualSite || r.site.toUpperCase() === actualSite.toUpperCase());
}

function updateHelpTicketStatus(token,ticketId,status,response) {
  const user = getSession_(token);

  const row = findRowById_(APP.SHEETS.HELP_TICKETS,1,ticketId);
  if (row < 2) throw new Error('Ticket not found.');

  const sh = getSheet_(APP.SHEETS.HELP_TICKETS);
  const current = sh.getRange(row,1,1,13).getDisplayValues()[0];
  const site = clean_(current[1]);

  if (
    !['HEAD_ADMIN','HEAD_COORDINATOR'].includes(user.role) &&
    clean_(user.site).toUpperCase() !== site.toUpperCase()
  ) {
    throw new Error('You can only update tickets for your assigned site.');
  }

  sh.getRange(row,10).setValue(clean_(status).toUpperCase());
  sh.getRange(row,11).setValue(clean_(response));
  sh.getRange(row,13).setValue(new Date());

  return {success:true,message:'Help ticket updated.'};
}

/***************************************************************
 * USER MANAGEMENT
 ***************************************************************/

function getUsers(token) {
  const user = getSession_(token);
  assertRole_(user,['HEAD_ADMIN']);

  return getUsersRaw_().map(u => ({
    userId:u.userId,
    username:u.username,
    fullName:u.fullName,
    role:u.role,
    department:u.department,
    site:u.site,
    status:u.status,
    lastLogin:u.lastLogin
  }));
}

function addUser(token,data) {
  const user = getSession_(token);
  assertRole_(user,['HEAD_ADMIN']);

  const username = clean_(data.username).toLowerCase();
  const password = String(data.password || '');
  const role = clean_(data.role).toUpperCase();
  const site = role === 'HEAD_COORDINATOR' ? 'ALL' : clean_(data.site || 'ALL');
  if (['COORDINATOR','TL','VIEWER','HR'].includes(role) && site.toUpperCase()==='ALL') throw new Error('Select a specific assigned site for this role.');
  if (!['HEAD_ADMIN','HEAD_COORDINATOR','HR_HEAD','HR','COORDINATOR','TL','VIEWER'].includes(role)) throw new Error('Unsupported role.');

  if (!username || !password || !clean_(data.fullName) || !role) {
    throw new Error('Name, username, password and role are required.');
  }

  if (getUsersRaw_().some(u => u.username.toLowerCase() === username)) {
    throw new Error('Username already exists.');
  }

  const prefix =
    role === 'HEAD_ADMIN' ? 'ADM-' :
    role === 'HR_HEAD' ? 'HRH-' :
    role === 'HR' ? 'HR-' :
    role === 'HEAD_COORDINATOR' ? 'HCO-' :
    role === 'COORDINATOR' ? 'COORD-' :
    role === 'TL' ? 'TL-' : 'USR-';

  const id = nextRoleId_(prefix);

  getSheet_(APP.SHEETS.USERS).appendRow([
    id,
    username,
    hashPassword_(password),
    clean_(data.fullName),
    role,
    clean_(data.department),
    site,
    'ACTIVE',
    ''
  ]);

  logAudit_(user,'USERS','ADD USER',id,'',username);

  return {success:true,message:'User access created.',userId:id};
}

function setUserStatus(token,userId,status) {
  const user = getSession_(token);
  assertRole_(user,['HEAD_ADMIN']);

  if (user.userId === userId && clean_(status).toUpperCase() !== 'ACTIVE') {
    throw new Error('You cannot deactivate your own current account.');
  }

  const row = findRowById_(APP.SHEETS.USERS,1,userId);
  if (row < 2) throw new Error('User not found.');

  getSheet_(APP.SHEETS.USERS)
    .getRange(row,8)
    .setValue(clean_(status).toUpperCase());

  return {success:true};
}

function resetUserPassword(token,userId,newPassword) {
  const user = getSession_(token);
  assertRole_(user,['HEAD_ADMIN']);

  if (String(newPassword || '').length < 6) {
    throw new Error('Password must contain at least 6 characters.');
  }

  const row = findRowById_(APP.SHEETS.USERS,1,userId);
  if (row < 2) throw new Error('User not found.');

  getSheet_(APP.SHEETS.USERS)
    .getRange(row,3)
    .setValue(hashPassword_(newPassword));

  return {success:true,message:'Password reset.'};
}

/***************************************************************
 * AUDIT / SUMMARIES
 ***************************************************************/

function getAuditLog(token) {
  const user = getSession_(token);
  assertRole_(user,['HEAD_ADMIN']);

  const sh = getSheet_(APP.SHEETS.AUDIT_LOG);

  return {
    headers:sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0],
    rows:readRows_(sh).reverse().slice(0,500)
  };
}







/***************************************************************
 * RAW READERS
 ***************************************************************/

function getUsersRaw_() {
  return readRows_(getSheet_(APP.SHEETS.USERS)).map(r => ({
    userId:r[0],
    username:r[1],
    passwordHash:r[2],
    fullName:r[3],
    role:clean_(r[4]).toUpperCase(),
    department:clean_(r[5]).toUpperCase(),
    site:clean_(r[6]),
    status:clean_(r[7]).toUpperCase(),
    lastLogin:r[8]
  }));
}

function getEmployeesRaw_() {
  return readRows_(getSheet_(APP.SHEETS.EMPLOYEES)).map(r => ({
    employeeId:r[0],
    lastName:r[1],
    firstName:r[2],
    middleName:r[3],
    fullName:r[4],
    birthday:r[5],
    gender:r[6],
    contactNumber:r[7],
    email:r[8],
    address:r[9],
    emergencyContact:r[10],
    emergencyNumber:r[11],
    dateHired:r[12],
    client:r[13],
    site:clean_(r[14]),
    department:r[15],
    position:r[16],
    employmentType:r[17],
    status:clean_(r[18]).toUpperCase(),
    lastUpdated:r[19]
  }));
}

function getSitesRaw_() {
  return readRows_(getSheet_(APP.SHEETS.SITES)).map(r => ({
    siteId:r[0],
    siteName:clean_(r[1]),
    client:clean_(r[2]),
    location:clean_(r[3]),
    region:clean_(r[4]),
    requiredHC:Number(r[5]) || 0,
    coordinator:clean_(r[6]),
    teamLeader:clean_(r[7]),
    requiredDocuments:clean_(r[8]),
    status:clean_(r[9]).toUpperCase(),
    notes:clean_(r[10])
  }));
}

function getDocumentsRaw_() {
  return readRows_(getSheet_(APP.SHEETS.DOCUMENTS)).map(r => ({
    documentId:r[0],
    employeeId:r[1],
    employeeName:r[2],
    site:clean_(r[3]),
    documentType:r[4],
    fileName:r[5],
    driveFileId:r[6],
    driveUrl:r[7],
    expirationDate:r[8],
    status:clean_(r[9]).toUpperCase(),
    remarks:r[10],
    uploadedBy:r[11],
    uploadedDate:r[12]
  }));
}

function getReportsRaw_() {
  return readRows_(getSheet_(APP.SHEETS.SITE_REPORTS)).map(r => ({
    reportId:r[0],
    site:clean_(r[1]),
    department:r[2],
    reportedBy:r[3],
    role:r[4],
    category:r[5],
    subject:r[6],
    description:r[7],
    employeeInvolved:r[8],
    priority:r[9],
    status:clean_(r[10]).toUpperCase(),
    managementResponse:r[11],
    reportedDate:r[12],
    lastUpdated:r[13]
  }));
}

function getHelpRaw_() {
  return readRows_(getSheet_(APP.SHEETS.HELP_TICKETS)).map(r => ({
    ticketId:r[0],
    site:clean_(r[1]),
    department:r[2],
    reportedBy:r[3],
    role:r[4],
    category:r[5],
    subject:r[6],
    description:r[7],
    priority:r[8],
    status:clean_(r[9]).toUpperCase(),
    managementResponse:r[10],
    createdDate:r[11],
    lastUpdated:r[12]
  }));
}

/***************************************************************
 * RESTORE / NORMALIZE
 ***************************************************************/

function normalizeBulkEmployee_(raw,defaultSite) {
  const o = {};
  Object.keys(raw || {}).forEach(k => {
    o[String(k).trim().toLowerCase()] = raw[k];
  });

  return {
    employeeId:pick_(o,['employee id','employeeid','id']),
    lastName:pick_(o,['last name','lastname','surname']),
    firstName:pick_(o,['first name','firstname','given name']),
    middleName:pick_(o,['middle name','middlename','middle initial']),
    birthday:pick_(o,['birthday','birthdate','date of birth']),
    gender:pick_(o,['gender','sex']),
    contactNumber:pick_(o,['contact number','phone','mobile','cp number']),
    email:pick_(o,['email','email address']),
    address:pick_(o,['address']),
    emergencyContact:pick_(o,['emergency contact']),
    emergencyNumber:pick_(o,['emergency number','emergency phone']),
    dateHired:pick_(o,['date hired','hired date']),
    client:pick_(o,['client']),
    site:pick_(o,['site','assigned site']) || defaultSite,
    department:pick_(o,['department']),
    position:pick_(o,['position','job title']),
    employmentType:pick_(o,['employment type','type']),
    status:pick_(o,['status']) || 'ACTIVE'
  };
}

function restoreEmployees_(oldData) {
  if (!oldData || oldData.length < 2) return;

  const headers = oldData[0].map(h => clean_(h).toLowerCase());
  const out = [];

  oldData.slice(1).forEach(row => {
    if (!row.some(Boolean)) return;

    const get = names => getFromOld_(headers,row,names);
    const first = get(['first name']);
    const middle = get(['middle name']);
    const last = get(['last name']);
    const full = get(['full name']) || [first,middle,last].filter(Boolean).join(' ').toUpperCase();

    if (!full && !get(['employee id'])) return;

    out.push([
      get(['employee id']),
      last,first,middle,full,
      get(['birthday']),get(['gender']),get(['contact number']),
      get(['email']),get(['address']),get(['emergency contact']),
      get(['emergency number']),get(['date hired']),get(['client']),
      get(['site']),get(['department']),get(['position']),
      get(['employment type']),get(['status']) || 'ACTIVE',
      new Date()
    ]);
  });

  if (out.length) {
    getSheet_(APP.SHEETS.EMPLOYEES).getRange(2,1,out.length,20).setValues(out);
  }
}

function restoreSites_(oldData) {
  if (!oldData || oldData.length < 2) return;

  const headers = oldData[0].map(h => clean_(h).toLowerCase());
  const out = [];

  oldData.slice(1).forEach(row => {
    const get = names => getFromOld_(headers,row,names);
    const name = get(['site name']);
    if (!name) return;

    out.push([
      get(['site id']) || 'SITE-' + String(out.length+1).padStart(3,'0'),
      name,
      get(['client','primary client']),
      get(['location','province / city']),
      get(['region','region / area']),
      get(['required hc']),
      get(['coordinator']),
      get(['team leader']),
      get(['required documents']),
      get(['status']) || 'ACTIVE',
      get(['notes'])
    ]);
  });

  if (out.length) {
    getSheet_(APP.SHEETS.SITES).getRange(2,1,out.length,11).setValues(out);
  }
}

function restoreClients_(oldData) {
  if (!oldData || oldData.length < 2) return;

  const headers = oldData[0].map(h => clean_(h).toLowerCase());
  const out = [];

  oldData.slice(1).forEach(row => {
    const get = names => getFromOld_(headers,row,names);
    const name = get(['client name']);
    if (!name) return;

    out.push([
      get(['client id']) || 'CLIENT-' + String(out.length+1).padStart(3,'0'),
      name,
      get(['contact person']),
      get(['contact number']),
      get(['email']),
      get(['address']),
      get(['contract start']),
      get(['contract end']),
      get(['status']) || 'ACTIVE'
    ]);
  });

  if (out.length) {
    getSheet_(APP.SHEETS.CLIENTS).getRange(2,1,out.length,9).setValues(out);
  }
}

function getFromOld_(headers,row,names) {
  for (const name of names) {
    const i = headers.indexOf(String(name).toLowerCase());
    if (i !== -1) return row[i] || '';
  }
  return '';
}

/***************************************************************
 * HELPERS
 ***************************************************************/

function getSheet_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error(name+' sheet not found. Run setupBesiWorkforce().');
  return sh;
}

function readRows_(sh) {
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();

  if (lastRow < 2 || lastCol < 1) return [];

  return sh.getRange(2,1,lastRow-1,lastCol).getDisplayValues();
}

function clean_(v) {
  return String(v == null ? '' : v).trim();
}

function numberOrBlank_(v) {
  if (v === '' || v === null || typeof v === 'undefined') return '';
  const n = Number(v);
  return isNaN(n) ? '' : n;
}

function pick_(obj,keys) {
  for (const key of keys) {
    if (obj[key] !== undefined && clean_(obj[key]) !== '') return obj[key];
  }
  return '';
}

function validateEmployee_(d) {
  if (!d) throw new Error('Employee information is missing.');
  if (!clean_(d.lastName)) throw new Error('Last Name is required.');
  if (!clean_(d.firstName)) throw new Error('First Name is required.');
  if (!clean_(d.site)) throw new Error('Site is required.');
}

function buildFullName_(d) {
  return [
    clean_(d.firstName),
    clean_(d.middleName),
    clean_(d.lastName)
  ].filter(Boolean).join(' ').toUpperCase();
}

function employeeToRow_(id,d,fullName) {
  return [
    id,
    clean_(d.lastName).toUpperCase(),
    clean_(d.firstName).toUpperCase(),
    clean_(d.middleName).toUpperCase(),
    fullName,
    clean_(d.birthday),
    clean_(d.gender).toUpperCase(),
    clean_(d.contactNumber),
    clean_(d.email),
    clean_(d.address),
    clean_(d.emergencyContact),
    clean_(d.emergencyNumber),
    clean_(d.dateHired),
    clean_(d.client),
    clean_(d.site),
    clean_(d.department),
    clean_(d.position),
    clean_(d.employmentType),
    clean_(d.status || 'ACTIVE').toUpperCase(),
    new Date()
  ];
}

function generateEmployeeId_() {
  return generateEmployeeIdFromSet_(
    new Set(getEmployeesRaw_().map(e => e.employeeId))
  );
}

function generateEmployeeIdFromSet_(set) {
  let max = 1000;

  set.forEach(id => {
    const m = String(id || '').match(/(\d+)$/);
    if (m) max = Math.max(max,Number(m[1]));
  });

  return 'BESI-' + (max+1);
}

function nextId_(sh,col,prefix,pad) {
  const vals = sh.getLastRow() < 2
    ? []
    : sh.getRange(2,col,sh.getLastRow()-1,1).getDisplayValues().flat();

  let max = 0;

  vals.forEach(v => {
    const m = String(v || '').match(/(\d+)$/);
    if (m) max = Math.max(max,Number(m[1]));
  });

  return prefix + String(max+1).padStart(pad,'0');
}

function nextRoleId_(prefix) {
  let max = 0;

  getUsersRaw_().forEach(u => {
    if (u.userId.startsWith(prefix)) {
      const m = u.userId.match(/(\d+)$/);
      if (m) max = Math.max(max,Number(m[1]));
    }
  });

  return prefix + String(max+1).padStart(4,'0');
}

function findRowById_(sheetName,col,id) {
  const sh = getSheet_(sheetName);

  if (sh.getLastRow() < 2) return -1;

  const values = sh.getRange(2,col,sh.getLastRow()-1,1).getDisplayValues();

  for (let i=0;i<values.length;i++) {
    if (String(values[i][0]) === String(id)) return i+2;
  }

  return -1;
}

function hashPassword_(password) {
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(password || ''),
    Utilities.Charset.UTF_8
  );

  return digest.map(b => {
    const n=(b+256)%256;
    return ('0'+n.toString(16)).slice(-2);
  }).join('');
}





function getOrCreateFolder_(name) {
  const folders=DriveApp.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(name);
}

function getOrCreateChildFolder_(parent,name) {
  const folders=parent.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return parent.createFolder(name);
}

function logAudit_(user,module,action,recordId,oldValue,newValue) {
  getSheet_(APP.SHEETS.AUDIT_LOG).appendRow([
    new Date(),
    user && user.fullName ? user.fullName : 'SYSTEM',
    user && user.role ? user.role : 'SYSTEM',
    user && user.department ? user.department : '',
    module || '',
    action || '',
    recordId || '',
    oldValue || '',
    newValue || ''
  ]);
}


/** Run once in a NEW spreadsheet-bound project to copy workforce records. */
function migrateV41WorkforceRecords(sourceSpreadsheetId) {
  sourceSpreadsheetId = sourceSpreadsheetId || PropertiesService.getScriptProperties().getProperty('SOURCE_SPREADSHEET_ID');
  if (!sourceSpreadsheetId) throw new Error('Set SOURCE_SPREADSHEET_ID in Script Properties first.');
  const source = SpreadsheetApp.openById(clean_(sourceSpreadsheetId));
  const target = SpreadsheetApp.getActiveSpreadsheet();
  if (source.getId() === target.getId()) throw new Error('Source and target must differ.');
  setupBesiWorkforce();
  const defs = getDefinitions_();
  const names = Object.keys(defs).filter(n => n !== APP.SHEETS.SETTINGS);
  names.forEach(n => {
    const dest = target.getSheetByName(n);
    const allowed = n === APP.SHEETS.USERS ? 2 : 1;
    if (dest.getLastRow()>allowed) throw new Error('Target '+n+' already contains records.');
  });
  const totals = {};
  names.forEach(n => {
    const old = source.getSheetByName(n);
    if (!old || old.getLastRow()<2) {totals[n]=0;return;}
    const all = old.getDataRange().getValues();
    const headers = all[0].map(h=>String(h).trim().toLowerCase());
    const dest = target.getSheetByName(n);
    const mapped = all.slice(1).filter(row=>row.some(v=>v!==''))
      .filter(row=>{
        if (n!==APP.SHEETS.USERS) return true;
        const role=String(row[headers.indexOf('role')]||'').toUpperCase();
        const username=String(row[headers.indexOf('username')]||'').toLowerCase();
        return ['HEAD_ADMIN','HEAD_COORDINATOR','HR_HEAD','HR','COORDINATOR','TL','VIEWER'].includes(role) && username!=='admin';
      }).map(row=>defs[n].map(label=>{
        const i=headers.indexOf(label.toLowerCase());
        return i<0?'':row[i];
      }));
    if (mapped.length) dest.getRange(dest.getLastRow()+1,1,mapped.length,defs[n].length).setValues(mapped);
    totals[n]=mapped.length;
  });
  return {success:true,imported:totals};
}
