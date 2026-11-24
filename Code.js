function doGet(e) {
  if (e && e.parameter && (e.parameter.code || e.parameter.error)) finishSignIn_(e.parameter);
  return HtmlService.createHtmlOutputFromFile("Index").setTitle("EZ-Ensemble");
}

// Run once from the editor so the demo sheet belongs to the project owner.
function setupDemo_() {
  const properties = PropertiesService.getScriptProperties();
  if (!properties.getProperty("DEMO_SHEET_ID")) {
    const spreadsheet = SpreadsheetApp.create("EZ-Ensemble Plumbing Demo");
    const sheet = spreadsheet.getSheets()[0].setName("Demo");
    sheet.getRange("A1").setValue("Hello from EZ-Ensemble!");
    properties.setProperty("DEMO_SHEET_ID", spreadsheet.getId());
  }
  console.log(SpreadsheetApp.openById(properties.getProperty("DEMO_SHEET_ID")).getUrl());
}

function getDemoMessage() {
  const id = PropertiesService.getScriptProperties().getProperty("DEMO_SHEET_ID");
  if (!id) throw new Error("Run setupDemo_ in the Apps Script editor first.");
  const sheet = SpreadsheetApp.openById(id).getSheetByName("Demo");
  if (!sheet) throw new Error("The demo spreadsheet needs a tab named Demo.");
  return sheet.getRange("A1").getDisplayValue();
}

function initializeDemo() {
  setupDemo_();
}

const demoAccounts_ = ['enricsoni14@gmail.com', 'enrson27@bergen.org'];
const columns_ = ['Instrument', 'Tone', 'Pitch', 'Rhythm'];

function microsoftConfig_() {
  const p = PropertiesService.getScriptProperties().getProperties();
  if (!p.MS_CLIENT_ID || !p.MS_CLIENT_SECRET || !p.MS_TENANT_ID)
    throw new Error('Microsoft sign-in needs an Entra app registration. Set MS_CLIENT_ID, MS_CLIENT_SECRET and MS_TENANT_ID in Script Properties.');
  return p;
}

function startSignIn() {
  const p = microsoftConfig_();
  const state = Utilities.getUuid() + Utilities.getUuid();
  CacheService.getUserCache().put('oauthState:' + state, 'pending', 600);
  const params = {client_id:p.MS_CLIENT_ID, response_type:'code', redirect_uri:ScriptApp.getService().getUrl(),
    scope:'https://graph.microsoft.com/User.Read', state:state, prompt:'select_account'};
  return 'https://login.microsoftonline.com/' + encodeURIComponent(p.MS_TENANT_ID) + '/oauth2/v2.0/authorize?' +
    Object.entries(params).map(([k,v]) => k + '=' + encodeURIComponent(v)).join('&');
}

function finishSignIn_(query) {
  const cache = CacheService.getUserCache();
  const key = 'oauthState:' + query.state;
  const lock = LockService.getUserLock();
  lock.waitLock(5000);
  try {
    if (!/^[a-f0-9-]{72}$/i.test(query.state || '') || cache.get(key) !== 'pending')
      throw new Error('This sign-in link expired or was already used. Open the web app to start again.');
    cache.remove(key);
  } finally {
    lock.releaseLock();
  }
  if (!query.code) throw new Error('Microsoft sign-in was canceled. Open the app and try again.');
  const p = microsoftConfig_();
  const token = UrlFetchApp.fetch('https://login.microsoftonline.com/' + encodeURIComponent(p.MS_TENANT_ID) + '/oauth2/v2.0/token',
    {method:'post', muteHttpExceptions:true, payload:{client_id:p.MS_CLIENT_ID, client_secret:p.MS_CLIENT_SECRET,
      grant_type:'authorization_code', code:query.code, redirect_uri:ScriptApp.getService().getUrl()}});
  if (token.getResponseCode() !== 200) throw new Error('Microsoft could not complete sign-in. Check the registered redirect URI and client secret.');
  const profile = JSON.parse(UrlFetchApp.fetch('https://graph.microsoft.com/v1.0/me?$select=displayName,userPrincipalName,mail',
    {headers:{Authorization:'Bearer ' + JSON.parse(token.getContentText()).access_token}}).getContentText());
  const email = [profile.userPrincipalName, profile.mail]
    .filter(value => typeof value === 'string')
    .map(value => value.toLowerCase())
    .find(value => demoAccounts_.includes(value));
  if (!email) throw new Error('Choose the personal Microsoft account or Bergen account registered for this private demo.');
  cache.put('scorer', JSON.stringify({name:profile.displayName,email:email}), 3600);
}

function currentScorer() {
  return JSON.parse(CacheService.getUserCache().get('scorer') || 'null');
}

function scoringBook_() {
  if (!currentScorer()) throw new Error('Sign in with Microsoft to access scores.');
  return SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('DEMO_SHEET_ID'));
}

function updateScores(scores) {
  const book = scoringBook_();
  if (!Array.isArray(scores) || scores.length !== 3 || !scores.every(Number.isFinite)) throw new Error('Enter three numerical scores.');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = book.getSheetByName('Audition Scores') || book.insertSheet('Audition Scores');
    sheet.getRange('A1:D2').setValues([columns_, ['Flute', ...scores]]);
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
}

function retrieveScores() {
  const sheet = scoringBook_().getSheetByName('Audition Scores');
  return sheet ? sheet.getRange('A1:D2').getDisplayValues() : [];
}
