// http://qiita.com/kiita312/items/d5ffd0207b8411c159e1
// https://developers.google.com/webmaster-tools/search-console-api-original/v3/searchanalytics/query

const OAUTH_CLIENT_ID = PropertiesService.getScriptProperties().getProperty('OAUTH_CLIENT_ID');
const OAUTH_CLIENT_SECRET = PropertiesService.getScriptProperties().getProperty('OAUTH_CLIENT_SECRET');
const SITE_URL = PropertiesService.getScriptProperties().getProperty('SITE_URL');
const SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');

declare var OAuth2: any;

function getWebmastersService(): any {
  return OAuth2.createService('webmasters')
    .setAuthorizationBaseUrl('https://accounts.google.com/o/oauth2/auth')
    .setTokenUrl('https://accounts.google.com/o/oauth2/token')
    .setClientId(OAUTH_CLIENT_ID)
    .setClientSecret(OAUTH_CLIENT_SECRET)
    .setCallbackFunction('authCallback')
    .setPropertyStore(PropertiesService.getUserProperties())
    .setScope('https://www.googleapis.com/auth/webmasters.readonly')
    .setParam('access_type', 'offline');
}

function init() {
  let wmService = getWebmastersService();
  if (!wmService.hasAccess()) {
    Logger.log(wmService.getAuthorizationUrl());
  }
}

function authCallback(request) {
  let wmService = getWebmastersService();
  let isAuthorized = wmService.handleCallback(request);
  if (isAuthorized) {
    return HtmlService.createHtmlOutput('Success! You can close this tab.');
  } else {
    return HtmlService.createHtmlOutput('Denied. You can close this tab');
  }
}

const METRIC_NAMES = ['ctr', 'clicks', 'impressions', 'position'] as (keyof QueryResponseRow)[];

function main() {
  let date = new Date()
  date.setDate(date.getDate() - 3);
  importData(date);
}

function importData(date: Date) {
  Logger.log(`importData(${date})`);

  let spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheetsByName: { [metric: string]: GoogleAppsScript.Spreadsheet.Sheet } = {};

  for (let metric of METRIC_NAMES) {
    let sheet = spreadsheet.getSheetByName(metric);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(metric)
      sheet.setFrozenRows(1);
      sheet.setFrozenColumns(1);
    }
    sheetsByName[metric] = sheet;
  }

  let dateString = date.toISOString().substring(0, 10);

  let resp = fetchData(date);
  for (let metric of METRIC_NAMES) {
    let sheet = sheetsByName[metric];

    let col = sheet.getLastColumn();
    if (col === 0) {
      col = 1;
    }
    sheet.getRange(1, col+1).setValue(dateString);
    Logger.log(`(1, ${col+1}) = ${dateString}`);

    let pairs = resp.rows.map((dataRow) => {
      return { key: dataRow.keys[0], value: dataRow[metric] };
    });

    let range = sheet.getRange(1, 1, sheet.getLastRow() + pairs.length, sheet.getLastColumn());
    let sheetValues = range.getValues();
    let newRowIndex = findIndex(sheet.getRange('A2:A').getValues(), (vv) => vv[0] === '');
    if (newRowIndex === -1) {
      throw new Error('no empty row found');
    }

    for (let {key, value} of pairs) {
      let rowIndex = findIndex(sheetValues, (row) => row[0] === key);
      if (rowIndex === -1) {
        rowIndex = ++newRowIndex;
        sheetValues[rowIndex][0] = key;
      }
      sheetValues[rowIndex][col] = value;
    }
    range.setValues(sheetValues);
  }
}

interface QueryResponse {
  rows: QueryResponseRow[];
}

interface QueryResponseRow {
  keys: string[];
  ctr: number;
  clicks: number;
  impressions: number;
  position: number;
}

function fetchData(date: Date): QueryResponse {
  let dateString = date.toISOString().substring(0, 10);
  let query = {
    rowLimit: 50,
    startDate: dateString,
    endDate: dateString,
    dimensions: [ 'query' ]
  };
  let wmService = getWebmastersService();
  let response = UrlFetchApp.fetch(`https://www.googleapis.com/webmasters/v3/sites/${encodeURIComponent(SITE_URL)}/searchAnalytics/query`, {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: `Bearer ${wmService.getAccessToken()}`
    },
    payload: JSON.stringify(query)
  });
  return JSON.parse(response.getContentText());
}

function findIndex<T>(array: T[], pred: (T) => boolean): number {
  for (let i = 0; i < array.length; i++) {
    if (pred(array[i])) {
      return i;
    }
  }
  return -1;
}
