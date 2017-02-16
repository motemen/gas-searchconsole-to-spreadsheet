var OAUTH_CLIENT_ID = PropertiesService.getScriptProperties().getProperty('OAUTH_CLIENT_ID');
var OAUTH_CLIENT_SECRET = PropertiesService.getScriptProperties().getProperty('OAUTH_CLIENT_SECRET');
var SITE_URL = PropertiesService.getScriptProperties().getProperty('SITE_URL');
var SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
function getWebmastersService() {
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
    var wmService = getWebmastersService();
    if (!wmService.hasAccess()) {
        Logger.log(wmService.getAuthorizationUrl());
    }
}
function authCallback(request) {
    var wmService = getWebmastersService();
    var isAuthorized = wmService.handleCallback(request);
    if (isAuthorized) {
        return HtmlService.createHtmlOutput('Success! You can close this tab.');
    }
    else {
        return HtmlService.createHtmlOutput('Denied. You can close this tab');
    }
}
var METRIC_NAMES = ['ctr', 'clicks', 'impressions', 'position'];
function main() {
    var date = new Date();
    date.setDate(date.getDate() - 3);
    importData(date);
}
function importData(date) {
    Logger.log("importData(" + date + ")");
    var spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheetsByName = {};
    for (var _i = 0, METRIC_NAMES_1 = METRIC_NAMES; _i < METRIC_NAMES_1.length; _i++) {
        var metric = METRIC_NAMES_1[_i];
        var sheet = spreadsheet.getSheetByName(metric);
        if (!sheet) {
            sheet = spreadsheet.insertSheet(metric);
            sheet.setFrozenRows(1);
            sheet.setFrozenColumns(1);
        }
        sheetsByName[metric] = sheet;
    }
    var dateString = date.toISOString().substring(0, 10);
    var resp = fetchData(date);
    var _loop_1 = function (metric) {
        var sheet = sheetsByName[metric];
        var col = sheet.getLastColumn();
        if (col === 0) {
            col = 1;
        }
        sheet.getRange(1, col + 1).setValue(dateString);
        Logger.log("(1, " + (col + 1) + ") = " + dateString);
        var pairs = resp.rows.map(function (dataRow) {
            return { key: dataRow.keys[0], value: dataRow[metric] };
        });
        var range = sheet.getRange(1, 1, sheet.getLastRow() + pairs.length, sheet.getLastColumn());
        var sheetValues = range.getValues();
        var newRowIndex = findIndex(sheet.getRange('A2:A').getValues(), function (vv) { return vv[0] === ''; });
        if (newRowIndex === -1) {
            throw new Error('no empty row found');
        }
        var _loop_2 = function (key, value) {
            var rowIndex = findIndex(sheetValues, function (row) { return row[0] === key; });
            if (rowIndex === -1) {
                rowIndex = ++newRowIndex;
                sheetValues[rowIndex][0] = key;
            }
            sheetValues[rowIndex][col] = value;
        };
        for (var _i = 0, pairs_1 = pairs; _i < pairs_1.length; _i++) {
            var _a = pairs_1[_i], key = _a.key, value = _a.value;
            _loop_2(key, value);
        }
        range.setValues(sheetValues);
    };
    for (var _a = 0, METRIC_NAMES_2 = METRIC_NAMES; _a < METRIC_NAMES_2.length; _a++) {
        var metric = METRIC_NAMES_2[_a];
        _loop_1(metric);
    }
}
function fetchData(date) {
    var dateString = date.toISOString().substring(0, 10);
    var query = {
        rowLimit: 50,
        startDate: dateString,
        endDate: dateString,
        dimensions: ['query']
    };
    var wmService = getWebmastersService();
    var response = UrlFetchApp.fetch("https://www.googleapis.com/webmasters/v3/sites/" + encodeURIComponent(SITE_URL) + "/searchAnalytics/query", {
        method: 'post',
        contentType: 'application/json',
        headers: {
            Authorization: "Bearer " + wmService.getAccessToken()
        },
        payload: JSON.stringify(query)
    });
    return JSON.parse(response.getContentText());
}
function findIndex(array, pred) {
    for (var i = 0; i < array.length; i++) {
        if (pred(array[i])) {
            return i;
        }
    }
    return -1;
}
