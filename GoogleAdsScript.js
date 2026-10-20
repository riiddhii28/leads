/**
 * Google Ads Scripts — runs inside Google Ads, NOT inside the Sheets-bound Apps Script editor.
 *
 * Configure SPREADSHEET_URL and CLIENT_ID before running. The destination spreadsheet
 * must already have the Google tab created by Code.gs / setupPrototype().
 *
 * This script pulls campaign-level performance rows (Record Type = PERFORMANCE) for the
 * Master summary, plus ad-level rows (Record Type = DETAIL) for drill-down. The Master
 * ignores DETAIL rows so metrics are not counted twice.
 *
 * It does not retrieve individual CRM/IVR leads; those must come from the client's CRM/IVR/lead source.
 */

var CONFIG = {
  SPREADSHEET_URL: 'PASTE_THE_CLIENT_SPREADSHEET_URL_HERE',
  CLIENT_ID: 'CLIENT_DEMO_001',
  CLIENT_NAME: 'Demo Client'
};

var DETAIL_HEADERS = [
  'Record Type', 'Client ID', 'Client Name', 'Platform', 'Data Mode',
  'Account ID', 'Account Name', 'Currency', 'Date', 'Campaign ID', 'Campaign Name',
  'Ad Group / Ad Set ID', 'Ad Group / Ad Set Name', 'Ad ID', 'Ad Name', 'Spend',
  'Clicks', 'Impressions', 'Reported Conversions', 'Reported Leads', 'Cost Per Lead',
  'Lead ID', 'Lead Created At', 'Lead Name', 'Lead Email', 'Lead Phone', 'Lead Source',
  'Form ID', 'Lead Status', 'Record Origin', 'Custom Fields JSON', 'Record Key', 'Synced At'
];

function main() {
  if (!CONFIG.SPREADSHEET_URL || CONFIG.SPREADSHEET_URL.indexOf('PASTE_') === 0) {
    throw new Error('Set CONFIG.SPREADSHEET_URL to the full URL of the client spreadsheet first.');
  }

  var spreadsheet = SpreadsheetApp.openByUrl(CONFIG.SPREADSHEET_URL);
  var sheet = spreadsheet.getSheetByName('Google');
  if (!sheet) throw new Error('The Google tab was not found. Run setupPrototype() in the spreadsheet first.');

  var account = AdsApp.currentAccount();
  var customerId = normalizeAccountId_(account.getCustomerId());
  var accountName = account.getName();
  var currency = account.getCurrencyCode();
  var syncedAt = new Date();
  var incomingRows = [];

  // Campaign-level reporting is the source used by the Master summary.
  var campaignQuery =
    'SELECT customer.id, customer.descriptive_name, customer.currency_code, ' +
    'campaign.id, campaign.name, segments.date, ' +
    'metrics.cost_micros, metrics.clicks, metrics.impressions, metrics.conversions ' +
    'FROM campaign ' +
    'WHERE segments.date DURING LAST_30_DAYS ' +
    "AND campaign.status != 'REMOVED'";
  appendReportRows_(AdsApp.report(campaignQuery), incomingRows, {
    recordType: 'PERFORMANCE',
    accountId: customerId,
    accountName: accountName,
    currency: currency,
    syncedAt: syncedAt
  });

  // Ad-level detail is retained for reporting/drill-down but excluded from Master totals.
  var adQuery =
    'SELECT customer.id, customer.descriptive_name, customer.currency_code, ' +
    'campaign.id, campaign.name, ad_group.id, ad_group.name, ' +
    'ad_group_ad.ad.id, ad_group_ad.ad.name, segments.date, ' +
    'metrics.cost_micros, metrics.clicks, metrics.impressions, metrics.conversions ' +
    'FROM ad_group_ad ' +
    'WHERE segments.date DURING LAST_30_DAYS ' +
    "AND campaign.status != 'REMOVED' AND ad_group.status != 'REMOVED' " +
    "AND ad_group_ad.status != 'REMOVED'";
  appendReportRows_(AdsApp.report(adQuery), incomingRows, {
    recordType: 'DETAIL',
    accountId: customerId,
    accountName: accountName,
    currency: currency,
    syncedAt: syncedAt
  });

  var result = upsertDetailRows_(sheet, incomingRows);
  Logger.log('Google Ads sync finished. Account ' + customerId + ': ' + incomingRows.length + ' rows fetched; ' + result.added + ' added, ' + result.updated + ' updated.');
  Logger.log('Open the spreadsheet and choose Lead Retrieval → Refresh Master summary to rebuild the daily summary.');
}

function appendReportRows_(report, destination, options) {
  var iterator = report.rows();
  while (iterator.hasNext()) {
    var row = iterator.next();
    var dateKey = String(row['segments.date']);
    var reportedSpend = Number(row['metrics.cost_micros'] || 0) / 1000000;
    var clicks = Number(row['metrics.clicks'] || 0);
    var impressions = Number(row['metrics.impressions'] || 0);
    var conversions = Number(row['metrics.conversions'] || 0);
    var customerId = normalizeAccountId_(row['customer.id'] || options.accountId);
    var accountName = row['customer.descriptive_name'] || options.accountName;
    var currency = row['customer.currency_code'] || options.currency;
    var campaignId = String(row['campaign.id'] || '');
    var campaignName = row['campaign.name'] || '';
    var adGroupId = row['ad_group.id'] ? String(row['ad_group.id']) : '';
    var adGroupName = row['ad_group.name'] || '';
    var adId = row['ad_group_ad.ad.id'] ? String(row['ad_group_ad.ad.id']) : '';
    var adName = row['ad_group_ad.ad.name'] || '';

    var recordKey;
    if (options.recordType === 'PERFORMANCE') {
      recordKey = 'GOOGLE|PERFORMANCE|' + dateKey + '|' + customerId + '|' + campaignId;
    } else {
      recordKey = 'GOOGLE|DETAIL|' + dateKey + '|' + customerId + '|' + campaignId + '|' + adGroupId + '|' + adId;
    }

    var values = {
      'Record Type': options.recordType,
      'Client ID': CONFIG.CLIENT_ID,
      'Client Name': CONFIG.CLIENT_NAME,
      'Platform': 'Google',
      'Data Mode': 'LIVE',
      'Account ID': customerId,
      'Account Name': accountName,
      'Currency': currency,
      'Date': dateFromKey_(dateKey),
      'Campaign ID': campaignId,
      'Campaign Name': campaignName,
      'Ad Group / Ad Set ID': adGroupId,
      'Ad Group / Ad Set Name': adGroupName,
      'Ad ID': adId,
      'Ad Name': adName,
      'Spend': round2_(reportedSpend),
      'Clicks': clicks,
      'Impressions': impressions,
      'Reported Conversions': conversions,
      // Conversions are not automatically the same as qualified/CRM/IVR leads.
      'Reported Leads': '',
      'Cost Per Lead': '',
      'Lead ID': '',
      'Lead Created At': '',
      'Lead Name': '',
      'Lead Email': '',
      'Lead Phone': '',
      'Lead Source': '',
      'Form ID': '',
      'Lead Status': '',
      'Record Origin': 'GOOGLE_ADS_SCRIPTS',
      'Custom Fields JSON': '',
      'Record Key': recordKey,
      'Synced At': options.syncedAt
    };
    destination.push(DETAIL_HEADERS.map(function (header) {
      return values[header] === undefined || values[header] === null ? '' : values[header];
    }));
  }
}

function upsertDetailRows_(sheet, incomingRows) {
  if (!incomingRows.length) return { added: 0, updated: 0 };
  var keyIndex = DETAIL_HEADERS.indexOf('Record Key');
  var lastRow = sheet.getLastRow();
  var existingRows = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, DETAIL_HEADERS.length).getValues() : [];
  var indexByKey = {};
  for (var i = 0; i < existingRows.length; i++) {
    var key = String(existingRows[i][keyIndex] || '');
    if (key) indexByKey[key] = i;
  }
  var added = 0;
  var updated = 0;
  for (var j = 0; j < incomingRows.length; j++) {
    var newKey = String(incomingRows[j][keyIndex] || '');
    if (!newKey) throw new Error('Refusing to write a row without Record Key.');
    if (Object.prototype.hasOwnProperty.call(indexByKey, newKey)) {
      existingRows[indexByKey[newKey]] = incomingRows[j];
      updated++;
    } else {
      indexByKey[newKey] = existingRows.length;
      existingRows.push(incomingRows[j]);
      added++;
    }
  }
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, DETAIL_HEADERS.length).clearContent();
  if (existingRows.length) sheet.getRange(2, 1, existingRows.length, DETAIL_HEADERS.length).setValues(existingRows);
  var finalRow = sheet.getLastRow();
  if (finalRow > 1) {
    sheet.getRange(2, 9, finalRow - 1, 1).setNumberFormat('yyyy-mm-dd');
    sheet.getRange(2, 16, finalRow - 1, 1).setNumberFormat('#,##0.00');
  }
  return { added: added, updated: updated };
}

function normalizeAccountId_(id) {
  return String(id || '').replace(/-/g, '');
}

function dateFromKey_(dateKey) {
  var parts = String(dateKey).split('-');
  return new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]), 12, 0, 0);
}

function round2_(value) {
  var number = Number(value);
  if (!isFinite(number)) number = 0;
  return Math.round((number + Number.EPSILON) * 100) / 100;
}
