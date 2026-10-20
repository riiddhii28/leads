/**
 * Lead Retrieval Prototype — spreadsheet-bound Google Apps Script
 *
 * FIRST RUN: Run setupPrototype() only in a NEW, blank prototype spreadsheet.
 * It creates exactly three tabs (Master, Google, Meta), removes other tabs,
 * clears those three tabs, and seeds clearly labelled synthetic demo data.
 *
 * Live Meta functions require Script Properties (instructions supplied separately).
 * This script does NOT require an API key or advertising spend for the prototype.
 */

const APP_CONFIG = Object.freeze({
  CLIENT_ID: 'CLIENT_DEMO_001',
  CLIENT_NAME: 'Demo Client',
  GRAPH_API_VERSION: 'v26.0',
  TIME_ZONE: 'Asia/Kolkata',
  REPORTING_CURRENCY: 'INR'
});

const TAB_NAMES = Object.freeze({ MASTER: 'Master', GOOGLE: 'Google', META: 'Meta' });

const DETAIL_HEADERS = Object.freeze([
  'Record Type',
  'Client ID',
  'Client Name',
  'Platform',
  'Data Mode',
  'Account ID',
  'Account Name',
  'Currency',
  'Date',
  'Campaign ID',
  'Campaign Name',
  'Ad Group / Ad Set ID',
  'Ad Group / Ad Set Name',
  'Ad ID',
  'Ad Name',
  'Spend',
  'Clicks',
  'Impressions',
  'Reported Conversions',
  'Reported Leads',
  'Cost Per Lead',
  'Lead ID',
  'Lead Created At',
  'Lead Name',
  'Lead Email',
  'Lead Phone',
  'Lead Source',
  'Form ID',
  'Lead Status',
  'Record Origin',
  'Custom Fields JSON',
  'Record Key',
  'Synced At'
]);

const MASTER_HEADERS = Object.freeze([
  'Date',
  'Google Spend',
  'Google / IVR Leads',
  'Google Cost Per Lead',
  'Google Clicks',
  'Google Impressions',
  'Meta Spend',
  'Meta Leads',
  'Meta Cost Per Lead',
  'Meta Clicks',
  'Meta Impressions',
  'Total Spend',
  'Total Leads',
  'Overall CPL',
  'Last Updated'
]);

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Lead Retrieval')
    .addItem('Set up / reset demo prototype', 'setupPrototype')
    .addItem('Refresh Master summary', 'rebuildMaster')
    .addSeparator()
    .addItem('Sync Meta TEST lead from form', 'syncMetaTestLeads')
    .addItem('Sync Meta LIVE leads from Pages', 'syncMetaLiveLeadsFromPages')
    .addItem('Sync Meta Ads performance', 'syncMetaAdsPerformance')
    .addSeparator()
    .addItem('Remove demo seed data', 'removeDemoData')
    .addToUi();
}

/**
 * Run once in the new blank spreadsheet. This action is intentionally destructive
 * to other tabs/contents in this prototype file, so it asks for confirmation.
 */
function setupPrototype() {
  const ui = SpreadsheetApp.getUi();
  const answer = ui.alert(
    'Build the prototype spreadsheet?',
    'This will create/refresh Master, Google and Meta, remove any other tabs, and replace data in those tabs with synthetic DEMO data. Use only in your new blank prototype spreadsheet.',
    ui.ButtonSet.YES_NO
  );
  if (answer !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Open the target Google Sheet and run this script from Extensions → Apps Script.');
  ss.setSpreadsheetTimeZone(APP_CONFIG.TIME_ZONE);

  ensureExactlyThreeTabs_(ss);
  const master = ss.getSheetByName(TAB_NAMES.MASTER);
  const google = ss.getSheetByName(TAB_NAMES.GOOGLE);
  const meta = ss.getSheetByName(TAB_NAMES.META);

  ensureColumnCapacity_(master, MASTER_HEADERS.length);
  ensureColumnCapacity_(google, DETAIL_HEADERS.length);
  ensureColumnCapacity_(meta, DETAIL_HEADERS.length);

  [master, google, meta].forEach(function (sheet) {
    const existingFilter = sheet.getFilter();
    if (existingFilter) existingFilter.remove();
    sheet.clear();
  });

  master.getRange(1, 1, 1, MASTER_HEADERS.length).setValues([MASTER_HEADERS]);
  google.getRange(1, 1, 1, DETAIL_HEADERS.length).setValues([DETAIL_HEADERS]);
  meta.getRange(1, 1, 1, DETAIL_HEADERS.length).setValues([DETAIL_HEADERS]);

  styleMasterSheet_(master);
  styleDetailSheet_(google);
  styleDetailSheet_(meta);

  seedDemoData_(google, meta);
  rebuildMaster();

  ui.alert(
    'Prototype created',
    'Your spreadsheet now has exactly three tabs: Master, Google and Meta. The rows are synthetic DEMO records, not real advertising data or real people.',
    ui.ButtonSet.OK
  );
}

function ensureExactlyThreeTabs_(ss) {
  let master = ss.getSheetByName(TAB_NAMES.MASTER);
  if (!master) {
    const sheets = ss.getSheets();
    if (sheets.length > 0) {
      sheets[0].setName(TAB_NAMES.MASTER);
    } else {
      ss.insertSheet(TAB_NAMES.MASTER);
    }
  }
  if (!ss.getSheetByName(TAB_NAMES.GOOGLE)) ss.insertSheet(TAB_NAMES.GOOGLE);
  if (!ss.getSheetByName(TAB_NAMES.META)) ss.insertSheet(TAB_NAMES.META);

  ss.getSheets().forEach(function (sheet) {
    const name = sheet.getName();
    if ([TAB_NAMES.MASTER, TAB_NAMES.GOOGLE, TAB_NAMES.META].indexOf(name) === -1) {
      ss.deleteSheet(sheet);
    }
  });
}

function ensureColumnCapacity_(sheet, requiredColumns) {
  const currentColumns = sheet.getMaxColumns();
  if (currentColumns < requiredColumns) sheet.insertColumnsAfter(currentColumns, requiredColumns - currentColumns);
}

function styleMasterSheet_(sheet) {
  ensureColumnCapacity_(sheet, MASTER_HEADERS.length);
  sheet.setFrozenRows(1);
  sheet.setTabColor('#1a73e8');
  sheet.getRange(1, 1, 1, MASTER_HEADERS.length)
    .setBackground('#1f4e78').setFontColor('#ffffff').setFontWeight('bold').setWrap(true);
  sheet.setRowHeight(1, 38);
  sheet.setColumnWidth(1, 110);
  sheet.setColumnWidths(2, 1, 130);
  sheet.setColumnWidths(3, 1, 150);
  sheet.setColumnWidths(4, 1, 145);
  sheet.setColumnWidths(5, 2, 125);
  sheet.setColumnWidths(7, 1, 130);
  sheet.setColumnWidths(8, 1, 100);
  sheet.setColumnWidths(9, 1, 130);
  sheet.setColumnWidths(10, 5, 125);
  if (sheet.getFilter()) sheet.getFilter().remove();
  sheet.getRange(1, 1, Math.max(1, sheet.getLastRow()), MASTER_HEADERS.length).createFilter();
}

function styleDetailSheet_(sheet) {
  ensureColumnCapacity_(sheet, DETAIL_HEADERS.length);
  sheet.setFrozenRows(1);
  sheet.setTabColor(sheet.getName() === TAB_NAMES.GOOGLE ? '#188038' : '#1877f2');
  sheet.getRange(1, 1, 1, DETAIL_HEADERS.length)
    .setBackground('#1f4e78').setFontColor('#ffffff').setFontWeight('bold').setWrap(true);
  sheet.setRowHeight(1, 42);
  sheet.setColumnWidths(1, DETAIL_HEADERS.length, 135);
  sheet.setColumnWidth(3, 165);
  sheet.setColumnWidth(7, 190);
  sheet.setColumnWidth(11, 190);
  sheet.setColumnWidth(13, 190);
  sheet.setColumnWidth(15, 190);
  sheet.setColumnWidth(24, 180);
  sheet.setColumnWidth(25, 220);
  sheet.setColumnWidth(27, 230);
  sheet.setColumnWidth(31, 260);
  sheet.setColumnWidth(32, 300);
  if (sheet.getFilter()) sheet.getFilter().remove();
  sheet.getRange(1, 1, Math.max(1, sheet.getLastRow()), DETAIL_HEADERS.length).createFilter();
}

function seedDemoData_(googleSheet, metaSheet) {
  const googleRows = [];
  const metaRows = [];
  let fakeLeadNumber = 0;

  const demoConfigs = [
    {
      platform: 'Google', sheetRows: googleRows, accountId: 'DEMO-G-ACCOUNT-001',
      accountName: 'Demo Google Account A', campaignId: 'DEMO-G-CAMP-001',
      campaignName: 'Search Leads — Demo', groupId: 'DEMO-G-AG-001',
      groupName: 'Search Ad Group — Demo', adId: 'DEMO-G-AD-001', adName: 'Search Ad — Demo',
      leadSource: 'Google-attributed CRM / IVR (sample)',
      daily: [
        { daysAgo: 2, spend: 4200, clicks: 82, impressions: 4200, leads: 19 },
        { daysAgo: 1, spend: 4600, clicks: 91, impressions: 4800, leads: 22 },
        { daysAgo: 0, spend: 5000, clicks: 104, impressions: 5200, leads: 25 }
      ]
    },
    {
      platform: 'Google', sheetRows: googleRows, accountId: 'DEMO-G-ACCOUNT-002',
      accountName: 'Demo Google Account B', campaignId: 'DEMO-G-CAMP-002',
      campaignName: 'Brand Search — Demo', groupId: 'DEMO-G-AG-002',
      groupName: 'Brand Ad Group — Demo', adId: 'DEMO-G-AD-002', adName: 'Brand Ad — Demo',
      leadSource: 'Google-attributed CRM / IVR (sample)',
      daily: [
        { daysAgo: 2, spend: 1400, clicks: 35, impressions: 1500, leads: 5 },
        { daysAgo: 1, spend: 1550, clicks: 38, impressions: 1650, leads: 6 },
        { daysAgo: 0, spend: 1600, clicks: 40, impressions: 1700, leads: 6 }
      ]
    },
    {
      platform: 'Meta', sheetRows: metaRows, accountId: 'act_DEMO_META_001',
      accountName: 'Demo Meta Account A', campaignId: 'DEMO-M-CAMP-001',
      campaignName: 'Instant Form Leads — Demo', groupId: 'DEMO-M-SET-001',
      groupName: 'Lead Generation Ad Set — Demo', adId: 'DEMO-M-AD-001', adName: 'Instant Form Ad — Demo',
      leadSource: 'Meta Instant Form (sample)', formId: 'DEMO-M-FORM-001',
      daily: [
        { daysAgo: 2, spend: 3000, clicks: 90, impressions: 4000, leads: 20 },
        { daysAgo: 1, spend: 3200, clicks: 95, impressions: 4150, leads: 21 },
        { daysAgo: 0, spend: 3300, clicks: 100, impressions: 4300, leads: 23 }
      ]
    },
    {
      platform: 'Meta', sheetRows: metaRows, accountId: 'act_DEMO_META_002',
      accountName: 'Demo Meta Account B', campaignId: 'DEMO-M-CAMP-002',
      campaignName: 'Retargeting — Demo', groupId: 'DEMO-M-SET-002',
      groupName: 'Retargeting Ad Set — Demo', adId: 'DEMO-M-AD-002', adName: 'Retargeting Ad — Demo',
      leadSource: 'Meta Instant Form (sample)', formId: 'DEMO-M-FORM-002',
      daily: [
        { daysAgo: 2, spend: 1000, clicks: 22, impressions: 1100, leads: 3 },
        { daysAgo: 1, spend: 1200, clicks: 26, impressions: 1250, leads: 4 },
        { daysAgo: 0, spend: 1300, clicks: 28, impressions: 1400, leads: 5 }
      ]
    }
  ];

  demoConfigs.forEach(function (cfg) {
    cfg.daily.forEach(function (day) {
      const dateKey = dateKeyDaysAgo_(day.daysAgo);
      const currency = APP_CONFIG.REPORTING_CURRENCY;
      const perf = detailRow_({
        'Record Type': 'PERFORMANCE', 'Client ID': APP_CONFIG.CLIENT_ID,
        'Client Name': APP_CONFIG.CLIENT_NAME, 'Platform': cfg.platform, 'Data Mode': 'DEMO',
        'Account ID': cfg.accountId, 'Account Name': cfg.accountName, 'Currency': currency,
        'Date': dateObjectFromKey_(dateKey), 'Campaign ID': cfg.campaignId,
        'Campaign Name': cfg.campaignName, 'Ad Group / Ad Set ID': cfg.groupId,
        'Ad Group / Ad Set Name': cfg.groupName, 'Ad ID': cfg.adId, 'Ad Name': cfg.adName,
        'Spend': day.spend, 'Clicks': day.clicks, 'Impressions': day.impressions,
        'Reported Conversions': cfg.platform === 'Google' ? day.leads : '',
        'Reported Leads': cfg.platform === 'Meta' ? day.leads : day.leads,
        'Cost Per Lead': day.leads > 0 ? day.spend / day.leads : '',
        'Record Origin': 'SYNTHETIC_DEMO_DATA',
        'Custom Fields JSON': JSON.stringify({ demo: true, note: 'Fictional metrics for prototype testing only.' }),
        'Record Key': cfg.platform.toUpperCase() + '|DEMO|PERFORMANCE|' + dateKey + '|' + cfg.accountId + '|' + cfg.campaignId + '|' + cfg.adId,
        'Synced At': new Date()
      });
      cfg.sheetRows.push(perf);

      // Create one synthetic lead row per sample lead so the Master count demonstrates
      // how actual CRM/IVR or Meta Instant Form lead records are counted.
      for (let i = 1; i <= day.leads; i++) {
        fakeLeadNumber++;
        const leadId = 'DEMO-' + cfg.platform.toUpperCase() + '-LEAD-' + dateKey.replace(/-/g, '') + '-' + cfg.campaignId + '-' + i;
        const lead = detailRow_({
          'Record Type': 'LEAD', 'Client ID': APP_CONFIG.CLIENT_ID,
          'Client Name': APP_CONFIG.CLIENT_NAME, 'Platform': cfg.platform, 'Data Mode': 'DEMO',
          'Account ID': cfg.accountId, 'Account Name': cfg.accountName, 'Currency': currency,
          'Date': dateObjectFromKey_(dateKey), 'Campaign ID': cfg.campaignId,
          'Campaign Name': cfg.campaignName, 'Ad Group / Ad Set ID': cfg.groupId,
          'Ad Group / Ad Set Name': cfg.groupName, 'Ad ID': cfg.adId, 'Ad Name': cfg.adName,
          'Lead ID': leadId, 'Lead Created At': dateKey + 'T12:00:00+05:30',
          'Lead Name': 'Sample ' + cfg.platform + ' Lead ' + fakeLeadNumber,
          'Lead Email': 'sample-' + cfg.platform.toLowerCase() + '-' + String(fakeLeadNumber).padStart(4, '0') + '@example.test',
          'Lead Phone': 'TEST-ONLY-' + String(fakeLeadNumber).padStart(4, '0'),
          'Lead Source': cfg.leadSource, 'Form ID': cfg.formId || '',
          'Lead Status': 'SAMPLE', 'Record Origin': 'SYNTHETIC_DEMO_DATA',
          'Custom Fields JSON': JSON.stringify({ demo: true, sample_field: 'Example value ' + fakeLeadNumber }),
          'Record Key': cfg.platform.toUpperCase() + '|DEMO|LEAD|' + leadId,
          'Synced At': new Date()
        });
        cfg.sheetRows.push(lead);
      }
    });
  });

  writeDetailRows_(googleSheet, googleRows);
  writeDetailRows_(metaSheet, metaRows);
}

function detailRow_(overrides) {
  const record = {};
  DETAIL_HEADERS.forEach(function (header) { record[header] = ''; });
  record['Client ID'] = APP_CONFIG.CLIENT_ID;
  record['Client Name'] = APP_CONFIG.CLIENT_NAME;
  record['Currency'] = APP_CONFIG.REPORTING_CURRENCY;
  record['Synced At'] = new Date();
  Object.keys(overrides || {}).forEach(function (key) { record[key] = overrides[key]; });
  return DETAIL_HEADERS.map(function (header) {
    return record[header] === undefined || record[header] === null ? '' : record[header];
  });
}

function writeDetailRows_(sheet, rows) {
  if (rows.length) sheet.getRange(2, 1, rows.length, DETAIL_HEADERS.length).setValues(rows);
  formatDetailData_(sheet);
}

function formatDetailData_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  sheet.getRange(2, 9, lastRow - 1, 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(2, 16, lastRow - 1, 1).setNumberFormat('#,##0.00');
  sheet.getRange(2, 21, lastRow - 1, 1).setNumberFormat('#,##0.00');
  sheet.getRange(2, 33, lastRow - 1, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  if (sheet.getFilter()) sheet.getFilter().remove();
  sheet.getRange(1, 1, lastRow, DETAIL_HEADERS.length).createFilter();
}

/** Rebuild the daily client summary from PERFORMANCE and LEAD rows. */
function rebuildMaster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName(TAB_NAMES.MASTER);
  const google = ss.getSheetByName(TAB_NAMES.GOOGLE);
  const meta = ss.getSheetByName(TAB_NAMES.META);
  if (!master || !google || !meta) throw new Error('Missing one of the required tabs. Run setupPrototype() in the prototype spreadsheet first.');

  const totals = {};
  const currencies = new Set();
  const index = {};
  DETAIL_HEADERS.forEach(function (h, i) { index[h] = i; });

  [google, meta].forEach(function (sheet) {
    const platformData = sheet.getDataRange().getValues();
    if (platformData.length < 2) return;
    platformData.slice(1).forEach(function (row) {
      const type = String(row[index['Record Type']] || '').trim().toUpperCase();
      const mode = String(row[index['Data Mode']] || '').trim().toUpperCase();
      if (mode !== 'DEMO' && mode !== 'LIVE') return; // TEST leads must not affect reported results.

      const platform = String(row[index['Platform']] || sheet.getName()).trim();
      let dateKey = dateKeyFromCell_(row[index['Date']]);
      if (!dateKey && row[index['Lead Created At']]) dateKey = dateKeyFromCell_(row[index['Lead Created At']]);
      if (!dateKey) return;

      if (!totals[dateKey]) {
        totals[dateKey] = {
          googleSpend: 0, googleClicks: 0, googleImpressions: 0, googleLeads: 0,
          metaSpend: 0, metaClicks: 0, metaImpressions: 0, metaLeads: 0,
          googleLeadIds: new Set(), metaLeadIds: new Set()
        };
      }
      const day = totals[dateKey];

      if (type === 'PERFORMANCE') {
        const currency = String(row[index['Currency']] || '').trim().toUpperCase();
        if (currency && currency !== 'UNKNOWN') currencies.add(currency);
        const spend = numberOrZero_(row[index['Spend']]);
        const clicks = numberOrZero_(row[index['Clicks']]);
        const impressions = numberOrZero_(row[index['Impressions']]);
        if (platform === 'Google') {
          day.googleSpend += spend;
          day.googleClicks += clicks;
          day.googleImpressions += impressions;
        } else if (platform === 'Meta') {
          day.metaSpend += spend;
          day.metaClicks += clicks;
          day.metaImpressions += impressions;
        }
      } else if (type === 'LEAD') {
        const leadId = String(row[index['Lead ID']] || row[index['Record Key']] || '').trim();
        if (!leadId) return;
        if (platform === 'Google') day.googleLeadIds.add(leadId);
        else if (platform === 'Meta') day.metaLeadIds.add(leadId);
      }
    });
  });

  if (currencies.size > 1) {
    throw new Error('Master was not refreshed because performance rows use multiple currencies (' + Array.from(currencies).join(', ') + '). Convert amounts to one client reporting currency before combining them.');
  }

  const dateKeys = Object.keys(totals).sort();
  const now = new Date();
  const out = dateKeys.map(function (dateKey) {
    const day = totals[dateKey];
    const googleLeads = day.googleLeadIds.size;
    const metaLeads = day.metaLeadIds.size;
    const totalSpend = day.googleSpend + day.metaSpend;
    const totalLeads = googleLeads + metaLeads;
    return [
      dateObjectFromKey_(dateKey),
      round2_(day.googleSpend), googleLeads, googleLeads ? round2_(day.googleSpend / googleLeads) : '',
      day.googleClicks, day.googleImpressions,
      round2_(day.metaSpend), metaLeads, metaLeads ? round2_(day.metaSpend / metaLeads) : '',
      day.metaClicks, day.metaImpressions,
      round2_(totalSpend), totalLeads, totalLeads ? round2_(totalSpend / totalLeads) : '', now
    ];
  });

  if (master.getFilter()) master.getFilter().remove();
  master.clearContents();
  master.getRange(1, 1, 1, MASTER_HEADERS.length).setValues([MASTER_HEADERS]);
  if (out.length) master.getRange(2, 1, out.length, MASTER_HEADERS.length).setValues(out);
  styleMasterSheet_(master);
  const lastRow = master.getLastRow();
  if (lastRow > 1) {
    master.getRange(2, 1, lastRow - 1, 1).setNumberFormat('yyyy-mm-dd');
    [2, 4, 7, 9, 12, 14].forEach(function (col) {
      master.getRange(2, col, lastRow - 1, 1).setNumberFormat('₹#,##0.00');
    });
    [3, 5, 6, 8, 10, 11, 13].forEach(function (col) {
      master.getRange(2, col, lastRow - 1, 1).setNumberFormat('#,##0');
    });
    master.getRange(2, 15, lastRow - 1, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  }
  SpreadsheetApp.flush();
}

/** Pull one test lead from Meta's form test_leads edge. TEST rows are excluded from Master totals. */
function syncMetaTestLeads() {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('META_PAGE_ACCESS_TOKEN');
  const formId = props.getProperty('META_TEST_FORM_ID');
  if (!token || !formId) {
    throw new Error('Add META_PAGE_ACCESS_TOKEN and META_TEST_FORM_ID in Apps Script → Project Settings → Script Properties, then run again.');
  }

  const leads = graphGetAll_(formId + '/test_leads', {
    fields: 'id,created_time,ad_id,form_id,field_data',
    limit: '100'
  }, token);
  const rows = leads.map(function (lead) { return metaLeadToRow_(lead, 'TEST', 'META_TEST_LEAD_API'); });
  const result = upsertRows_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB_NAMES.META), rows);
  rebuildMaster();
  SpreadsheetApp.getActiveSpreadsheet().toast('Meta test leads: ' + leads.length + ' fetched; ' + result.added + ' added, ' + result.updated + ' updated. TEST rows do not count in Master.', 'Lead Retrieval', 8);
}

/**
 * Import live Meta Instant Form leads from all configured Pages.
 * Only run after the Page token and permissions are configured and you intend to import real leads.
 */
function syncMetaLiveLeadsFromPages() {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('META_PAGE_ACCESS_TOKEN');
  const pageIds = (props.getProperty('META_PAGE_IDS') || '').split(',').map(function (x) { return x.trim(); }).filter(Boolean);
  if (!token || !pageIds.length) {
    throw new Error('Add META_PAGE_ACCESS_TOKEN and META_PAGE_IDS (comma-separated Page IDs) in Script Properties first.');
  }

  const rows = [];
  const pendingCheckpoints = {};
  pageIds.forEach(function (pageId) {
    const forms = graphGetAll_(pageId + '/leadgen_forms', { fields: 'id,name,status', limit: '100' }, token);
    forms.forEach(function (form) {
      const propKey = 'META_LEAD_LAST_SYNC_' + form.id;
      const lastSyncedUnix = Number(props.getProperty(propKey) || 0);
      const params = { fields: 'id,created_time,ad_id,form_id,field_data', limit: '100' };
      if (lastSyncedUnix > 0) {
        params.filtering = JSON.stringify([{ field: 'time_created', operator: 'GREATER_THAN', value: Math.max(0, lastSyncedUnix - 60) }]);
      }
      const leads = graphGetAll_(form.id + '/leads', params, token);
      let maxCreatedUnix = lastSyncedUnix;
      leads.forEach(function (lead) {
        const createdUnix = lead.created_time ? Math.floor(new Date(lead.created_time).getTime() / 1000) : 0;
        if (createdUnix > maxCreatedUnix) maxCreatedUnix = createdUnix;
        rows.push(metaLeadToRow_(lead, 'LIVE', 'META_LEAD_ADS_API', form.name || ''));
      });
      if (maxCreatedUnix > lastSyncedUnix) pendingCheckpoints[propKey] = String(Math.max(0, maxCreatedUnix - 60));
    });
  });

  const result = upsertRows_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB_NAMES.META), rows);
  Object.keys(pendingCheckpoints).forEach(function (key) { props.setProperty(key, pendingCheckpoints[key]); });
  rebuildMaster();
  SpreadsheetApp.getActiveSpreadsheet().toast('Meta live leads: ' + rows.length + ' fetched; ' + result.added + ' added, ' + result.updated + ' updated.', 'Lead Retrieval', 8);
}

/** Pull daily ad-level Meta performance for the configured accounts and last 30 complete days. */
function syncMetaAdsPerformance() {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('META_ADS_ACCESS_TOKEN');
  const accountIds = (props.getProperty('META_AD_ACCOUNT_IDS') || '').split(',').map(function (x) { return x.trim(); }).filter(Boolean);
  if (!token || !accountIds.length) {
    throw new Error('Add META_ADS_ACCESS_TOKEN and META_AD_ACCOUNT_IDS (comma-separated act_... IDs) in Script Properties first.');
  }

  const range = lastCompleteDaysRange_(30);
  const incoming = [];
  const fields = 'account_id,account_name,account_currency,campaign_id,campaign_name,adset_id,adset_name,ad_id,ad_name,date_start,spend,clicks,impressions,actions';

  accountIds.forEach(function (accountId) {
    const reportRows = graphGetAll_(accountId + '/insights', {
      level: 'ad',
      time_increment: '1',
      time_range: JSON.stringify({ since: range.since, until: range.until }),
      fields: fields,
      limit: '500'
    }, token);

    reportRows.forEach(function (report) {
      const dateKey = report.date_start || range.until;
      const reportedLeads = getMetaReportedLeadCount_(report.actions);
      incoming.push(detailRow_({
        'Record Type': 'PERFORMANCE', 'Client ID': APP_CONFIG.CLIENT_ID,
        'Client Name': APP_CONFIG.CLIENT_NAME, 'Platform': 'Meta', 'Data Mode': 'LIVE',
        'Account ID': report.account_id || accountId,
        'Account Name': report.account_name || '', 'Currency': report.account_currency || 'UNKNOWN',
        'Date': dateObjectFromKey_(dateKey), 'Campaign ID': report.campaign_id || '',
        'Campaign Name': report.campaign_name || '',
        'Ad Group / Ad Set ID': report.adset_id || '',
        'Ad Group / Ad Set Name': report.adset_name || '',
        'Ad ID': report.ad_id || '', 'Ad Name': report.ad_name || '',
        'Spend': numberOrZero_(report.spend), 'Clicks': numberOrZero_(report.clicks),
        'Impressions': numberOrZero_(report.impressions), 'Reported Leads': reportedLeads,
        'Record Origin': 'META_MARKETING_API_INSIGHTS',
        'Custom Fields JSON': JSON.stringify({ actions: report.actions || [] }),
        'Record Key': 'META|PERFORMANCE|' + dateKey + '|' + (report.account_id || accountId) + '|' + (report.campaign_id || '') + '|' + (report.ad_id || ''),
        'Synced At': new Date()
      }));
    });
  });

  replacePerformanceRange_(
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB_NAMES.META),
    'Meta', accountIds, range.since, range.until, incoming
  );
  rebuildMaster();
  SpreadsheetApp.getActiveSpreadsheet().toast('Meta performance rows refreshed for ' + range.since + ' to ' + range.until + ': ' + incoming.length + ' rows.', 'Lead Retrieval', 8);
}

/** Uses one explicit action type to avoid adding overlapping Meta lead action types together. */
function getMetaReportedLeadCount_(actions) {
  if (!Array.isArray(actions)) return '';
  let action = actions.find(function (item) { return item.action_type === 'lead'; });
  if (!action) action = actions.find(function (item) { return item.action_type === 'onsite_conversion.lead_grouped'; });
  return action ? numberOrZero_(action.value) : '';
}

function metaLeadToRow_(lead, dataMode, origin, formName) {
  const customFields = {};
  (lead.field_data || []).forEach(function (field) {
    const values = Array.isArray(field.values) ? field.values : [];
    customFields[field.name] = values.join(', ');
  });
  const firstName = customFields.first_name || '';
  const lastName = customFields.last_name || '';
  const name = customFields.full_name || [firstName, lastName].filter(Boolean).join(' ');
  const created = lead.created_time || '';
  const dateKey = created ? Utilities.formatDate(new Date(created), APP_CONFIG.TIME_ZONE, 'yyyy-MM-dd') : dateKeyDaysAgo_(0);
  const leadId = String(lead.id || '');
  return detailRow_({
    'Record Type': 'LEAD', 'Client ID': APP_CONFIG.CLIENT_ID,
    'Client Name': APP_CONFIG.CLIENT_NAME, 'Platform': 'Meta', 'Data Mode': dataMode,
    'Account ID': '', 'Account Name': '', 'Currency': '', 'Date': dateObjectFromKey_(dateKey),
    'Ad ID': lead.ad_id || '', 'Lead ID': leadId, 'Lead Created At': created,
    'Lead Name': name, 'Lead Email': customFields.email || '',
    'Lead Phone': customFields.phone_number || customFields.phone || '',
    'Lead Source': dataMode === 'TEST' ? 'Meta Instant Form TEST' : 'Meta Instant Form',
    'Form ID': lead.form_id || '', 'Lead Status': dataMode === 'TEST' ? 'TEST' : 'NEW',
    'Record Origin': origin || 'META_LEAD_ADS_API',
    'Custom Fields JSON': JSON.stringify({ formName: formName || '', fields: customFields }),
    'Record Key': 'META|' + dataMode + '|LEAD|' + leadId,
    'Synced At': new Date()
  });
}

function graphGetAll_(path, params, token) {
  const baseUrl = 'https://graph.facebook.com/' + APP_CONFIG.GRAPH_API_VERSION + '/' + path.replace(/^\//, '');
  const query = Object.keys(params || {}).map(function (key) {
    return encodeURIComponent(key) + '=' + encodeURIComponent(String(params[key]));
  }).join('&');
  let nextUrl = baseUrl + (query ? '?' + query : '');
  const all = [];
  let pageCount = 0;

  while (nextUrl && pageCount < 200) {
    const response = UrlFetchApp.fetch(nextUrl, {
      method: 'get',
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });
    const status = response.getResponseCode();
    const body = response.getContentText();
    let json;
    try { json = JSON.parse(body); } catch (e) { throw new Error('Meta returned an unreadable response (HTTP ' + status + ').'); }
    if (status < 200 || status >= 300 || json.error) {
      const message = json.error && json.error.message ? json.error.message : body.slice(0, 500);
      throw new Error('Meta API error (HTTP ' + status + '): ' + message);
    }
    if (Array.isArray(json.data)) all.push.apply(all, json.data);
    nextUrl = json.paging && json.paging.next ? json.paging.next : '';
    pageCount++;
  }
  if (nextUrl) throw new Error('Meta pagination exceeded 200 pages. Narrow the date range or sync in smaller batches.');
  return all;
}

function upsertRows_(sheet, incomingRows) {
  if (!sheet) throw new Error('The target sheet does not exist. Run setupPrototype() first.');
  if (!incomingRows.length) return { added: 0, updated: 0 };
  const keyIndex = DETAIL_HEADERS.indexOf('Record Key');
  const lastRow = sheet.getLastRow();
  const existingRows = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, DETAIL_HEADERS.length).getValues() : [];
  const rowIndexByKey = new Map();
  existingRows.forEach(function (row, i) {
    const key = String(row[keyIndex] || '').trim();
    if (key) rowIndexByKey.set(key, i);
  });
  let added = 0;
  let updated = 0;
  incomingRows.forEach(function (row) {
    const key = String(row[keyIndex] || '').trim();
    if (!key) throw new Error('A record is missing Record Key; refusing to write a potentially duplicated row.');
    if (rowIndexByKey.has(key)) {
      existingRows[rowIndexByKey.get(key)] = row;
      updated++;
    } else {
      rowIndexByKey.set(key, existingRows.length);
      existingRows.push(row);
      added++;
    }
  });
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, DETAIL_HEADERS.length).clearContent();
  if (existingRows.length) sheet.getRange(2, 1, existingRows.length, DETAIL_HEADERS.length).setValues(existingRows);
  formatDetailData_(sheet);
  return { added: added, updated: updated };
}

/** Replace live PERFORMANCE rows for the specified accounts/date window; preserves DEMO and LEAD rows. */
function replacePerformanceRange_(sheet, platform, accountIds, since, until, incomingRows) {
  const idSet = new Set(accountIds.map(function (id) { return String(id).replace(/^act_/, '').replace(/-/g, ''); }));
  const lastRow = sheet.getLastRow();
  const allRows = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, DETAIL_HEADERS.length).getValues() : [];
  const ix = {};
  DETAIL_HEADERS.forEach(function (h, i) { ix[h] = i; });
  const keepRows = allRows.filter(function (row) {
    const sameTarget = String(row[ix['Record Type']]).toUpperCase() === 'PERFORMANCE' &&
      String(row[ix['Platform']]) === platform && String(row[ix['Data Mode']]).toUpperCase() === 'LIVE' &&
      idSet.has(String(row[ix['Account ID']]).replace(/^act_/, '').replace(/-/g, ''));
    if (!sameTarget) return true;
    const key = dateKeyFromCell_(row[ix['Date']]);
    return !(key && key >= since && key <= until);
  });
  const map = new Map();
  keepRows.forEach(function (row, i) {
    const key = String(row[ix['Record Key']] || '');
    if (key) map.set(key, i);
  });
  incomingRows.forEach(function (row) {
    const key = String(row[ix['Record Key']] || '');
    if (!key) throw new Error('A performance row is missing Record Key.');
    if (map.has(key)) keepRows[map.get(key)] = row;
    else { map.set(key, keepRows.length); keepRows.push(row); }
  });
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, DETAIL_HEADERS.length).clearContent();
  if (keepRows.length) sheet.getRange(2, 1, keepRows.length, DETAIL_HEADERS.length).setValues(keepRows);
  formatDetailData_(sheet);
}

/** Remove only rows generated by setupPrototype(); TEST and LIVE records are preserved. */
function removeDemoData() {
  const ui = SpreadsheetApp.getUi();
  const answer = ui.alert(
    'Remove synthetic demo rows?',
    'This will permanently remove rows marked DEMO from Google and Meta and rebuild Master. TEST and LIVE rows will be kept.',
    ui.ButtonSet.YES_NO
  );
  if (answer !== ui.Button.YES) return;
  [TAB_NAMES.GOOGLE, TAB_NAMES.META].forEach(function (tabName) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tabName);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    const rows = sheet.getRange(2, 1, lastRow - 1, DETAIL_HEADERS.length).getValues();
    const modeIndex = DETAIL_HEADERS.indexOf('Data Mode');
    const keep = rows.filter(function (row) { return String(row[modeIndex]).toUpperCase() !== 'DEMO'; });
    sheet.getRange(2, 1, lastRow - 1, DETAIL_HEADERS.length).clearContent();
    if (keep.length) sheet.getRange(2, 1, keep.length, DETAIL_HEADERS.length).setValues(keep);
    formatDetailData_(sheet);
  });
  rebuildMaster();
  ui.alert('Demo rows removed. Any TEST rows remain in Meta but do not count in Master.');
}

function lastCompleteDaysRange_(numberOfDays) {
  const today = new Date();
  today.setHours(12, 0, 0, 0);
  const end = new Date(today);
  end.setDate(end.getDate() - 1);
  const start = new Date(end);
  start.setDate(start.getDate() - numberOfDays + 1);
  return {
    since: Utilities.formatDate(start, APP_CONFIG.TIME_ZONE, 'yyyy-MM-dd'),
    until: Utilities.formatDate(end, APP_CONFIG.TIME_ZONE, 'yyyy-MM-dd')
  };
}

function dateKeyDaysAgo_(daysAgo) {
  const date = new Date();
  date.setHours(12, 0, 0, 0);
  date.setDate(date.getDate() - daysAgo);
  return Utilities.formatDate(date, APP_CONFIG.TIME_ZONE, 'yyyy-MM-dd');
}

function dateObjectFromKey_(dateKey) {
  const parts = String(dateKey).slice(0, 10).split('-');
  return new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]), 12, 0, 0);
}

function dateKeyFromCell_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, APP_CONFIG.TIME_ZONE, 'yyyy-MM-dd');
  }
  const text = String(value).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(text)) return text.slice(0, 10);
  const parsed = new Date(text);
  if (!isNaN(parsed.getTime())) return Utilities.formatDate(parsed, APP_CONFIG.TIME_ZONE, 'yyyy-MM-dd');
  return '';
}

function numberOrZero_(value) {
  const n = Number(value);
  return isFinite(n) ? n : 0;
}

function round2_(value) {
  return Math.round((numberOrZero_(value) + Number.EPSILON) * 100) / 100;
}
