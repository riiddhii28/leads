LEAD RETRIEVAL PROTOTYPE — FILES

1) Code.gs
   Paste into the Apps Script project attached to the new Google Sheet.
   Run setupPrototype() once. Confirm the warning. It creates exactly Master, Google, Meta
   and seeds synthetic DEMO rows. Do not re-run on a working/live spreadsheet because it resets those tabs.

   Optional Meta setup uses Script Properties under Apps Script > Project Settings > Script Properties:
     META_PAGE_ACCESS_TOKEN = Page access token for permitted Page/Form lead retrieval
     META_TEST_FORM_ID = form ID used in Meta Lead Ads Testing Tool
     META_PAGE_IDS = comma-separated Page IDs for live lead retrieval
     META_ADS_ACCESS_TOKEN = token with ads_read access for Insights reporting
     META_AD_ACCOUNT_IDS = comma-separated Meta ad account IDs, e.g. act_123456789

   Available functions:
     syncMetaTestLeads()            retrieves test lead records from /{FORM_ID}/test_leads
     syncMetaLiveLeadsFromPages()   retrieves live lead records from forms on configured Pages
     syncMetaAdsPerformance()       retrieves ad-level daily Meta spend/clicks/impressions for last 30 complete days
     rebuildMaster()                recalculates Master
     removeDemoData()               removes DEMO records only (confirmation required)

   TEST rows are deliberately excluded from Master totals. The live Meta lead function will import actual personal data; only use it when you intend to do so and have appropriate permission.

2) GoogleAdsScript.js
   This is a DIFFERENT script runtime. Do NOT paste it into the spreadsheet's Extensions > Apps Script project.
   Open the target Google Ads account > Tools > Bulk actions > Scripts, create a new script, paste it there,
   update CONFIG.SPREADSHEET_URL and CLIENT_ID, authorize, and run. It retrieves campaign-level daily metrics
   for the current Google Ads account and writes them to the existing Google tab.

   For a client with multiple Google Ads accounts, run it in each authorised account with the same spreadsheet URL
   and client ID, or later replace it with an MCC/manager-account variant.

CURRENCY NOTE
Master combines spends numerically, so all accounts for one client must use the same currency for now.
The bound Apps Script refuses to rebuild Master if it detects multiple currencies. Currency conversion should be
added before combining accounts with different currencies.

No credentials/tokens are included in the code. Never commit or share real access tokens.
