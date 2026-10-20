Lead Retrieval & Advertising Reporting Prototype — Handoff
Purpose
Build a reusable lead-retrieval and advertising-reporting prototype for multiple clients. Each client has one Google Sheets spreadsheet with exactly three tabs:
1. Master — daily summary across all Google Ads and Meta Ads accounts connected to that client.
2. Google — detailed Google advertising performance rows and Google-attributed lead records.
3. Meta — detailed Meta advertising performance rows and Meta lead records.
Google Apps Script is used for spreadsheet setup and Meta integration. Google Ads reporting should use a separate Google Ads Script because AdsApp runs inside the Google Ads scripting environment, not in a spreadsheet-bound Apps Script project. The Google Ads Script can write into the same spreadsheet with SpreadsheetApp.
Current status
- The intended Drive folder name is Lead Retrieval Prototype.
- The intended spreadsheet name is CLIENT_DEMO_001_Lead_Reporting.
- A code package was previously provided with:
  - Code.gs — spreadsheet setup, synthetic sample data, Master summary logic, and Meta integration functions.
  - GoogleAdsScript.js — Google Ads Scripts reporting connector.
  - README.txt — setup notes.
- The user has not confirmed that setupPrototype() ran successfully. Do not assume the sheet is configured.
- Meta app, Page/form IDs, access token and permissions have not been confirmed/configured.
- Google Ads Script has not been configured against a real account.
- Looker Studio has not been built.
- No real client accounts have been connected.
- The scripts have not been validated against live Meta or Google Ads accounts. Do not claim production readiness.
Expected spreadsheet design
Master
One summary row per date, with fields such as:
- Date
- Google Spend
- Google / IVR Leads
- Google CPL
- Google Clicks
- Google Impressions
- Meta Spend
- Meta Leads
- Meta CPL
- Meta Clicks
- Meta Impressions
- Total Spend
- Total Leads
- Overall CPL
- Last Updated
Check the actual Code.gs before relying on exact headers or function names.
Google
One tab for all Google Ads accounts connected to this client. Keep account ID/name so multiple accounts can coexist. Include performance rows and individual lead rows, differentiated by a Record Type field (e.g. PERFORMANCE, LEAD) to prevent spend being multiplied by the number of leads. Expected details include date, account/campaign/ad group/ad IDs, spend, clicks, impressions, conversions, available lead details, attribution and record origin.
Meta
One tab for all Meta ad accounts connected to this client. Include account ID/name and campaign/ad set/ad/form IDs where available. Use Record Type to distinguish performance rows from individual lead rows. Include lead IDs, creation time, authorised contact fields, source, attribution, status and client-specific fields where available.

Do not add a fourth visible tab or create separate tabs per ad account unless the user changes this requirement. Credentials should not be stored in visible sheet cells or public GitHub.
Data rules
- Each client gets a separate spreadsheet; do not mix clients' leads or credentials.
- One client can have multiple Google Ads accounts and multiple Meta ad accounts.
- Do not assume campaign IDs are globally unique; use account ID with campaign ID and date/reporting level where appropriate.
- Performance rows and individual lead rows have different data granularity. Never sum spend across lead rows.
- Keep platform-reported conversions separate from actual leads from CRM, forms or IVR.
- Define the Master tab's lead-count and attribution rules before calculating CPL.
- Use stable source IDs to avoid duplicates; update existing records on subsequent syncs.
- Mark test records and exclude them from production totals.
- Preserve source currency; do not combine different currencies without conversion.
- Do not invent missing attribution; mark it unknown/unattributed.
- Protect lead PII and tokens. Never commit access tokens or client data to GitHub.
Prototype approach
Start with synthetic data, then validate:
1. Exactly three tabs and expected headers.
2. Master daily totals and CPL, including zero-lead cases.
3. Refresh/idempotency: rerunning syncs must not duplicate records.
4. Correct distinction between performance rows and lead rows.
5. Meta test-lead retrieval.
6. Google reporting approach and dashboard.
7. Only then, real client integrations.
Google Ads API test accounts are useful for API/configuration testing but cannot serve ads, so they do not provide real spend, impression or conversion performance. Use synthetic data for those metrics during the free prototype.
Meta integration requirements
Use the Meta Lead Ads Testing Tool:
https://developers.facebook.com/tools/lead-ads-testing

A test lead will not automatically appear in Sheets. The integration needs an authorised Meta app/token, Page and lead-form access, the current required permissions, and the relevant Page/form IDs. Confirm the latest requirements in official Meta docs before implementation. Store secrets in Apps Script Properties where appropriate, restrict editor access, and never hard-code or commit tokens.

Meta Ads performance insights and individual lead-form retrieval are separate capabilities and may require different permissions. A test lead does not create real ad spend, clicks or impressions.
Google Ads integration requirements
GoogleAdsScript.js should run in the Google Ads Scripts environment, not in the spreadsheet-bound Apps Script editor. Configure the target spreadsheet ID and write to the Google tab. Confirm GAQL fields, reporting level, date range, account selection and output headers before running. For manager accounts, explicitly handle child-account iteration and access.

Official references:
- Google Ads Scripts reporting: https://developers.google.com/google-ads/scripts/docs/concepts/reports
- Google Ads API test accounts: https://developers.google.com/google-ads/api/docs/best-practices/test-accounts
- Google Ads API access levels: https://developers.google.com/google-ads/api/docs/api-policy/access-levels
- Meta Lead Ads Testing Tool: https://developers.facebook.com/tools/lead-ads-testing
Remaining work — recommended order
Phase 1: Confirm setup
- [ ] Add the existing Code.gs, GoogleAdsScript.js, and README.txt to the repository.
- [ ] Inspect actual files and align this document with real function names and column headers.
- [ ] Ask whether the user has created the spreadsheet and run setupPrototype().
- [ ] If it failed, debug the exact error before proceeding.
Phase 2: Validate prototype
- [ ] Confirm exactly Master, Google, and Meta.
- [ ] Validate headers, formatting, sample records and IDs.
- [ ] Manually verify daily rollups, CPL and zero-lead behavior.
- [ ] Ensure test leads and performance rows aggregate correctly.
- [ ] Ensure repeated syncs update records instead of duplicating them.
- [ ] Consider timezone, currencies, API pagination/limits, error logging and retries.
Phase 3: Meta test lead
- [ ] Verify the user has access to a suitable Facebook Page and Instant Form.
- [ ] Configure a Meta developer app as required.
- [ ] Create a test lead using Meta's Lead Ads Testing Tool.
- [ ] Confirm required permissions/token type using current official docs.
- [ ] Store tokens securely (e.g. Apps Script Properties), never in GitHub.
- [ ] Retrieve the test lead into Meta with Record Type=LEAD and a test marker.
- [ ] Test stable-ID deduplication, token errors/expiry, and safe logging without exposing PII.
- [ ] Implement Meta Insights reporting separately if needed.
Phase 4: Google Ads
- [ ] Inspect GoogleAdsScript.js and verify GAQL fields against the selected resource/reporting level.
- [ ] Set the correct spreadsheet ID and Google tab.
- [ ] Test sample data first; use a suitable authorised account for live testing later.
- [ ] Handle single account vs manager account with multiple child accounts.
- [ ] Use stable keys that include account ID, campaign ID, date and reporting level.
- [ ] Verify currency, date range, conversions and click metric definitions.
- [ ] Ensure refreshes do not duplicate rows.
Phase 5: Looker Studio
- [ ] Use Master for the overview and platform tabs for detail.
- [ ] Add scorecards for spend, leads and CPL; time trends; platform comparison; campaign table; date/account filters.
- [ ] Avoid blending different row granularities in ways that double-count metrics.
- [ ] Ensure test leads do not inflate production totals.
- [ ] Create one report per client and keep client access separated.
Phase 6: Production readiness
- [ ] Identify the real customer lead source(s): CRM, website form, IVR/call system or another source.
- [ ] Define authoritative lead counts and attribution rules for Google/IVR and Meta.
- [ ] Define deduplication across ad platforms, CRM and IVR.
- [ ] Decide refresh frequency, backfill range, timezone and currency handling.
- [ ] Add monitoring, safe error reporting, retry/quota handling and refresh status.
- [ ] Apply least-privilege permissions, token protection, PII minimisation and retention rules.
- [ ] Test with one authorised client before onboarding more clients.
- [ ] Create reusable client/account configuration and onboarding steps.
Instructions for the next AI assistant
Continue step by step. First ask whether the user ran setupPrototype() and what they see. If they have not run it, walk them through it. If they report an error, troubleshoot that exact error. If it succeeds, validate the three tabs and sample calculations before moving to Meta test-lead setup.

The user asked for code previously, but has not confirmed execution. Inspect the actual repository files before claiming a function exists or integration is complete. If code is missing in the new chat, ask the user to upload the ZIP/repository or inspect the checked-in files. Do not invent the implementation's state.
User requirements/preferences
- Wants explicit, beginner-friendly, step-by-step instructions: exact file names, where to paste code, which function to run, and what result to expect.
- Wants a free prototype first, using sample data and Meta test leads; connect real accounts later.
- Wants one spreadsheet per client with exactly three tabs: Master, Google, Meta.
- Master contains key summary data; platform tabs contain detailed advertising data and individual leads.
- Wants research-backed guidance and official documentation where possible.
- Intends to commit this Markdown handoff and source files to GitHub so another AI/chat/account can continue.
