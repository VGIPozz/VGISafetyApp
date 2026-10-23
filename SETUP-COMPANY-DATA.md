# Company data setup (Excel workbook in SharePoint)

Once this is set up, the home page:

- **Looks up the Site Main ID.** Type it in and the app searches **Job Information 2027**, then **2026**, then **2025**. It fills in the **Site Secondary ID**, **VGI job number**, **Site address** and **Site city**, and every form receives them. If the ID appears more than once in that year, the crew lead picks the right job.
- **Shows dropdowns** for the **crew lead and workers** (from the Staff List, with their title filled in) and the **customer** (from the Customer List). Each dropdown has an "Other" choice for anyone or anything not on the list, such as subcontractors and visitors.

How it works: inside Teams, the app signs the person in through Teams (Microsoft "nested app authentication"). It then downloads the workbook from SharePoint with **read-only** access and keeps a copy on the phone, so lookups still work with no signal. There's no server and no password or client secret, and none of the workbook's data is stored on GitHub.

Until the settings below are filled in, everything is typed by hand as before.

---

## 1. The workbook (you)

Use **Safety_App_Data.xlsx** in a SharePoint document library, for example the Safety site's **Documents** library. It can sit in a folder.

| Sheet | Columns (row 1) | Notes |
|---|---|---|
| Job Information 2027, Job Information 2026, Job Information 2025 | Site Main ID · Site Secondary ID · VGi Job number · Site Address · Site City | Searched newest year first. Add **Job Information 2028** next year and it's picked up automatically. |
| Staff List | Employee · Title | Blank rows are skipped. |
| Customer List | Customer List (heading), then one customer per row | Blank rows are skipped. |

Cells holding `0`, a single space, or `_Drop Down Selection` are treated as blank. Formulas are fine; the app reads the last calculated values.

Everyone who uses the app (at least crew leads) needs **read** access to the workbook. Edit access can stay with the office.

## 2. Register the app in Microsoft Entra (IT admin)

In the [Microsoft Entra admin center](https://entra.microsoft.com):

1. Open **App registrations** (under **Entra ID**, or **Identity → Applications** in older menus), then click **New registration**.
   - **Name:** `VGI Safety App`
   - **Supported account types:** *Accounts in this organizational directory only*
   - **Redirect URI:** platform **Single-page application (SPA)**, value `brk-multihub://vgipozz.github.io` (the domain only, with no path or trailing slash)
2. Copy the **Application (client) ID** and **Directory (tenant) ID** from **Overview**.
3. Go to **API permissions → Add a permission → Microsoft Graph → Delegated**, add **Sites.Read.All**, then click **Grant admin consent**.
   - *Tighter option:* use the delegated **Sites.Selected** permission, and grant the app read access to just the one site. Then set `scope: 'Sites.Selected'` below.
4. Don't create a client secret.

## 3. Fill in `settings.js`

```js
window.SAFETY_SETTINGS = {
  microsoft: { clientId: '…', tenantId: '…', scope: 'Sites.Read.All' },
  data: {
    siteUrl: 'https://violingroup.sharepoint.com/sites/Safety',   // the site's address
    library: 'Documents',
    filePath: 'Safety_App_Data.xlsx',                            // or 'Folder/Safety_App_Data.xlsx'
    jobSheets: 'Job Information', staffSheet: 'Staff List', customerSheet: 'Customer List'
  }
};
```

None of these values are secret. A person still has to be signed in to Teams with a company account that can open the workbook.

## 4. Check it

Open the **Site Safety** tab in Teams. At the top of **Job details**, you should see:

> ● 200 jobs, 18 staff, 7 customers (updated 9:41 a.m.)

| Message | What to do |
|---|---|
| *Sign in once to load jobs, staff and customers.* [Sign in] | Tap **Sign in** (only needed until admin consent is granted). |
| *…the company data file wasn't found…* | Check `siteUrl`, `library` and `filePath`. |
| *…you don't have access to the company data file…* | Give that person read access to the workbook. |
| *No job found for "X" in 2027, 2026, 2025.* | The ID isn't in any job sheet, so type the site details in by hand. Check for a typo in the sheet. |
| *Job lookup and staff lists load when the app is opened in Teams.* | Expected in a normal web browser. |
| *…saved copy from…* (amber dot) | Offline, or opened outside Teams. It refreshes the next time it opens in Teams with signal. |

Updating the data: edit the workbook in SharePoint. Phones pick up the changes the next time the app opens in Teams.
