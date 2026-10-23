/* =====================================================================
   Company settings. Fill these in once IT has registered the app in
   Microsoft Entra (see SETUP-COMPANY-DATA.md). Leave them blank and the
   app works as before: everything is typed in by hand.

   None of these values are secret. They only say which Microsoft 365
   organization and which file to use; every person still has to be
   signed in to Teams with a company account to read the file.
   ===================================================================== */
window.SAFETY_SETTINGS = {
  microsoft: {
    clientId: '',   // Application (client) ID from the Entra app registration
    tenantId: '',   // Directory (tenant) ID from the same page
    scope: 'Sites.Read.All'   // Graph permission granted by IT (read-only). Use 'Sites.Selected' if IT limits access to one site.
  },
  data: {
    // The SharePoint site that holds the workbook, e.g. 'https://violingroup.sharepoint.com/sites/Safety'
    siteUrl: '',
    // Document library the file is in ('Documents' is the default library, shown as "Shared Documents" in links)
    library: 'Documents',
    // Path to the workbook inside that library, e.g. 'Safety_App_Data.xlsx' or 'Safety App/Safety_App_Data.xlsx'
    filePath: '',
    // Sheet names in the workbook (only change these if the sheets are renamed)
    jobSheets: 'Job Information',   // every sheet starting with this, e.g. "Job Information 2027", searched newest year first
    staffSheet: 'Staff List',
    customerSheet: 'Customer List'
  }
};
