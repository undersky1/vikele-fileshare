# SharePoint Storage Provider

Vikele FileShare can store uploaded files in a SharePoint Online document library using the Microsoft Graph API.

## Prerequisites

- A Microsoft 365 tenant with SharePoint Online
- An Azure AD (Entra ID) App Registration with **application** permissions
- The SharePoint site and document library must already exist

## Step 1: Register an Azure AD App

1. Go to [Azure Portal → App registrations](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade)
2. Click **New registration**
3. Name it (e.g. "Vikele FileShare")
4. Set **Supported account types** to "Accounts in this organizational directory only"
5. Click **Register**
6. Note the **Application (client) ID** and **Directory (tenant) ID**

## Step 2: Create a Client Secret

1. In your app registration, go to **Certificates & secrets**
2. Click **New client secret**
3. Set an expiry and click **Add**
4. Copy the **Value** immediately (you won't see it again)

## Step 3: Grant API Permissions

1. Go to **API permissions → Add a permission → Microsoft Graph → Application permissions**
2. Add `Sites.ReadWrite.All`
3. Click **Grant admin consent** for your organisation

## Step 4: Find your Site ID and Drive ID

### Site ID

Use Graph Explorer:

```
GET https://graph.microsoft.com/v1.0/sites/{hostname}:/{relative-path-to-site}
```

Example:
```
GET https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/sites/fileshare
```

The `id` field in the response is your Site ID.

### Drive ID

```
GET https://graph.microsoft.com/v1.0/sites/{siteId}/drives
```

Pick the `id` of the document library you want to use.

## Step 5: Configure in Vikele FileShare

1. Log in as admin
2. Go to **Admin → Configuration → SharePoint**
3. Fill in Tenant ID, Client ID, Client Secret, Site ID, Drive ID
4. Optionally set a Folder Path (e.g. `vikele-uploads`)
5. Toggle **Enabled** to on
6. Click **Save**

All new file shares will now store in your SharePoint document library.

## Limitations

- Maximum single-file upload size is ~250 MB (simple upload). For larger files a Graph upload session would be needed.
- SharePoint API rate limits apply (~10,000 calls per 10 minutes per tenant).
- The Azure AD client secret expires and must be rotated before expiry.
