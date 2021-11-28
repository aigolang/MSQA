For EWS PowerShell script with OAuth authentication, we need an app registration in AAD portal.
1. Create an App registration from AAD portal.
2. Create a secret for the App registration, and copy the value into a temp notepad, notice that the secret value will be hidden after the page refreshed.

- For Application Type permission:
  1. Go to "API permissions" -> "Add a permission" -> "APIs my organization uses" -> search "Office 365 Exchange Online" -> "Application permissions" -> "full_access_as_app permission" -> Add permissions.
  2. Click "Grant admin consent for <Contoso>".
  3. Go to "Overview", copy the values of "Application (client) ID", "Directory (tenant) ID". Application secret is in the temp notepad at step 2, then update them in the PowerShell script.
  4. Notice that 2 DLL files are required, the names are in the script.
