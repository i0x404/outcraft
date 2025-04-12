# Outlook Email Craftsman

This script automates the creation of Outlook email accounts using the Microsoft Graph API.

## Prerequisites

1. PowerShell 5.1 or higher
2. Microsoft Azure AD tenant with appropriate permissions
3. Registered application in Azure AD with the following permissions:
   - Microsoft Graph API: `User.ReadWrite.All`

## Setup

1. Register an application in Azure AD:
   - Go to [Azure Portal](https://portal.azure.com)
   - Navigate to Azure Active Directory > App registrations > New registration
   - Name your application (e.g., "Outlook Account Creator")
   - Set the redirect URI to "http://localhost" (type: Web)
   - Click Register

2. Grant API permissions:
   - In your registered app, go to API permissions
   - Click "Add a permission"
   - Select Microsoft Graph > Application permissions
   - Search for and select "User.ReadWrite.All"
   - Click "Add permissions"
   - Click "Grant admin consent for [your tenant]"

3. Create a client secret:
   - In your registered app, go to Certificates & secrets
   - Click "New client secret"
   - Add a description and select an expiration period
   - Click "Add"
   - **Important**: Copy the secret value immediately as it won't be shown again

4. Configure the script:
   - Update the `config.json` file with your Azure AD application details:
     - ClientId: Your application (client) ID
     - ClientSecret: The client secret you created
     - TenantId: Your Azure AD tenant ID

## Usage

1. Ensure you have the required PowerShell modules:
   ```powershell
   Install-Module -Name MSAL.PS -Scope CurrentUser -Force
   ```

2. Run the script:
   ```powershell
   .\outlook_creator.ps1
   ```

3. Follow the prompts to enter user information:
   - First name
   - Last name
   - Date of birth (MM/DD/YYYY)
   - Place of birth

4. The script will:
   - Generate email address options based on the user's name
   - Check if the email addresses are available
   - Create a secure random password
   - Create the Outlook account using Microsoft Graph API
   - Display the account details

## Troubleshooting

If you encounter errors:

1. Verify your Azure AD application has the correct permissions
2. Check that your client secret is valid and not expired
3. Ensure your tenant ID is correct
4. Verify network connectivity to Microsoft Graph API endpoints

## Production Use

For production environments:

1. Store credentials securely (consider using Azure Key Vault)
2. Implement proper error handling and logging
3. Consider implementing rate limiting to avoid throttling
4. Set up monitoring for the script's execution

## License

This project is licensed under the MIT License - see the LICENSE file for details.
