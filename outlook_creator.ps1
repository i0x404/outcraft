# Load required modules
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Load configuration
function Get-Configuration {
    $configPath = Join-Path -Path $PSScriptRoot -ChildPath "config.json"

    if (-not (Test-Path -Path $configPath)) {
        Write-Error "Configuration file not found at: $configPath"
        exit 1
    }

    try {
        $config = Get-Content -Path $configPath -Raw | ConvertFrom-Json
        return $config
    }
    catch {
        Write-Error "Failed to load configuration: $_"
        exit 1
    }
}

# Get authentication token for Microsoft Graph API
function Get-MicrosoftGraphToken {
    param (
        [PSCustomObject]$Config
    )

    try {
        # Check if MSAL.PS module is installed
        if (-not (Get-Module -ListAvailable -Name MSAL.PS)) {
            Write-Host "Installing MSAL.PS module..." -ForegroundColor Yellow
            Install-Module -Name MSAL.PS -Scope CurrentUser -Force
        }

        # Import the module
        Import-Module MSAL.PS

        # Get token using client credentials flow
        $msalParams = @{
            ClientId     = $Config.MicrosoftGraph.ClientId
            ClientSecret = (ConvertTo-SecureString $Config.MicrosoftGraph.ClientSecret -AsPlainText -Force)
            TenantId     = $Config.MicrosoftGraph.TenantId
            Scopes       = $Config.MicrosoftGraph.Scopes
        }

        $authResult = Get-MsalToken @msalParams
        return $authResult.AccessToken
    }
    catch {
        Write-Error "Failed to get Microsoft Graph token: $_"
        return $null
    }
}

# Check if email exists using Microsoft Graph API
function Test-EmailExists {
    param (
        [string]$EmailAddress,
        [string]$AccessToken
    )
    Write-Host "Checking if $EmailAddress already exists..." -ForegroundColor Yellow

    try {
        $headers = @{
            "Authorization" = "Bearer $AccessToken"
            "Content-Type"  = "application/json"
        }

        # Check if the email exists using Microsoft Graph API
        $uri = "https://graph.microsoft.com/v1.0/users?`$filter=mail eq '$EmailAddress' or userPrincipalName eq '$EmailAddress'"
        $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get

        # If any users are returned, the email exists
        return ($response.value.Count -gt 0)
    }
    catch {
        Write-Warning "Error checking if email exists: $_"
        # In case of error, assume the email might exist to be safe
        return $true
    }
}

# Create a new Outlook account using Microsoft Graph API
function New-OutlookAccount {
    param (
        [string]$FirstName,
        [string]$LastName,
        [string]$DateOfBirth,
        [string]$PlaceOfBirth,
        [string]$EmailAddress,
        [string]$Password,
        [string]$AccessToken
    )

    Write-Host "Creating new Outlook account with the following details:" -ForegroundColor Green
    Write-Host "  Email: $EmailAddress" -ForegroundColor Green
    Write-Host "  Name: $FirstName $LastName" -ForegroundColor Green
    Write-Host "  Date of Birth: $DateOfBirth" -ForegroundColor Green
    Write-Host "  Place of Birth: $PlaceOfBirth" -ForegroundColor Green

    try {
        $headers = @{
            "Authorization" = "Bearer $AccessToken"
            "Content-Type"  = "application/json"
        }

        # Prepare the user object
        $displayName = "$FirstName $LastName"
        $mailNickname = $EmailAddress.Split('@')[0]

        $userObject = @{
            accountEnabled = $true
            displayName = $displayName
            mailNickname = $mailNickname
            userPrincipalName = $EmailAddress
            passwordProfile = @{
                forceChangePasswordNextSignIn = $false
                password = $Password
            }
            givenName = $FirstName
            surname = $LastName
        } | ConvertTo-Json

        Write-Host "Creating account..." -ForegroundColor Yellow

        # Create the user using Microsoft Graph API
        $uri = "https://graph.microsoft.com/v1.0/users"
        $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method Post -Body $userObject

        Write-Host "Account created successfully!" -ForegroundColor Green

        return @{
            EmailAddress = $EmailAddress
            Password = $Password
            FirstName = $FirstName
            LastName = $LastName
            UserId = $response.id
        }
    }
    catch {
        Write-Error "Failed to create Outlook account: $_"
        throw "Failed to create Outlook account: $_"
    }
}

function New-SecurePassword {
    # Generate a secure random password
    $length = 12
    $nonAlphanumeric = 2
    $digits = 2
    $uppercase = 2

    $charSet = 'abcdefghijklmnopqrstuvwxyz'
    $random = New-Object System.Random

    # Start with required characters
    $password = ""

    # Add uppercase letters
    for ($i = 0; $i -lt $uppercase; $i++) {
        $password += $charSet[$random.Next(0, $charSet.Length)].ToString().ToUpper()
    }

    # Add digits
    for ($i = 0; $i -lt $digits; $i++) {
        $password += $random.Next(0, 10).ToString()
    }

    # Add special characters
    $specialChars = '!@#$%^&*()-_=+[]{}|;:,.<>?'
    for ($i = 0; $i -lt $nonAlphanumeric; $i++) {
        $password += $specialChars[$random.Next(0, $specialChars.Length)]
    }

    # Fill the rest with lowercase letters
    while ($password.Length -lt $length) {
        $password += $charSet[$random.Next(0, $charSet.Length)]
    }

    # Shuffle the password characters
    $passwordArray = $password.ToCharArray()
    $n = $passwordArray.Length
    while ($n -gt 1) {
        $n--
        $k = $random.Next(0, $n + 1)
        $temp = $passwordArray[$k]
        $passwordArray[$k] = $passwordArray[$n]
        $passwordArray[$n] = $temp
    }

    return -join $passwordArray
}

function Get-ValidInput {
    param (
        [string]$Prompt,
        [string]$ErrorMessage,
        [scriptblock]$ValidationScriptBlock
    )

    $isValid = $false
    $userInput = ""

    do {
        $userInput = Read-Host -Prompt $Prompt

        # Check if input is not empty or whitespace
        if (-not [string]::IsNullOrWhiteSpace($userInput)) {
            $isValid = $true
        } else {
            Write-Host $ErrorMessage -ForegroundColor Red
            $isValid = $false
        }

    } while (-not $isValid)

    return $userInput
}

# Main script execution starts here
try {
    Clear-Host
    Write-Host "===== Outlook Account Creator =====" -ForegroundColor Cyan
    Write-Host "This script will create a new Outlook email account." -ForegroundColor Cyan
    Write-Host "Please provide the required information below." -ForegroundColor Cyan
    Write-Host "===============================" -ForegroundColor Cyan

    # Load configuration
    $config = Get-Configuration
    Write-Host "Environment: $($config.Environment)" -ForegroundColor Cyan

    # Get Microsoft Graph API token
    $accessToken = Get-MicrosoftGraphToken -Config $config
    if (-not $accessToken) {
        Write-Error "Failed to get Microsoft Graph API token. Please check your configuration."
        exit 1
    }

    # Get first name
    $firstName = Get-ValidInput -Prompt "Enter first name" -ErrorMessage "First name cannot be empty" -ValidationScriptBlock { $true }

    # Get last name
    $lastName = Get-ValidInput -Prompt "Enter last name" -ErrorMessage "Last name cannot be empty" -ValidationScriptBlock { $true }

    # Get date of birth
    $dateOfBirth = Get-ValidInput -Prompt "Enter date of birth (MM/DD/YYYY)" -ErrorMessage "Please enter a valid date in MM/DD/YYYY format" -ValidationScriptBlock {
        param($input)
        try {
            $date = [DateTime]::ParseExact($input, "MM/dd/yyyy", $null)
            return $true
        }
        catch {
            Write-Host "Please enter a valid date in MM/DD/YYYY format" -ForegroundColor Red
            return $false
        }
    }

    # Get place of birth
    $placeOfBirth = Get-ValidInput -Prompt "Enter place of birth" -ErrorMessage "Place of birth cannot be empty" -ValidationScriptBlock { $true }

    # Generate email address options
    $domain = "outlook.com"
    $emailOptions = @(
        "$($firstName).$($lastName)@$domain",
        "$($firstName.Substring(0,1))$($lastName)@$domain",
        "$($firstName)$($lastName.Substring(0,1))@$domain"
    )

    # Find an available email address
    $emailAddress = $null
    foreach ($email in $emailOptions) {
        $emailToCheck = $email.ToLower() -replace '[^a-z0-9.@]', ''

        Write-Host "Checking availability of: $emailToCheck" -ForegroundColor Yellow

        if (-not (Test-EmailExists -EmailAddress $emailToCheck -AccessToken $accessToken)) {
            $emailAddress = $emailToCheck
            break
        }
    }

    # If no email address is available from the predefined options, create one with a number
    if (-not $emailAddress) {
        $attemptCount = 1
        do {
            $emailToCheck = "$($firstName).$($lastName)$attemptCount@$domain".ToLower() -replace '[^a-z0-9.@]', ''

            Write-Host "Checking availability of: $emailToCheck" -ForegroundColor Yellow

            if (-not (Test-EmailExists -EmailAddress $emailToCheck -AccessToken $accessToken)) {
                $emailAddress = $emailToCheck
                break
            }

            $attemptCount++
        } while ($attemptCount -le 10)
    }

    if (-not $emailAddress) {
        Write-Host "Failed to find an available email address after multiple attempts." -ForegroundColor Red
        Write-Host "Please try again with different name information." -ForegroundColor Red
        exit 1
    }

    # Generate a secure password
    $password = New-SecurePassword

    # Create the account
    $account = New-OutlookAccount -FirstName $firstName -LastName $lastName -DateOfBirth $dateOfBirth -PlaceOfBirth $placeOfBirth -EmailAddress $emailAddress -Password $password -AccessToken $accessToken

    # Display the results
    Write-Host "`n===== Account Creation Successful =====" -ForegroundColor Green
    Write-Host "Email Address: $($account.EmailAddress)" -ForegroundColor Cyan
    Write-Host "Password: $($account.Password)" -ForegroundColor Cyan
    Write-Host "Name: $($account.FirstName) $($account.LastName)" -ForegroundColor Cyan
    Write-Host "User ID: $($account.UserId)" -ForegroundColor Cyan
    Write-Host "===============================" -ForegroundColor Green

    Write-Host "`nPlease store these credentials securely." -ForegroundColor Yellow
    Write-Host "You can now use these credentials to log into Outlook." -ForegroundColor Yellow
}
catch {
    Write-Error "An error occurred: $_"

    # Get detailed error information
    if ($_.Exception.Response) {
        $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
        $reader.BaseStream.Position = 0
        $reader.DiscardBufferedData()
        $responseBody = $reader.ReadToEnd()
        Write-Error "Response body: $responseBody"
    }

    exit 1
}