<#

.SYNOPSIS
    Office 365 Management Activity API demonstration

.DESCRIPTION
    A minimal proof-of-concept to demonstrate how to connect to the Office 365 Management Activity API with PowerShell. Requires tenant id, client id (a.k.a. app id) and client secret for a registered application with application permissions to the API.

.LINK
    See https://docs.microsoft.com/en-us/office/office-365-management-api/ for more information about the Office 365 Management Activity API

#>


Param(
    [Parameter(Mandatory = $true)]
    [string]$tenant_id,
    [Parameter(Mandatory = $true)]
    [string]$client_id,
    [Parameter(Mandatory = $true)]
    [string]$client_secret,
    [Parameter(Mandatory = $false)]
    [switch]$debug_mode
)

[string]$manage_api_url = 'https://manage.office.com'
[string]$protocol = 'TLS12'
[string]$regex_guid = '^[0-9a-z]{8}-[0-9a-z]{4}-[0-9a-z]{4}-[0-9a-z]{4}-[0-9a-z]{12}$'
[string]$regex_secret = '^[0-9a-zA-Z~_.-]{34,}$'
[string]$grant_type = 'client_credentials'


Function Set-Protocol {
    $Error.Clear()
    Try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::$protocol
    }
    Catch {
        [array]$error_clone = $Error.Clone()
        [string]$error_message = $error_clone | Where-Object { $null -ne $_.Exception } | Select-Object -First 1 | Select-Object -ExpandProperty Exception | Select-Object -ExpandProperty Message
        Write-Host "Error: Failed to set the protocol to [$protocol] due to [$error_message]"
        Exit
    }    
}

Function Confirm-Script-Parameters {
    Set-Protocol
    If ( $tenant_id -notmatch $regex_guid) {
        Write-Host "[$tenant_id] does not appear to be a valid tenant id" -ForegroundColor Red
        Exit
    }
    If ( $client_id -notmatch $regex_guid) {
        Write-Host "[$client_id] does not appear to be a valid client id" -ForegroundColor Red
        Exit
    }
    If ( $client_secret -notmatch $regex_secret) {
        Write-Host "[$client_secret] does not appear to be a valid client secret" -ForegroundColor Red
        Exit
    }
}

Function New-Request-Headers-Parameters {
    [OutputType([hashtable])]
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [ValidateLength(36,36)]
        [string]$tenant_id,
        [Parameter(Mandatory = $true)]
        [ValidateLength(36,36)]
        [string]$client_id,
        [Parameter(Mandatory = $true)]
        [string]$client_secret
    )
    [string]$uri = "https://login.microsoftonline.com/$tenant_id/oauth2/token?api-version=1.0"
    [hashtable]$body = @{
        'client_id'     = $client_id;
        'resource'      = $manage_api_url;
        'client_secret' = $client_secret;
        'grant_type'    = $grant_type;
    }
    [hashtable]$parameters = @{
        'Uri'             = $uri;
        'Method'          = 'POST';
        'Body'            = $body;
        'ContentType'     = 'application/x-www-form-urlencoded';
        'UseBasicParsing' = $true;
    }
    Return $parameters
}

Function New-Parameters {
    [OutputType([hashtable])]
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [string]$command
    )    
    [string]$base_url = $manage_api_url + '/api/v1.0/' + $tenant_id + '/activity/feed/'
    [string]$final_url = ($base_url + $command)
    [hashtable]$parameters = @{
        'Uri'             = $final_url;
        'Method'          = 'GET';        
        'ContentType'     = 'application/json';
        'UseBasicParsing' = $true;
    }
    Return $parameters
}

Function Invoke-API {
    [OutputType([PSCustomObject])]
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [hashtable]$parameters,
        [Parameter(Mandatory = $false)]
        [hashtable]$headers
    )
    If ( $null -ne $headers) {
        $parameters.Add('Headers', $headers)
    }
    If( $debug_mode -eq $true)
    {
        [string]$parameters_display = ConvertTo-Json -InputObject $parameters -Compress
        Write-Host "Sending parameters to Office 365 Management API " -NoNewLine
        Write-host $parameters_display -ForegroundColor Yellow
    }
    $Error.Clear()
    Try {
        $ProgressPreference = 'SilentlyContinue'
        [PSCustomObject]$response = Invoke-WebRequest @parameters
        $ProgressPreference = 'Stop'
    }
    Catch {
        [array]$error_clone = $Error.Clone()
        [string]$error_message = $error_clone | Where-Object { $null -ne $_.Exception } | Select-Object -First 1 | Select-Object -ExpandProperty Exception | Select-Object -ExpandProperty Message
        Write-Host "Error: Invoke-WebRequest failed due to [$error_message]" -ForegroundColor Red
        Exit
    }
    If ($response.StatusCode -isnot [int]) {
        Write-Host 'Somehow there was no status code received?' -ForegroundColor Red
        Exit
    }
    [int]$status_code = $response.StatusCode
    If ( $status_code -ne 200) {
        Write-Host "Error: Received status code [$status_code] instead of 200. Please look into this." -ForegroundColor Red
        Exit
    }
    [string]$response_content = $response.Content
    [int]$response_content_length = $response_content.Length
    If ($response_content_length -eq 0) {
        Write-Host 'Error: Somehow the response content was empty!' -ForegroundColor Red
        Exit
    }
    [PSCustomObject]$response_content_object = $response_content | ConvertFrom-Json
    Return $response_content_object
}

Function Connect-To-API {
    [OutputType([hashtable])]
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [hashtable]$parameters
    )
    If( $debug_mode -eq $true)
    {
        [string]$parameters_display = ConvertTo-Json -InputObject $parameters -Compress
        Write-Host "Connecting to Office 365 Management API " -NoNewLine
        Write-Host $parameters_display -ForegroundColor Yellow
    }
    [PSCustomObject]$response_content_object = Invoke-API -parameters $parameters	
    [string]$token_type = $response_content_object.token_type
    If ($token_type -cne 'Bearer') {
        Write-Host -entry_type Error -log_message "Somehow the token type received is not exactly 'Bearer' (case-sensitive)" -ForegroundColor Red
        Exit
    }
    [string]$access_token = $response_content_object.access_token
    [int]$access_token_length = $access_token.Length
    If ($access_token_length -eq 0) {
        Write-Host 'Somehow the access token is 0 characters in length' -ForegroundColor Red
        Exit
    }
    [string]$authorization_string = "$token_type $access_token"
    [hashtable]$headers = @{ Authorization = $authorization_string; }
    Return $headers
}

Function New-Headers {
    [OutputType([hashtable])]
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [ValidateLength(36,36)]
        [string]$tenant_id,
        [Parameter(Mandatory = $true)]
        [ValidateLength(36,36)]
        [string]$client_id,
        [Parameter(Mandatory = $true)]
        [string]$client_secret
    )
    Confirm-Script-Parameters
    [hashtable]$parameters = New-Request-Headers-Parameters -tenant_id $tenant_id -client_id $client_id -client_secret $client_secret
    [hashtable]$headers = Connect-To-API -parameters $parameters
    Return $headers
}

Function Get-Subscriptions {
    [OutputType([PSCustomObject])]
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $false)]
        [hashtable]$headers
    )
    [hashtable]$parameters = New-Parameters -command 'subscriptions/list'
    [PSCustomObject]$subscriptions = Invoke-API -parameters $parameters -headers $headers
    Return $subscriptions
}

[hashtable]$headers = New-Headers -tenant_id $tenant_id -client_id $client_id -client_secret $client_secret
[array]$subscriptions = Get-Subscriptions -headers $headers
Return $subscriptions

