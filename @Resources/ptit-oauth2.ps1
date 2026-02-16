param([switch]$Login, $SessionFile = "session.dat", $ExpireOffset = 30)

$ApiUrl = "https://gwdu.ptit.edu.vn"
$AuthUrl = "${ApiUrl}/sso/realms/ptit/protocol/openid-connect/auth"
$TokenUrl = "${ApiUrl}/sso/realms/ptit/protocol/openid-connect/token"
$ClientId = "ptit-connect"
$Scope = "email offline_access openid profile"
$RedirectUri = "http://localhost:42069/callback/"

function Exchange-Token {
    param ($Code, $RefreshToken)

    $body = if ($Code) {
        @{
            grant_type = "authorization_code"
            code = $Code
            client_id = $ClientId
            redirect_uri = $RedirectUri
        }
    } elseif ($RefreshToken) {
        @{
            grant_type = "refresh_token"
            scope = $Scope
            client_id = $ClientId
            refresh_token = $RefreshToken
        }
    }

    $response = Invoke-RestMethod -Method Post -Uri $TokenUrl -Body $body
    return @{
        accessToken = $response.access_token
        refreshToken = $response.refresh_token
        expireDate = (Get-Date).AddSeconds($response.expires_in).ToString("o")
    }
}

function Store-Tokens {
    param ($Tokens)

    $json = $Tokens | ConvertTo-Json
    $secured = $json | ConvertTo-SecureString -AsPlainText -Force
    $encrypted = $secured | ConvertFrom-SecureString

    $encrypted | Set-Content $SessionFile
}

function Load-Tokens {
    if (-not (Test-Path $SessionFile -PathType Leaf)) {
        return $null
    }

    $encrypted = Get-Content $SessionFile
    $secured = $encrypted | ConvertTo-SecureString
    $ptr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secured)
    try {
        $json = [Runtime.InteropServices.Marshal]::PtrToStringAuto($ptr)
    } finally {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr)       
    }

    return $json | ConvertFrom-Json
}

function Start-LoginFlow {
    $authRequest = "${authUrl}?response_type=code&client_id=${ClientId}&scope=${Scope}&redirect_uri=${RedirectUri}"
    Start-Process $authRequest

    $listener = [System.Net.HttpListener]::new()
    $listener.Prefixes.add($RedirectUri)
    $listener.Start()

    $context = $listener.GetContext()
    $code = $context.Request.QueryString["code"]

    $responseString = "Login successful. You can close this window."
    $buffer = [System.Text.Encoding]::UTF8.GetBytes($responseString)
    $context.Response.OutputStream.Write($buffer, 0, $buffer.Length)
    $context.Response.Close()
    $listener.Stop()

    return $code
}

function Check-LoginStatus {
    $tokens = Load-Tokens
    if (-not $tokens) {
        return $false
    }

    $expireDate = [DateTimeOffset]::Parse($tokens.expireDate)
    if ($expireDate -lt (Get-Date).AddSeconds($ExpireOffset)) {
        $tokens = Exchange-Token -RefreshToken $tokens.refreshToken
        Store-Tokens -Tokens $tokens
    }

    return $true
}

if ($Login) {
    try {
        $code = Start-LoginFlow
        $tokens = Exchange-Token -Code $code
        Store-Tokens -Tokens $tokens
        return $true
    } catch {
        return $false
    }
}

return Check-LoginStatus
