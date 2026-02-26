param([switch]$Login, $Port = 42069, $Timeout = 5,
    $SessionFile = "session.dat", $EventFile = "events.inc",
    $ExpireOffset = 30, $DateSpan = 7,
    $Locale = "vi-VN", $DateFormat = "dddd, dd MMMM", $TimeFormat = "HH:mm")

$ApiUrl = "https://gwdu.ptit.edu.vn"
$AuthUrl = "${ApiUrl}/sso/realms/ptit/protocol/openid-connect/auth"
$TokenUrl = "${ApiUrl}/sso/realms/ptit/protocol/openid-connect/token"
$ClientId = "ptit-connect"
$Scope = "email offline_access openid profile"
$RedirectUri = "http://localhost:${Port}/callback/"
$EventSources = @{
    QldtThoiKhoaBieu = "/qldt/thoi-khoa-bieu/sv"
    QldtAssignment = "/qldt/assignment/lich/sinh-vien"
    KhaoThiLichThi = "/khao-thi/lich-thi/lich-thi/sv"
    SlinkSuKien = "/slink/su-kien/user"
}

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

    $response = Invoke-RestMethod -Method Post -Uri $TokenUrl -Body $body -TimeoutSec $Timeout
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

if ($Login) {
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

    $tokens = Exchange-Token -Code $code
    Store-Tokens -Tokens $tokens
} else {
    if (-not (Test-Path $SessionFile -PathType Leaf)) {
        return $false
    }

    $encrypted = Get-Content $SessionFile
    $secured = $encrypted | ConvertTo-SecureString
    $ptr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secured)
    try {
        $json = [Runtime.InteropServices.Marshal]::PtrToStringAuto($ptr)
    } finally {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr)
    }

    $tokens = $json | ConvertFrom-Json
    $expireDate = [DateTimeOffset]::Parse($tokens.expireDate)
    if ($expireDate -lt (Get-Date).AddSeconds($ExpireOffset)) {
        $tokens = Exchange-Token -RefreshToken $tokens.refreshToken
        Store-Tokens -Tokens $tokens
    }
}

$now = Get-Date
$from = $now.ToString("o")
$to = $now.AddDays($DateSpan).ToString("o")

$events = @()
$headers = @{
    Authorization = "Bearer $($tokens.accessToken)"
    Accept = "application/json"
}
foreach ($source in $EventSources.Keys) {
    $path = $EventSources[$source]
    $endpoint = "${ApiUrl}${path}/from/${from}/to/${to}"

    try {
        $response = Invoke-RestMethod -Uri $endpoint -Headers $headers -TimeoutSec $Timeout
        if (-not $response.success) {
            continue
        }

        foreach ($data in $response.data) {
            $startDate =
            $event = @{
                title = $Source
                type = ""
                location = ""
                startDate = if ($data.thoiGianGiao) { $data.thoiGianGiao } else { $data.thoiGianBatDau }
                endDate = $data.thoiGianKetThuc
            }

            switch ($source) {
                "QldtThoiKhoaBieu" {
                    $event.title = if ($data.lopHocPhan.hocPhan.ten) {
                        $data.lopHocPhan.hocPhan.ten
                    } elseif ($data.lopHocPhan.maHocPhan) {
                        $data.lopHocPhan.maHocPhan
                    } else {
                        $data.tenLopHocPhan
                    }
                    $event.type = 1
                    $event.location = $data.phongHoc
                }
                "QldtAssignment" {
                    $event.title = $data.noiDung
                    $event.type = 3
                    $event.location = $data.tenLopHocPhan
                }
                "KhaoThiLichThi" {
                    $event.title = $data.danhSachHocPhan.ten -join ", "
                    $event.type = 2
                    $event.location = $data.phong.ma
                }
                "SlinkSuKien" {
                    $event.title = $data.tenSuKien
                    $event.type = switch ($data.loaiSuKien) {
                        "Chung" { 0 }
                        "Họp lớp" { 4 }
                        "Cá nhân" { 5 }
                        "Khác" { 6 }
                    }
                    $event.location = $data.diaDiem
                }
            }

            $events += $event
        }
    } catch {
        continue
    }
}

if (-not $events) {
    $meters = @"
[Variables]
Generated=1

[MeterNoEvent]
Meter=String
MeterStyle=StyleNoEvent
"@
} else {

    $culture = [System.Globalization.CultureInfo]::GetCultureInfo($Locale)
    $dates = $events | Sort-Object { $_.startDate }, { $_.endDate } | Group-Object { $_.startDate }

    $meters = @"
[Variables]
Generated=1
MeterFirst=MeterDate0
MeterLast=MeterDate$($dates.Count - 1)
`n
"@

    $eventNo = 0
    for ($dateNo = 0; $dateNo -lt $dates.Count; ++$dateNo) {
        $date = $dates[$dateNo]
        $dateText = [DateTimeOffset]::parse($date.Name).ToLocalTime().ToString($DateFormat, $culture).Normalize([System.Text.NormalizationForm]::FormC)
        $meters += @"
[MeterDate${dateNo}]
Meter=String
MeterStyle=StyleDate
Text=${dateText}
`n
"@

    foreach ($event in $date.Group) {
        $timeText = [DateTimeOffset]::parse($event.startDate).ToLocalTime().ToString($TimeFormat) + " - " + [DateTimeOffset]::parse($event.endDate).ToLocalTime().ToString($TimeFormat)
        $meters += @"
[MeterCard${eventNo}]
Meter=Shape
MeterStyle=StyleCard

[MeterAccent${eventNo}]
Meter=Shape
MeterStyle=StyleAccent$($event.type)

[MeterTitle${eventNo}]
Meter=String
MeterStyle=StyleTitle
Text=$($event.title)

[MeterTime${eventNo}]
Meter=String
MeterStyle=StyleTime
Text=${timeText}

[MeterLocation${eventNo}]
Meter=String
MeterStyle=StyleLocation
Text=$($event.location)
`n
"@
            ++$eventNo
        }
    }
}

$meters | Out-File -FilePath $EventFile
