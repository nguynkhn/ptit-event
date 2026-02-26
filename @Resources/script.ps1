param(
    $Action,
    $Port = 42069, $TimeoutSec = 5, $ExpireOffsetSec = 30,
    $SessionFile = "session.dat",
    [DateTime]$From, [DateTime]$To, $DateSpanDay = 7, $DateStart = "Now",
    $DateFormat = "dddd, dd MMMM", $TimeFormat = "HH:mm",
    [System.Globalization.CultureInfo]$Culture = "vi-VN"
)

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

function Exchange-Tokens {
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

    $response = Invoke-RestMethod -Method Post -Uri $TokenUrl -Body $body -TimeoutSec $TimeoutSec
    return @{
        AccessToken = $response.access_token
        RefreshToken = $response.refresh_token
        ExpireDate = (Get-Date).AddSeconds($response.expires_in).ToString("o")
    }
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

    $tokens = $json | ConvertFrom-Json
    $expireDate = [DateTimeOffset]::Parse($tokens.ExpireDate)
    if ($expireDate -lt (Get-Date).AddSeconds($ExpireOffsetSec)) {
        $tokens = Exchange-Tokens -RefreshToken $tokens.RefreshToken
        Store-Tokens -Tokens $tokens
    }

    return $tokens
}

function Store-Tokens {
    param ($Tokens)

    $json = $Tokens | ConvertTo-Json
    $secured = $json | ConvertTo-SecureString -AsPlainText -Force
    $encrypted = $secured | ConvertFrom-SecureString

    $encrypted | Set-Content $SessionFile
}

function Start-LoginFlow {
    $authRequest = "${AuthUrl}?response_type=code&client_id=${ClientId}&scope=${Scope}&redirect_uri=${RedirectUri}"
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

function Parse-Event {
    param($Source, $Data)

    $startDate = if ($Data.thoiGianGiao) { $Data.thoiGianGiao } else { $Data.thoiGianBatDau }
    $event = @{
        Title = $Source
        Type = ""
        Location = ""
        StartDate = [DateTimeOffset]::parse($startDate).ToLocalTime()
        EndDate = [DateTimeOffset]::parse($data.thoiGianKetThuc).ToLocalTime()
    }

     switch ($Source) {
        "QldtThoiKhoaBieu" {
            $event.Title = if ($Data.lopHocPhan.hocPhan.ten) {
                $Data.lopHocPhan.hocPhan.ten
            } elseif ($Data.lopHocPhan.maHocPhan) {
                $Data.lopHocPhan.maHocPhan
            } else {
                $Data.tenLopHocPhan
            }
            $event.Type = 1
            $event.Location = $Data.phongHoc
        }
        "QldtAssignment" {
            $event.Title = $Data.noiDung
            $event.Type = 3
            $event.location = $Data.tenLopHocPhan
        }
        "KhaoThiLichThi" {
            $event.Title = $Data.danhSachHocPhan.ten -join ", "
            $event.Type = 2
            $event.Location = $Data.phong.ma
        }
        "SlinkSuKien" {
            $event.Title = $Data.tenSuKien
            $event.Type = switch ($Data.loaiSuKien) {
                "Chung" { 0 }
                "Họp lớp" { 4 }
                "Cá nhân" { 5 }
                "Khác" { 6 }
            }
            $event.Location = $Data.diaDiem
        }
    }

    return $event
}

function Fetch-Events {
    param($AccessToken)

    $events = @()
    $headers = @{
        Authorization = "Bearer ${AccessToken}"
        Accept = "application/json"
    }
    foreach ($source in $EventSources.Keys) {
        $path = $EventSources[$source]
        $endpoint = "${ApiUrl}${path}/from/{0:o}/to/{1:o}" -f $From, $To

        try {
            $response = Invoke-RestMethod -Uri $endpoint -Headers $headers -TimeoutSec $TimeoutSec
            if (-not $response.success) {
                continue
            }

            foreach ($data in $response.data) {
                $event = Parse-Event -Source $source -Data $data
                $events += $event
            }
        } catch {
            continue
        }
    }

    $sortedEvents = $events | Sort-Object { $_.StartDate }, { $_.EndDate }
    return $sortedEvents | Group-Object { $_.StartDate.Date }
}

function Generate-Meters {
    param($EventGroups)

    $eventNo = 0
    $meters = @"
[Variables]
Generated=1
`n
"@

    for ($dateNo = 0; $dateNo -lt $EventGroups.Count; ++$dateNo) {
        $group = $EventGroups[$dateNo]

        $date = [DateTimeOffset]::parse($group.Name)
        $dateText = $date.ToString($DateFormat, $Culture).Normalize([System.Text.NormalizationForm]::FormC)
        $meters += @"
[MeterDate${dateNo}]
Meter=String
MeterStyle=StyleDate
Text=${dateText}
`n
"@

        foreach ($event in $group.Group) {
            $timeText = "{0:${TimeFormat}} - {1:${TimeFormat}}" -f $event.StartDate, $event.EndDate
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

    return $meters
}

[Console]::OutputEncoding = [System.Text.Encoding]::Unicode
switch ($Action.ToLower()) {
    "Status" { return (Load-Tokens) -ne $null }
    "Login" {
        try {
            $code = Start-LoginFlow
            $tokens = Exchange-Tokens -Code $code
            Store-Tokens -Tokens $tokens
            return $true
        } catch {
            return $false
        }
    }
    "Generate" {
        $tokens = Load-Tokens

        if (-not $From -and -not $To) {
            $From = Get-Date
            switch ($DateStart.ToLower()) {
                "Now" {}
                "Today" { $From = $From.Date }
                "WeekStart" {
                    $firstDayOfWeek = $Culture.DateTimeFormat.FirstDayOfWeek
                    $diff = ($From.DayOfWeek - $firstDayOfWeek + 7) % 7
                    $From = $From.Date.AddDays(-$diff)
                }
                default {
                    $dateStartOffset = $DateStart -as [int]
                    if ($dateStartOffset -eq $null) {
                        Write-Error "Invalid DateStart"
                    }

                    $From = $From.Date.AddDays(-$dateStartOffset)
                }
            }
            $To = $From.AddDays($DateSpanDay)
        } elseif (-not $From -or-not $To) {
            Write-Error "Missing From or To"
        }

        $eventGroups = Fetch-Events -AccessToken $tokens.AccessToken
        return Generate-Meters -EventGroups $eventGroups
    }
    default { Write-Error "Invalid Action" }
}
