# Kim qaysi versiyada ekanini ko'rsatadi (O'zbekiston vaqti bo'yicha).
# Ishlatish:  .\holat.ps1
$keyFile = Join-Path $PSScriptRoot ".telemetry_key.txt"
if (-not (Test-Path $keyFile)) {
    Write-Host "Kalit fayl topilmadi: $keyFile" -ForegroundColor Red
    exit 1
}
$key = (Get-Content $keyFile).Trim()
$url = "https://soddahisobot-telemetry.tasks-bot.workers.dev/status?key=$key"

try {
    $clients = (Invoke-RestMethod $url).clients
} catch {
    Write-Host "Serverga ulanib bo'lmadi: $_" -ForegroundColor Red
    exit 1
}

$nowUtc = [datetime]::UtcNow
$style = [System.Globalization.DateTimeStyles]::AdjustToUniversal -bor `
         [System.Globalization.DateTimeStyles]::AssumeUniversal

$clients | ForEach-Object {
    $utc = [datetime]::Parse($_.last_seen, $null, $style)
    $mins = ($nowUtc - $utc).TotalMinutes
    $qachon = if ($mins -lt 60) { "{0:N0} daqiqa oldin" -f $mins }
              elseif ($mins -lt 1440) { "{0:N1} soat oldin" -f ($mins / 60) }
              else { "{0:N1} kun oldin" -f ($mins / 1440) }
    [pscustomobject]@{
        Kompyuter = $_.host
        Versiya   = $_.version
        Ochilgan  = $utc.AddHours(5).ToString("dd.MM HH:mm")
        Qachon    = $qachon
        Marta     = $_.ping_count
    }
} | Format-Table -AutoSize

Write-Host ("Hozir (UZT): " + $nowUtc.AddHours(5).ToString("dd.MM.yyyy HH:mm")) -ForegroundColor DarkGray
