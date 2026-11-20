# Skrypt budujacy pakiet dodatku NVDA (.nvda-addon) dla outlookSpeechFilter
# Po zbudowaniu pakietu oblicza jego SHA-256 i zapisuje wersje, adres pobierania oraz sume kontrolna
# do update_info.json i updates.json (aktualizator dodatku odrzuca pakiety bez zgodnej sumy SHA-256).
param(
    [Parameter(Mandatory=$false)]
    [string]$DownloadUrl = "https://github.com/strugo10/outlookSpeechFilter/releases/latest/download/outlookSpeechFilter.nvda-addon"
)
$ErrorActionPreference = "Stop"

$scriptDir = $PSScriptRoot
$sourceDir = $scriptDir
$addonName = "outlookSpeechFilter"
$addonFile = Join-Path $scriptDir "$addonName.nvda-addon"
$tempZip = Join-Path $scriptDir "temp_build_$addonName.zip"
$validatorScript = Join-Path $scriptDir "Test-AddonSyntax.ps1"
if (-not (Test-Path $validatorScript)) {
    $validatorScript = Join-Path (Split-Path -Parent $scriptDir) "Test-AddonSyntax.ps1"
}

# 1. KROK KONTROLNY: Obowiazkowa walidacja skladni Python i AST przed spakowaniem
if (-not (Test-Path $validatorScript)) {
    Write-Host "BLAD KRYTYCZNY: Brak skryptu walidacji Test-AddonSyntax.ps1! Pakowanie zostalo zablokowane." -ForegroundColor Red
    exit 1
}
Write-Host "Uruchamianie automatycznej walidacji skladni Python i AST..." -ForegroundColor Cyan
& powershell.exe -ExecutionPolicy Bypass -File $validatorScript -TargetDir $sourceDir
if ($LASTEXITCODE -ne 0) {
    Write-Host "BLAD: Wykryto usterke skladniowa lub AST! Pakowanie zostalo przerwane." -ForegroundColor Red
    exit 1
}

# 2. Wersja i minimalna wersja NVDA z manifest.ini
$manifestPath = Join-Path $sourceDir "manifest.ini"
$manifestText = [System.IO.File]::ReadAllText($manifestPath, [System.Text.Encoding]::UTF8)
$verMatch = [regex]::Match($manifestText, '(?m)^\s*version\s*=\s*"?([^"\r\n]+)"?\s*$')
if (-not $verMatch.Success) {
    Write-Host "BLAD: Nie znaleziono wersji w manifest.ini." -ForegroundColor Red
    exit 1
}
$version = $verMatch.Groups[1].Value.Trim()
$minMatch = [regex]::Match($manifestText, '(?m)^\s*minimumNVDAVersion\s*=\s*"?([^"\r\n]+)"?\s*$')
$minNvda = if ($minMatch.Success) { $minMatch.Groups[1].Value.Trim() } else { "2024.1.0" }

if (Test-Path $tempZip) { Remove-Item $tempZip -Force }
if (Test-Path $addonFile) { Remove-Item $addonFile -Force }

# Usun foldery __pycache__ przed spakowaniem
Get-ChildItem -Path $sourceDir -Recurse -Directory -Filter "__pycache__" | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue

# 3. Pakowanie wylacznie plikow dodatku (lista dozwolonych elementow).
# Do pakietu NIE trafiaja: update_info.json, updates.json, .gitignore, README.md, .git, skrypty budowania i walidacji.
$includeItems = @("manifest.ini", "installTasks.py", "globalPlugins", "locale", "sounds", "doc.html", "doc_pl.html", "doc_en.html")
$filesToPack = @()
foreach ($item in $includeItems) {
    $p = Join-Path $sourceDir $item
    if (Test-Path $p) { $filesToPack += (Get-Item $p).FullName }
}
Write-Host "Pakowanie dodatku w wersji $version z $sourceDir..."
Compress-Archive -Path $filesToPack -DestinationPath $tempZip -CompressionLevel Optimal
Rename-Item -Path $tempZip -NewName "$addonName.nvda-addon"

# 4. Kontrola zawartosci pakietu
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::OpenRead($addonFile)
try {
    $forbidden = $zip.Entries | Where-Object { $_.FullName -match '(^|[\\/])(update_info\.json|updates\.json|\.gitignore|README\.md|build\.(ps1|cmd)|Test-AddonSyntax\.ps1)$' -or $_.FullName -match '(^|[\\/])(\.git|__pycache__)([\\/]|$)' }
    if ($forbidden) {
        Write-Host "BLAD: Pakiet zawiera niedozwolone pliki: $($forbidden.FullName -join ', ')" -ForegroundColor Red
        exit 1
    }
} finally {
    $zip.Dispose()
}

# 5. Suma kontrolna SHA-256 i aktualizacja plikow informacji o aktualizacji
$sha256 = (Get-FileHash -Path $addonFile -Algorithm SHA256).Hash.ToLowerInvariant()
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

function Set-JsonProperty($obj, [string]$name, $value) {
    if ($obj.PSObject.Properties[$name]) {
        $obj.$name = $value
    } else {
        $obj | Add-Member -NotePropertyName $name -NotePropertyValue $value
    }
}

function Update-UpdateEntry($entry) {
    Set-JsonProperty $entry "version" $version
    Set-JsonProperty $entry "downloadUrl" $DownloadUrl
    Set-JsonProperty $entry "sha256" $sha256
    Set-JsonProperty $entry "minNVDAVersion" $minNvda
}

# update_info.json (schemat plaski: name, version, downloadUrl, sha256, changelog, minNVDAVersion)
$updateInfoPath = Join-Path $sourceDir "update_info.json"
if (Test-Path $updateInfoPath) {
    $info = [System.IO.File]::ReadAllText($updateInfoPath, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
    Update-UpdateEntry $info
    [System.IO.File]::WriteAllText($updateInfoPath, ($info | ConvertTo-Json -Depth 5), $utf8NoBom)
    Write-Host "Zaktualizowano update_info.json (wersja $version, sha256 $sha256)" -ForegroundColor Green
}

# updates.json (schemat z kluczem nazwy dodatku: { "outlookSpeechFilter": { ... } })
$updatesPath = Join-Path $sourceDir "updates.json"
if (Test-Path $updatesPath) {
    $all = [System.IO.File]::ReadAllText($updatesPath, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
    if (-not $all.PSObject.Properties[$addonName]) {
        $all | Add-Member -NotePropertyName $addonName -NotePropertyValue ([pscustomobject]@{})
    }
    Update-UpdateEntry $all.$addonName
    if ($info -and $info.PSObject.Properties["changelog"]) {
        Set-JsonProperty $all.$addonName "changelog" $info.changelog
    }
    [System.IO.File]::WriteAllText($updatesPath, ($all | ConvertTo-Json -Depth 5), $utf8NoBom)
    Write-Host "Zaktualizowano updates.json (wpis $addonName)" -ForegroundColor Green
}

Write-Host "Gotowe! Utworzono pakiet: $addonFile" -ForegroundColor Green
Write-Host "SHA-256: $sha256"
Write-Host "UWAGA: do wydania na GitHubie nalezy wgrac dokladnie ten plik - kazde ponowne budowanie zmienia sume SHA-256." -ForegroundColor Yellow
