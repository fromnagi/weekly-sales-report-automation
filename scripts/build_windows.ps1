$ErrorActionPreference = 'Stop'

$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$pyinstaller = Join-Path $repoRoot '.venv\Scripts\pyinstaller.exe'
if (-not (Test-Path -LiteralPath $pyinstaller -PathType Leaf)) {
    throw 'Install the repository build extra in .venv first: .\.venv\Scripts\python.exe -m pip install -e ".[build]"'
}
$version = (& $pyinstaller --version).Trim()
if ($LASTEXITCODE -ne 0 -or $version -ne '6.22.3') {
    throw "Expected PyInstaller 6.22.3 in .venv; found $version"
}
$helpText = (& $pyinstaller --help) -join "`n"
foreach ($option in @('--onedir', '--windowed', '--contents-directory')) {
    if (-not $helpText.Contains($option)) { throw "Installed PyInstaller does not support $option" }
}

$distRoot = Join-Path $repoRoot 'dist'
$workRoot = Join-Path $repoRoot 'build'
$package = Join-Path $distRoot 'Weekly Sales Report'
$entry = Join-Path $repoRoot 'src\weekly_sales_report\windows_app.py'

& $pyinstaller --onedir --windowed --contents-directory _internal `
    --name 'Weekly Sales Report' --paths (Join-Path $repoRoot 'src') `
    --distpath $distRoot --workpath $workRoot --specpath $workRoot `
    --noconfirm $entry
if ($LASTEXITCODE -ne 0) { throw "PyInstaller failed with exit code $LASTEXITCODE" }

foreach ($directory in @('input', 'output', 'examples')) {
    New-Item -ItemType Directory -Path (Join-Path $package $directory) -Force | Out-Null
}
Copy-Item -LiteralPath (Join-Path $repoRoot 'START_HERE.txt') -Destination $package
Copy-Item -LiteralPath (Join-Path $repoRoot 'INPUT_REQUIREMENTS.txt') -Destination $package
'Put your sales CSV or XLSX files in this folder, then double-click Weekly Sales Report.exe.' |
    Set-Content -LiteralPath (Join-Path $package 'input\PUT_FILES_HERE.txt') -Encoding UTF8
foreach ($name in @('01_north.csv', '02_south.xlsx', '03_legacy.csv')) {
    Copy-Item -LiteralPath (Join-Path $repoRoot "sample_data\input\$name") `
        -Destination (Join-Path $package "examples\$name")
}
$archive = Join-Path $distRoot 'WeeklySalesReport.zip'
Compress-Archive -LiteralPath $package -DestinationPath $archive -CompressionLevel Optimal -Force
Write-Host "Portable package: $package"
Write-Host "Delivery archive: $archive"
