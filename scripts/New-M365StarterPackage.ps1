param(
  [Parameter()]
  [string]$AppId = $env:APP_ID,

  [Parameter()]
  [string]$AppBaseUrl = $env:APP_BASE_URL,

  [Parameter()]
  [string]$TargetUrl = $env:TARGET_URL,

  [Parameter()]
  [string]$AppName = $env:APP_NAME,

  [Parameter()]
  [string]$AppDescription = $env:APP_DESCRIPTION,

  [Parameter()]
  [string]$AppVersion = $env:APP_VERSION,

  [Parameter()]
  [string]$PackageNameValue = $env:PACKAGE_NAME,

  [Parameter()]
  [string]$EntityId = $env:ENTITY_ID,

  [Parameter()]
  [string]$DialogTitle = $env:DIALOG_TITLE,

  [Parameter()]
  [string]$DialogSize = $env:DIALOG_SIZE,

  [Parameter()]
  [string]$AutoOpenOnLoad = $env:AUTO_OPEN_ON_LOAD,

  [Parameter()]
  [string]$AccentColor = $env:ACCENT_COLOR,

  [Parameter()]
  [string]$DeveloperName = $env:DEVELOPER_NAME,

  [Parameter()]
  [string]$DeveloperWebsiteUrl = $env:DEVELOPER_WEBSITE_URL,

  [Parameter()]
  [string]$DeveloperPrivacyUrl = $env:DEVELOPER_PRIVACY_URL,

  [Parameter()]
  [string]$DeveloperTermsUrl = $env:DEVELOPER_TERMS_URL,

  [Parameter()]
  [string]$OutputDir,

  [Parameter()]
  [string]$PackageName = "m365-outlook-starter.zip"
)

$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($OutputDir)) {
  $OutputDir = Join-Path $PSScriptRoot "..\manifest\dist"
}

$manifestGenerator = Join-Path $PSScriptRoot "New-M365StarterManifest.ps1"
$assetsDir = Join-Path $PSScriptRoot "..\assets"
$stagingDir = Join-Path $OutputDir "package"
$packagePath = Join-Path $OutputDir $PackageName

if (Test-Path -LiteralPath $stagingDir) {
  Remove-Item -LiteralPath $stagingDir -Recurse -Force
}

New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
New-Item -ItemType Directory -Path $stagingDir -Force | Out-Null

$manifestParams = @{
  OutputPath = (Join-Path $stagingDir "manifest.json")
}

foreach ($entry in @{
  AppId = $AppId
  AppBaseUrl = $AppBaseUrl
  TargetUrl = $TargetUrl
  AppName = $AppName
  AppDescription = $AppDescription
  AppVersion = $AppVersion
  PackageName = $PackageNameValue
  EntityId = $EntityId
  DialogTitle = $DialogTitle
  DialogSize = $DialogSize
  AutoOpenOnLoad = $AutoOpenOnLoad
  AccentColor = $AccentColor
  DeveloperName = $DeveloperName
  DeveloperWebsiteUrl = $DeveloperWebsiteUrl
  DeveloperPrivacyUrl = $DeveloperPrivacyUrl
  DeveloperTermsUrl = $DeveloperTermsUrl
}.GetEnumerator()) {
  if (-not [string]::IsNullOrWhiteSpace($entry.Value)) {
    $manifestParams[$entry.Key] = $entry.Value
  }
}

& $manifestGenerator @manifestParams

foreach ($iconName in @("outline.png", "color.png", "color32x32.png")) {
  $source = Join-Path $assetsDir $iconName
  if (-not (Test-Path -LiteralPath $source)) {
    throw "Missing icon asset: $source"
  }

  Copy-Item -LiteralPath $source -Destination (Join-Path $stagingDir $iconName) -Force
}

if (Test-Path -LiteralPath $packagePath) {
  Remove-Item -LiteralPath $packagePath -Force
}

Compress-Archive -Path (Join-Path $stagingDir '*') -DestinationPath $packagePath -Force

Write-Host "Package written to $packagePath"
