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
  [string]$PackageName = $env:PACKAGE_NAME,

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
  [string]$OutputPath,

  [Parameter()]
  [string]$TemplatePath
)

$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
  $OutputPath = Join-Path $PSScriptRoot "..\manifest\generated\manifest.json"
}

if ([string]::IsNullOrWhiteSpace($TemplatePath)) {
  $TemplatePath = Join-Path $PSScriptRoot "..\manifest\app.manifest.template.json"
}

function Get-GitHubRepositoryInfo {
  $remoteUrl = & git remote get-url origin 2>$null
  if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($remoteUrl)) {
    return $null
  }

  $branchName = & git branch --show-current 2>$null
  if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($branchName)) {
    $branchName = "main"
  }

  $match = [regex]::Match(
    $remoteUrl.Trim(),
    'github\.com[:/](?<owner>[^/]+)/(?<repo>[^/.]+?)(?:\.git)?$'
  )
  if (-not $match.Success) {
    return $null
  }

  $owner = $match.Groups["owner"].Value
  $repo = $match.Groups["repo"].Value

  return @{
    Owner = $owner
    Repository = $repo
    Branch = $branchName.Trim()
    RepositoryUrl = "https://github.com/$owner/$repo"
    ContentBaseUrl = "https://cdn.jsdelivr.net/gh/$owner/$repo@$($branchName.Trim())/web"
  }
}

function Get-DefaultAppName {
  $chars = foreach ($code in @(
    0x30B5, 0x30A4, 0x30C9, 0x30D0, 0x30FC,
    0x30E9, 0x30F3, 0x30C1, 0x30E3, 0x30FC
  )) {
    [char]$code
  }

  return -join $chars
}

function Get-RequiredValue {
  param(
    [string]$Value,
    [string]$Name
  )

  if ([string]::IsNullOrWhiteSpace($Value)) {
    throw "Missing required value: $Name"
  }

  return $Value.Trim()
}

function Get-GuidValue {
  param(
    [string]$Value,
    [string]$Name
  )

  $text = Get-RequiredValue -Value $Value -Name $Name
  $parsed = [Guid]::Empty
  if (-not [Guid]::TryParse($text, [ref]$parsed)) {
    throw "$Name must be a GUID: $text"
  }

  return $parsed.Guid
}

function Get-UriValue {
  param(
    [string]$Value,
    [string]$Name
  )

  $text = Get-RequiredValue -Value $Value -Name $Name
  $uri = $null
  if (-not [Uri]::TryCreate($text, [UriKind]::Absolute, [ref]$uri)) {
    throw "Invalid URI for ${Name}: $text"
  }

  if ($uri.Scheme -ne "https") {
    throw "$Name must be an absolute https URL: $text"
  }

  return $uri
}

function Get-DialogSizeValue {
  param([string]$Value)

  $candidate = if ([string]::IsNullOrWhiteSpace($Value)) {
    "large"
  } else {
    $Value.Trim().ToLowerInvariant()
  }

  if ($candidate -notin @("small", "medium", "large")) {
    throw "DIALOG_SIZE must be one of: small, medium, large"
  }

  return $candidate
}

function Get-BooleanText {
  param([string]$Value)

  if ([string]::IsNullOrWhiteSpace($Value)) {
    return "true"
  }

  switch ($Value.Trim().ToLowerInvariant()) {
    "1" { return "true" }
    "true" { return "true" }
    "yes" { return "true" }
    "on" { return "true" }
    "0" { return "false" }
    "false" { return "false" }
    "no" { return "false" }
    "off" { return "false" }
    default { throw "AUTO_OPEN_ON_LOAD must be true or false" }
  }
}

function Get-AccentColorValue {
  param([string]$Value)

  $candidate = if ([string]::IsNullOrWhiteSpace($Value)) {
    "#2563EB"
  } else {
    $Value.Trim()
  }

  if (-not $candidate.StartsWith("#")) {
    $candidate = "#$candidate"
  }

  if ($candidate -notmatch '^#[0-9A-Fa-f]{6}$') {
    throw "ACCENT_COLOR must be a 6 digit hex color like #2563EB"
  }

  return $candidate.ToUpperInvariant()
}

function Get-PackageNameValue {
  param(
    [string]$Value,
    [string]$FallbackSuffix
  )

  $candidate = if ([string]::IsNullOrWhiteSpace($Value)) {
    "com.m365.outlook.sidebar.$FallbackSuffix"
  } else {
    $Value.Trim().ToLowerInvariant()
  }

  if ($candidate -notmatch '^[a-z0-9][a-z0-9.-]*[a-z0-9]$') {
    throw "PACKAGE_NAME must use lowercase letters, digits, dots, or hyphens"
  }

  return $candidate
}

function Get-EntityIdValue {
  param(
    [string]$Value,
    [string]$FallbackSuffix
  )

  $candidate = if ([string]::IsNullOrWhiteSpace($Value)) {
    "tab-$FallbackSuffix"
  } else {
    $Value.Trim()
  }

  if ($candidate.Length -gt 64) {
    throw "ENTITY_ID must be 64 characters or fewer"
  }

  return $candidate
}

function Convert-ToManifestText {
  param(
    [string]$Template,
    [hashtable]$Values
  )

  $text = $Template
  foreach ($key in $Values.Keys) {
    $text = $text.Replace("__${key}__", [string]$Values[$key])
  }

  return $text
}

function Join-AppUrl {
  param(
    [Uri]$BaseUri,
    [string]$RelativePath
  )

  return "{0}/{1}" -f $BaseUri.AbsoluteUri.TrimEnd('/'), $RelativePath.TrimStart('/')
}

$repoInfo = Get-GitHubRepositoryInfo
if ([string]::IsNullOrWhiteSpace($AppId)) {
  $AppId = "a87ae817-082c-4707-9783-5fbf5b0f541f"
}
if ([string]::IsNullOrWhiteSpace($AppBaseUrl) -and $repoInfo) {
  $AppBaseUrl = $repoInfo.ContentBaseUrl
}
if ([string]::IsNullOrWhiteSpace($TargetUrl)) {
  $TargetUrl = "https://www.ctc-g.co.jp/"
}
if ([string]::IsNullOrWhiteSpace($AppName)) {
  $AppName = Get-DefaultAppName
}
if ([string]::IsNullOrWhiteSpace($AppDescription)) {
  $AppDescription = $AppName
}
if ([string]::IsNullOrWhiteSpace($DeveloperName) -and $repoInfo) {
  $DeveloperName = $repoInfo.Owner
}
if ([string]::IsNullOrWhiteSpace($DeveloperName)) {
  $DeveloperName = "z2250389"
}
if ([string]::IsNullOrWhiteSpace($DeveloperWebsiteUrl) -and $repoInfo) {
  $DeveloperWebsiteUrl = $repoInfo.RepositoryUrl
}
if ([string]::IsNullOrWhiteSpace($DeveloperPrivacyUrl) -and $repoInfo) {
  $DeveloperPrivacyUrl = $repoInfo.RepositoryUrl
}
if ([string]::IsNullOrWhiteSpace($DeveloperTermsUrl) -and $repoInfo) {
  $DeveloperTermsUrl = $repoInfo.RepositoryUrl
}

$appIdValue = Get-GuidValue -Value $AppId -Name "APP_ID"
$targetUri = Get-UriValue -Value $TargetUrl -Name "TARGET_URL"
$appNameValue = Get-RequiredValue -Value $AppName -Name "APP_NAME"
$appVersionValue = if ([string]::IsNullOrWhiteSpace($AppVersion)) { "1.0.0" } else { $AppVersion.Trim() }
$appDescriptionValue = if ([string]::IsNullOrWhiteSpace($AppDescription)) { $appNameValue } else { $AppDescription.Trim() }
$dialogTitleValue = if ([string]::IsNullOrWhiteSpace($DialogTitle)) { $appNameValue } else { $DialogTitle.Trim() }
$dialogSizeValue = Get-DialogSizeValue -Value $DialogSize
$autoOpenOnLoadValue = Get-BooleanText -Value $AutoOpenOnLoad
$accentColorValue = Get-AccentColorValue -Value $AccentColor
$idSuffix = $appIdValue.Substring(0, 8).ToLowerInvariant()
$packageNameValue = Get-PackageNameValue -Value $PackageName -FallbackSuffix $idSuffix
$entityIdValue = Get-EntityIdValue -Value $EntityId -FallbackSuffix $idSuffix
$developerNameValue = Get-RequiredValue -Value $DeveloperName -Name "DEVELOPER_NAME"
$developerWebsiteUri = Get-UriValue -Value $DeveloperWebsiteUrl -Name "DEVELOPER_WEBSITE_URL"
$developerPrivacyUri = Get-UriValue -Value $DeveloperPrivacyUrl -Name "DEVELOPER_PRIVACY_URL"
$developerTermsUri = Get-UriValue -Value $DeveloperTermsUrl -Name "DEVELOPER_TERMS_URL"
$contentUrl = $targetUri.AbsoluteUri

$template = Get-Content -LiteralPath $TemplatePath -Raw -Encoding UTF8
$values = @{
  APP_ID = $appIdValue
  APP_VERSION = $appVersionValue
  PACKAGE_NAME = $packageNameValue
  DEVELOPER_NAME = $developerNameValue
  DEVELOPER_WEBSITE_URL = $developerWebsiteUri.AbsoluteUri
  DEVELOPER_PRIVACY_URL = $developerPrivacyUri.AbsoluteUri
  DEVELOPER_TERMS_URL = $developerTermsUri.AbsoluteUri
  APP_NAME = $appNameValue
  APP_DESCRIPTION = $appDescriptionValue
  ACCENT_COLOR = $accentColorValue
  ENTITY_ID = $entityIdValue
  CONTENT_URL = $contentUrl
  WEBSITE_URL = $targetUri.AbsoluteUri
  APP_DOMAIN = $targetUri.Host
  TARGET_DOMAIN = $targetUri.Host
}

$manifestText = Convert-ToManifestText -Template $template -Values $values
$outputDir = Split-Path -Parent $OutputPath
if (-not (Test-Path -LiteralPath $outputDir)) {
  New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

Set-Content -LiteralPath $OutputPath -Value $manifestText -Encoding UTF8
Write-Host "Manifest written to $OutputPath"
