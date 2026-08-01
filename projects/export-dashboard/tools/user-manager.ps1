param(
  [Parameter(Mandatory=$true)]
  [ValidateSet("add", "get", "delete", "list", "invite")]
  [string]$Action,

  [string]$Email,
  [string]$Name,
  [string]$Org,
  [string]$Password
)

$ErrorActionPreference = "Stop"

function Load-DotEnv {
  param([string]$Path)
  if (!(Test-Path -LiteralPath $Path)) { return }
  Get-Content -LiteralPath $Path | ForEach-Object {
    $line = $_.Trim()
    if (!$line -or $line.StartsWith("#")) { return }
    $idx = $line.IndexOf("=")
    if ($idx -lt 1) { return }
    $key = $line.Substring(0, $idx).Trim()
    $value = $line.Substring($idx + 1).Trim().Trim('"').Trim("'")
    [Environment]::SetEnvironmentVariable($key, $value, "Process")
  }
}

function New-TempPassword {
  $chars = "ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789".ToCharArray()
  $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
  $bytes = New-Object byte[] 6
  $rng.GetBytes($bytes)
  -join ($bytes | ForEach-Object { $chars[$_ % $chars.Length] })
}

function Invoke-Upstash {
  param([array]$Command)
  $url = $env:UPSTASH_REDIS_REST_URL
  $token = $env:UPSTASH_REDIS_REST_TOKEN
  if (!$url -or !$token) {
    throw "Missing UPSTASH_REDIS_REST_URL or UPSTASH_REDIS_REST_TOKEN. Create tools\.env from tools\.env.example first."
  }
  $body = $Command | ConvertTo-Json -Compress -Depth 20
  Invoke-RestMethod `
    -Method Post `
    -Uri $url `
    -Headers @{ Authorization = "Bearer $token"; "Content-Type" = "application/json; charset=utf-8" } `
    -Body $body `
    -TimeoutSec 60
}

function Normalize-Email {
  param([string]$Value)
  if (!$Value) { throw "Email is required." }
  $Value.Trim().ToLowerInvariant()
}

function Parse-RedisUser {
  param($Value)
  if ($null -eq $Value) { return $null }
  if ($Value -is [string]) { return $Value | ConvertFrom-Json }
  $Value
}

function Print-Invite {
  param([string]$ToEmail, [string]$ToPassword, [string]$ToName)
  $familyName = if ($ToName) { $ToName.Substring(0, 1) } else { "" }
  $titleName = if ($familyName) { "$familyName`总" } else { "老师" }
  @"
$titleName好，

感谢您参加我们出海数据跟踪网站的内测，如国内访问遇到问题请连接梯子访问。

内测网址：https://dashboard.gfautodata.cn
登录账号：$ToEmail
登录密码：$ToPassword

出海数据网站包括中汽协乘用车出口、海关出口（分区域、分国家、分车企、分车型；独家数据）、Marklines终端实销情况等，更多功能陆续上线，数据月度更新。

欢迎多多提问题，我将进一步完善网站开发，将内测反馈意见修改后正式推出跟踪网站。

备注：请勿传播网站内数据，最终解释权归广发汽车。
"@
}

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Load-DotEnv (Join-Path $scriptDir ".env")

switch ($Action) {
  "add" {
    $emailNorm = Normalize-Email $Email
    if (!$Name) { throw "Name is required for add." }
    if (!$Org) { throw "Org is required for add." }
    if (!$Password) { $Password = New-TempPassword }

    $existing = Invoke-Upstash @("GET", "user:$emailNorm")
    if ($existing.result) {
      Write-Host "User already exists: $emailNorm"
      $user = Parse-RedisUser $existing.result
      Write-Host "Name: $($user.name)"
      Write-Host "Org : $($user.org)"
      Write-Host "Password: $($user.password)"
      exit 0
    }

    $record = [ordered]@{
      password = $Password
      name = $Name
      org = $Org
      createdAt = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    }
    $value = $record | ConvertTo-Json -Compress
    $resp = Invoke-Upstash @("SET", "user:$emailNorm", $value)
    if ($resp.result -ne "OK") { throw "Upstash SET failed: $($resp | ConvertTo-Json -Compress)" }

    Write-Host "Added: $emailNorm"
    Write-Host "Password: $Password"
    Write-Host ""
    Print-Invite $emailNorm $Password $Name
  }

  "get" {
    $emailNorm = Normalize-Email $Email
    $resp = Invoke-Upstash @("GET", "user:$emailNorm")
    $user = Parse-RedisUser $resp.result
    if (!$user) {
      Write-Host "Not found: $emailNorm"
      exit 1
    }
    [pscustomobject]@{
      email = $emailNorm
      name = $user.name
      org = $user.org
      password = $user.password
      createdAt = $user.createdAt
    } | Format-List
  }

  "delete" {
    $emailNorm = Normalize-Email $Email
    $resp = Invoke-Upstash @("DEL", "user:$emailNorm")
    Write-Host "Deleted count: $($resp.result)"
  }

  "list" {
    $keysResp = Invoke-Upstash @("KEYS", "user:*")
    $rows = @()
    foreach ($key in @($keysResp.result)) {
      $resp = Invoke-Upstash @("GET", $key)
      $user = Parse-RedisUser $resp.result
      if ($user) {
        $rows += [pscustomobject]@{
          email = $key.Replace("user:", "")
          name = $user.name
          org = $user.org
          createdAt = $user.createdAt
        }
      }
    }
    $rows | Sort-Object email | Format-Table -AutoSize
  }

  "invite" {
    $emailNorm = Normalize-Email $Email
    $resp = Invoke-Upstash @("GET", "user:$emailNorm")
    $user = Parse-RedisUser $resp.result
    if (!$user) {
      Write-Host "Not found: $emailNorm"
      exit 1
    }
    Print-Invite $emailNorm $user.password $user.name
  }
}
