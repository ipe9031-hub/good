$ErrorActionPreference = 'Stop'
Set-Location (Join-Path $PSScriptRoot '../..')
$dotnetExe = if ($env:DOTNET) { $env:DOTNET } else { 'dotnet' }
$probeKey = if ($env:PROBE_SIGNING_KEY) { $env:PROBE_SIGNING_KEY } else { Join-Path $env:LOCALAPPDATA 'PPTAIStudio/Build/native-probe.snk' }
$env:DOTNET_CLI_TELEMETRY_OPTOUT = '1'
$env:DOTNET_GENERATE_ASPNET_CERTIFICATE = 'false'
if (!(Test-Path $probeKey)) {
  New-Item -ItemType Directory -Path (Split-Path $probeKey) -Force | Out-Null
  $rsa = New-Object System.Security.Cryptography.RSACryptoServiceProvider(2048)
  $rsa.PersistKeyInCsp = $false
  try { [IO.File]::WriteAllBytes($probeKey, $rsa.ExportCspBlob($true)) } finally { $rsa.Dispose() }
}
foreach ($project in @('Probe.Setup', 'Probe.Checks')) {
  & $dotnetExe restore "native/$project/$project.csproj" --locked-mode "-p:ProbeSigningKey=$probeKey"
  if ($LASTEXITCODE -ne 0) { throw 'Restore failed' }
}
& $dotnetExe build native/Probe.Setup/Probe.Setup.csproj -c Release --no-restore "-p:ProbeSigningKey=$probeKey"
if ($LASTEXITCODE -ne 0) { throw 'Build failed' }
& $dotnetExe run --project native/Probe.Checks/Probe.Checks.csproj -c Release --no-restore "-p:ProbeSigningKey=$probeKey" -- native/Probe.Setup/bin/Release/net48
if ($LASTEXITCODE -ne 0) { throw 'Checks failed' }
