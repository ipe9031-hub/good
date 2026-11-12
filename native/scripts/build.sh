#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$0")/../.."
dotnet_bin="${DOTNET:-dotnet}"
probe_key="${PROBE_SIGNING_KEY:-${XDG_DATA_HOME:-$HOME/.local/share}/ppt-ai-studio/native-probe.snk}"
export DOTNET_CLI_TELEMETRY_OPTOUT=1 DOTNET_GENERATE_ASPNET_CERTIFICATE=false
python3 native/scripts/keygen.py "$probe_key"
for project in Probe.Setup Probe.Checks; do
  "$dotnet_bin" restore "native/$project/$project.csproj" --locked-mode -p:ProbeSigningKey="$probe_key"
done
"$dotnet_bin" build native/Probe.Setup/Probe.Setup.csproj -c Release --no-restore -p:ProbeSigningKey="$probe_key"
"$dotnet_bin" run --project native/Probe.Checks/Probe.Checks.csproj -c Release --no-restore -p:ProbeSigningKey="$probe_key" -- native/Probe.Setup/bin/Release/net48
