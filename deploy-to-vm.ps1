# Deploys NinjaClaw-fork build artifacts to vm-ninjaclaw-nano in chunks
# (Az RunCommand has ~256KB script-string limit; tarball b64 = ~310KB).

$ErrorActionPreference = 'Stop'

$tar = "C:\Users\gavishofir\source\NinjaClaw-fork\deploy.tar.gz"
if (-not (Test-Path $tar)) { throw "tarball not found: $tar" }

$bytes = [IO.File]::ReadAllBytes($tar)
$b64   = [Convert]::ToBase64String($bytes)
Write-Host "Tarball: $($bytes.Length) bytes, b64: $($b64.Length) chars"

$chunkSize = 90000
$chunks = @()
for ($i = 0; $i -lt $b64.Length; $i += $chunkSize) {
    $chunks += $b64.Substring($i, [Math]::Min($chunkSize, $b64.Length - $i))
}
Write-Host "Sending in $($chunks.Count) chunks"

$rg = "Agents-RG"
$vm = "vm-ninjaclaw-nano"

function Invoke-VMScript($scriptText) {
    $tmp = [IO.Path]::GetTempFileName()
    Set-Content -Path $tmp -NoNewline -Value $scriptText
    $r = Invoke-AzVMRunCommand -ResourceGroupName $rg -VMName $vm -CommandId "RunShellScript" -ScriptPath $tmp
    Remove-Item $tmp
    return ($r.Value | ForEach-Object { $_.Message }) -join "`n"
}

# Step 1: truncate b64 file with chunk 0
Write-Host "Sending chunk 1/$($chunks.Count)..."
$out = Invoke-VMScript ("#!/bin/bash`nset -e`nrm -f /tmp/deploy.b64`ncat > /tmp/deploy.b64 << 'B64EOF'`n" + $chunks[0] + "`nB64EOF`necho chunk1-len-`$(wc -c < /tmp/deploy.b64)")
Write-Host $out

# Step 2..N: append remaining chunks
for ($i = 1; $i -lt $chunks.Count; $i++) {
    Write-Host "Sending chunk $($i+1)/$($chunks.Count)..."
    $out = Invoke-VMScript ("#!/bin/bash`nset -e`ncat >> /tmp/deploy.b64 << 'B64EOF'`n" + $chunks[$i] + "`nB64EOF`necho chunk-$($i+1)-len-`$(wc -c < /tmp/deploy.b64)")
    Write-Host $out
}

# Step final: decode, extract, npm install
Write-Host "Decoding + extracting + npm install..."
$finalScript = @'
#!/bin/bash
set -e
base64 -d /tmp/deploy.b64 > /tmp/deploy.tar.gz
ls -la /tmp/deploy.tar.gz
sudo -u azureuser tar -xzf /tmp/deploy.tar.gz -C /home/azureuser/ninjaclaw-nano
rm /tmp/deploy.b64 /tmp/deploy.tar.gz
echo EXTRACTED_OK
sudo -u azureuser bash -c 'cd /home/azureuser/ninjaclaw-nano && npm install --omit=dev --ignore-scripts 2>&1 | tail -10'
echo NPM_INSTALL_DONE
'@
$out = Invoke-VMScript $finalScript
Write-Host $out

Write-Host "Restarting service..."
$restart = @'
#!/bin/bash
sudo -u azureuser bash -c 'export XDG_RUNTIME_DIR=/run/user/$(id -u); systemctl --user restart ninjaclaw-nano.service; sleep 4; systemctl --user is-active ninjaclaw-nano.service; journalctl --user -u ninjaclaw-nano.service -n 20 --no-pager'
'@
$out = Invoke-VMScript $restart
Write-Host $out

Write-Host "DEPLOY COMPLETE"
