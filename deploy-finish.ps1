$ErrorActionPreference = 'Stop'
$rg = "Agents-RG"; $vm = "vm-ninjaclaw-nano"

function Invoke-VMScript($scriptText) {
    $tmp = [IO.Path]::GetTempFileName()
    Set-Content -Path $tmp -NoNewline -Value $scriptText
    $r = Invoke-AzVMRunCommand -ResourceGroupName $rg -VMName $vm -CommandId "RunShellScript" -ScriptPath $tmp
    Remove-Item $tmp
    return ($r.Value | ForEach-Object { $_.Message }) -join "`n"
}

Write-Host "=== Step 1: cleanup stale /tmp files + verify b64 still here ==="
$out = Invoke-VMScript @'
#!/bin/bash
ls -la /tmp/deploy.b64 2>&1 | head -3
ls -la /tmp/deploy.tar.gz 2>&1 | head -3
rm -f /tmp/deploy.tar.gz
echo CLEANED
ls -la /tmp/deploy.b64 2>&1 | head -3
'@
Write-Host $out

Write-Host "=== Step 2: decode + extract + npm install ==="
$out = Invoke-VMScript @'
#!/bin/bash
set -e
if [ ! -f /tmp/deploy.b64 ]; then echo "ERROR: /tmp/deploy.b64 missing — re-run full deploy-to-vm.ps1"; exit 1; fi
base64 -d /tmp/deploy.b64 > /tmp/deploy.tar.gz
chown azureuser:azureuser /tmp/deploy.tar.gz
ls -la /tmp/deploy.tar.gz
sudo -u azureuser tar -xzf /tmp/deploy.tar.gz -C /home/azureuser/ninjaclaw-nano
echo EXTRACTED_OK
ls -la /home/azureuser/ninjaclaw-nano/dist/agent365/admin/ 2>&1 | head -10
ls -la /home/azureuser/ninjaclaw-nano/web_static/agent365.html 2>&1
sudo -u azureuser bash -c 'cd /home/azureuser/ninjaclaw-nano && npm install --omit=dev 2>&1 | tail -5'
echo NPM_INSTALL_DONE
rm -f /tmp/deploy.b64 /tmp/deploy.tar.gz
'@
Write-Host $out

Write-Host "=== Step 3: restart service ==="
$out = Invoke-VMScript @'
#!/bin/bash
sudo -u azureuser bash -c 'export XDG_RUNTIME_DIR=/run/user/$(id -u); systemctl --user restart ninjaclaw-nano.service; sleep 4; systemctl --user is-active ninjaclaw-nano.service; journalctl --user -u ninjaclaw-nano.service --since "30 seconds ago" --no-pager | tail -25'
'@
Write-Host $out

Write-Host "=== DONE ==="
