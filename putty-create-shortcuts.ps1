<#
Script to create PuTTY windows .lnk files in current folder.

It may be easier to start SSH sessions using shortcuts.
#>

$sessions = Get-ChildItem HKCU:SOFTWARE\SimonTatham\PuTTY\Sessions
$cur = Get-Location
$Wsh = New-Object -comObject WScript.Shell
foreach($session in $sessions) {
	$name = $session.name | Split-Path -Leaf
	$path = (Join-Path -Path $cur -ChildPath $name).ToString()
	$name = [system.uri]::UnescapeDataString($name)
	$path = [system.uri]::UnescapeDataString($path)
	#echo "${name}: ${path}"
	$shortcut = $Wsh.CreateShortcut("${path}.lnk")
	$shortcut.TargetPath = "putty.exe"
	$shortcut.Arguments = "-load `"${name}`""
	$shortcut.Save()
}
