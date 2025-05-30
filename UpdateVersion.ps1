param(
    [ValidateNotNullOrEmpty()]$ProjectDir = $Env:GITHUB_WORKSPACE,
    [ValidateNotNullOrEmpty()][System.String]$BuildNumber = $Env:GITHUB_RUN_NUMBER
)
$manifest=join-path -Path $ProjectDir  -ChildPath ".\DOrcDeployModule.psd1"
write-host "Using manifest: " $manifest
$version=[version]$(Import-PowerShellDataFile $manifest).ModuleVersion
if (-not $BuildNumber)
{
    write-host "Build not provided, incrementing current Revision"
    $newVersion = "{0}.{1}.{2}.{3}" -f $version.Major, $version.Minor, $version.Build, ($version.Revision + 1)
}else{
    $regexVersion = [regex]::matches($BuildNumber,"\d+\.\d+\.\d+\.\d+")
	$newVersion=[version]$regexVersion.Value
}
write-host "New version: " $newVersion
Update-ModuleManifest -Path $manifest -ModuleVersion $newVersion
