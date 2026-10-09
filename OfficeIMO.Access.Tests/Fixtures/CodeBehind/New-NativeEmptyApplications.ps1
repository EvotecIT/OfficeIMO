param([Parameter(Mandatory)][string]$OutputRoot)
$ErrorActionPreference='Stop'
New-Item -ItemType Directory -Path $OutputRoot | Out-Null
$module=Join-Path $OutputRoot 'FirstModule.bas'
[IO.File]::WriteAllText($module,"Attribute VB_Name = `"FirstModule`"`r`nOption Explicit`r`nPublic Function FirstValue() As Long`r`n FirstValue = 47`r`nEnd Function`r`n",[Text.Encoding]::ASCII)
$application=$null
try {
    $application=New-Object -ComObject Access.Application
    $application.Visible=$false; $application.AutomationSecurity=3
    foreach($profile in @(@{Name='first-jet4.mdb';Format=9},@{Name='first-ace12.accdb';Format=12})) {
        $path=Join-Path $OutputRoot $profile.Name
        $application.NewCurrentDatabase($path,$profile.Format)
        $application.CloseCurrentDatabase()
        Copy-Item -LiteralPath $path -Destination ($path+'.empty'+[IO.Path]::GetExtension($path))
        $application.OpenCurrentDatabase($path,$false)
        try {$application.LoadFromText(5,'FirstModule',$module)} finally {$application.CloseCurrentDatabase()}
    }
} finally {if($application){$application.Quit(2);[Runtime.InteropServices.Marshal]::FinalReleaseComObject($application)|Out-Null}}
