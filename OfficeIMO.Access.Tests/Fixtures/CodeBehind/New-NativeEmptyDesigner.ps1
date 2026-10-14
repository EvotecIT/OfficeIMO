param([Parameter(Mandatory)][string]$OutputPath)
$ErrorActionPreference='Stop'; if(Test-Path -LiteralPath $OutputPath){throw 'Output exists'}
$application=$null; $opened=$false
try {
    $application=New-Object -ComObject Access.Application; $application.Visible=$false;$application.AutomationSecurity=3
    if([int]$application.AutomationSecurity -ne 3){throw 'Macro disabling was not accepted'}
    $application.NewCurrentDatabase($OutputPath,12);$opened=$true
    $form=$application.CreateForm();$oldName=[string]$form.Name
    $label=$application.CreateControl($oldName,100,0,'','',200,200,2000,300)
    $label.Name='EventLabel';$label.Caption='Synthetic inert label'
    if($form.HasModule){throw 'Fixture unexpectedly contains code-behind'}
    $application.DoCmd.Save(2,$oldName);$application.DoCmd.Close(2,$oldName,1)
    [Runtime.InteropServices.Marshal]::FinalReleaseComObject($label)|Out-Null
    [Runtime.InteropServices.Marshal]::FinalReleaseComObject($form)|Out-Null
    $application.DoCmd.Rename('EventForm',2,$oldName)
    $application.CloseCurrentDatabase();$opened=$false
} finally {if($application){if($opened){$application.CloseCurrentDatabase()};$application.Quit(2);[Runtime.InteropServices.Marshal]::FinalReleaseComObject($application)|Out-Null}}
Get-FileHash -LiteralPath $OutputPath | Select-Object Hash
