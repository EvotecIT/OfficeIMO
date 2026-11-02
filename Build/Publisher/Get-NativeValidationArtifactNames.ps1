function Get-NativeValidationArtifactNames {
    <#
    .SYNOPSIS
    Returns every evidence filename planned for a native validation stage.
    .DESCRIPTION
    Shared by the controller and isolated worker before native activation, so
    stage-specific inventories and outputs use the same overwrite boundary.
    .PARAMETER Stage
    Inspect opens and inventories; Export also writes a read-only PDF reference;
    Resave saves and reopens; Full also
    exports both PDFs. Full preserves the default validation operation.
    .OUTPUTS
    Evidence filenames relative to the task-owned output directory.
    #>
    param(
        [ValidateSet('Inspect', 'Export', 'Resave', 'Full')]
        [string] $Stage = 'Full'
    )

    $names = @('candidate-input.pub', 'native-initial-inventory.json')
    if ($Stage -in @('Resave', 'Full')) {
        $names += @('native-resaved.pub', 'native-resaved-inventory.json')
    }
    if ($Stage -in @('Export', 'Full')) {
        $names += 'native-initial.pdf'
    }
    if ($Stage -eq 'Full') {
        $names += 'native-resaved.pdf'
    }
    return $names + @('native-validation.json')
}
