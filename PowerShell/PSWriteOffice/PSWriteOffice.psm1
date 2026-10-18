# to speed up development adding direct path to binaries, instead of the the Lib folder
$DevelopmentBasePath = Join-Path (Join-Path (Join-Path $PSScriptRoot 'Sources') 'PSWriteOffice') 'bin'
$DevelopmentConfiguration = if ($env:PSWRITEOFFICE_DEVELOPMENT_CONFIGURATION -in @('Debug', 'Release')) {
    $env:PSWRITEOFFICE_DEVELOPMENT_CONFIGURATION
} elseif (Test-Path (Join-Path $DevelopmentBasePath 'Release')) {
    'Release'
} else {
    'Debug'
}
$DevelopmentPath = Join-Path $DevelopmentBasePath $DevelopmentConfiguration
$DevelopmentFolderCore = "net8.0"
$DevelopmentFolderDefault = "net472"
$DevelopmentFramework = if ($PSVersionTable.PSEdition -eq 'Core') {
    $DevelopmentFolderCore
} else {
    $DevelopmentFolderDefault
}
$DevelopmentBinaryPath = Join-Path (Join-Path $DevelopmentPath $DevelopmentFramework) 'PSWriteOffice.dll'
$Development = if ($env:PSWRITEOFFICE_USE_DEVELOPMENT_BINARIES -eq 'false') {
    $false
} else {
    Test-Path $DevelopmentBinaryPath
}
$BinaryModules = @(
    "PSWriteOffice.dll"
)
$AssemblyFolders = Get-ChildItem -Path (Join-Path $PSScriptRoot 'Lib') -Directory -ErrorAction SilentlyContinue

function Import-PSWriteOfficeDevelopmentBinaryModule {
    param(
        [Parameter(Mandatory)]
        [string] $Path
    )

    $loaderTypeName = 'PSWriteOffice.DevelopmentModuleLoadContext.ModuleAssemblyLoadContext'
    if (-not ($loaderTypeName -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Loader;

namespace PSWriteOffice.DevelopmentModuleLoadContext;

public sealed class ModuleAssemblyLoadContext : AssemblyLoadContext
{
    private static readonly object Sync = new();
    private static readonly Dictionary<string, ModuleAssemblyLoadContext> Contexts = new(StringComparer.OrdinalIgnoreCase);
    private readonly string _assemblyDirectory;
    private readonly string _moduleAssemblyPath;
    private readonly AssemblyDependencyResolver _resolver;
    private Assembly _moduleAssembly;

    private ModuleAssemblyLoadContext(string moduleAssemblyPath, string contextName) : base(contextName, isCollectible: false)
    {
        _moduleAssemblyPath = Path.GetFullPath(moduleAssemblyPath);
        _assemblyDirectory = Path.GetDirectoryName(_moduleAssemblyPath) ?? string.Empty;
        _resolver = new AssemblyDependencyResolver(_moduleAssemblyPath);
    }

    public static Assembly LoadModule(string moduleAssemblyPath, string contextName)
    {
        if (string.IsNullOrWhiteSpace(moduleAssemblyPath))
        {
            throw new ArgumentException("Module assembly path is required.", nameof(moduleAssemblyPath));
        }

        string fullPath = Path.GetFullPath(moduleAssemblyPath);
        if (!File.Exists(fullPath))
        {
            throw new FileNotFoundException("Module assembly was not found.", fullPath);
        }

        lock (Sync)
        {
            if (!Contexts.TryGetValue(fullPath, out ModuleAssemblyLoadContext context))
            {
                context = new ModuleAssemblyLoadContext(fullPath, string.IsNullOrWhiteSpace(contextName) ? Path.GetFileNameWithoutExtension(fullPath) : contextName);
                Contexts[fullPath] = context;
            }

            return context.LoadMainModule();
        }
    }

    protected override Assembly Load(AssemblyName assemblyName)
    {
        if (assemblyName == null || string.IsNullOrWhiteSpace(assemblyName.Name))
        {
            return null;
        }

        AssemblyName loaderAssembly = typeof(ModuleAssemblyLoadContext).Assembly.GetName();
        if (AssemblyName.ReferenceMatchesDefinition(loaderAssembly, assemblyName))
        {
            return typeof(ModuleAssemblyLoadContext).Assembly;
        }

        if (string.Equals(assemblyName.Name, "System.Management.Automation", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        string assemblyPath = _resolver.ResolveAssemblyToPath(assemblyName);
        if (!string.IsNullOrWhiteSpace(assemblyPath) && File.Exists(assemblyPath))
        {
            return LoadFromAssemblyPath(assemblyPath);
        }

        string fallbackPath = Path.Combine(_assemblyDirectory, assemblyName.Name + ".dll");
        return File.Exists(fallbackPath) ? LoadFromAssemblyPath(fallbackPath) : null;
    }

    protected override IntPtr LoadUnmanagedDll(string unmanagedDllName)
    {
        string libraryPath = _resolver.ResolveUnmanagedDllToPath(unmanagedDllName);
        if (libraryPath != null)
        {
            return LoadUnmanagedDllFromPath(libraryPath);
        }

        return IntPtr.Zero;
    }

    private Assembly LoadMainModule()
    {
        if (_moduleAssembly == null)
        {
            _moduleAssembly = LoadFromAssemblyPath(_moduleAssemblyPath);
        }

        return _moduleAssembly;
    }
}
'@ -ErrorAction Stop
    }

    $importModule = Get-Command -Name Import-Module -Module Microsoft.PowerShell.Core
    $moduleAssembly = [PSWriteOffice.DevelopmentModuleLoadContext.ModuleAssemblyLoadContext]::LoadModule($Path, 'PSWriteOfficeDevelopment')
    $innerModule = & $importModule -Assembly $moduleAssembly -Force -PassThru -ErrorAction Stop

    if ($innerModule) {
        $addExportedCmdlet = [System.Management.Automation.PSModuleInfo].GetMethod(
            'AddExportedCmdlet',
            [System.Reflection.BindingFlags]'Instance, NonPublic'
        )
        if ($null -ne $addExportedCmdlet) {
            foreach ($cmdlet in $innerModule.ExportedCmdlets.Values) {
                $addExportedCmdlet.Invoke($ExecutionContext.SessionState.Module, @(, $cmdlet)) | Out-Null
            }

            $addExportedAlias = [System.Management.Automation.PSModuleInfo].GetMethod(
                'AddExportedAlias',
                [System.Reflection.BindingFlags]'Instance, NonPublic'
            )
            if ($null -ne $addExportedAlias) {
                foreach ($alias in $innerModule.ExportedAliases.Values) {
                    $aliasTarget = if ([string]::IsNullOrWhiteSpace($alias.Definition)) {
                        $alias.ResolvedCommandName
                    } else {
                        $alias.Definition
                    }

                    Set-Alias -Name $alias.Name -Value $aliasTarget -Scope Local -Force -ErrorAction Stop
                    $exportedAlias = $ExecutionContext.SessionState.InvokeCommand.GetCommand($alias.Name, [System.Management.Automation.CommandTypes]::Alias)
                    if ($null -ne $exportedAlias) {
                        $addExportedAlias.Invoke($ExecutionContext.SessionState.Module, @(, $exportedAlias)) | Out-Null
                    }
                }
            }
        } else {
            throw 'AddExportedCmdlet is not available on this PowerShell version.'
        }
    }
}

# ensure script file collections always exist (legacy folders were removed)
if (-not (Test-Path variable:Classes)) { $Classes = @() }
if (-not (Test-Path variable:Enums)) { $Enums = @() }
if (-not (Test-Path variable:Private)) { $Private = @() }
if (-not (Test-Path variable:Public)) { $Public = @() }

# Lets find which libraries we need to load
$Default = $false
$Core = $false
$Standard = $false
foreach ($A in $AssemblyFolders.Name) {
    if ($A -eq 'Default') {
        $Default = $true
    } elseif ($A -eq 'Core') {
        $Core = $true
    } elseif ($A -eq 'Standard') {
        $Standard = $true
    }
}
if ($Standard -and $Core -and $Default) {
    $FrameworkNet = 'Default'
    $Framework = 'Standard'
} elseif ($Standard -and $Core) {
    $Framework = 'Standard'
    $FrameworkNet = 'Standard'
} elseif ($Core -and $Default) {
    $Framework = 'Core'
    $FrameworkNet = 'Default'
} elseif ($Standard -and $Default) {
    $Framework = 'Standard'
    $FrameworkNet = 'Default'
} elseif ($Standard) {
    $Framework = 'Standard'
    $FrameworkNet = 'Standard'
} elseif ($Core) {
    $Framework = 'Core'
    $FrameworkNet = ''
} elseif ($Default) {
    $Framework = ''
    $FrameworkNet = 'Default'
} else {
    #Write-Error -Message 'No assemblies found'
}

# Ensure native runtime libraries are discoverable on Windows
if ($IsWindows) {
    $arch = [System.Runtime.InteropServices.RuntimeInformation]::ProcessArchitecture
    $archFolder = switch ($arch) {
        'X64' {
            'win-x64'
        }
        'X86' {
            'win-x86'
        }
        'Arm64' {
            'win-arm64'
        }
        'Arm' {
            'win-arm'
        }
        default {
            'win-x64'
        }
    }

    if ($Development) {
        $baseDir = if ($PSEdition -eq 'Core') {
            Join-Path $DevelopmentPath $DevelopmentFolderCore
        } else {
            Join-Path $DevelopmentPath $DevelopmentFolderDefault
        }
    } else {
        $baseDir = if ($PSEdition -eq 'Core') {
            Join-Path $PSScriptRoot "Lib/$Framework"
        } elseif ($FrameworkNet) {
            Join-Path $PSScriptRoot "Lib/$FrameworkNet"
        } else {
            $null
        }
    }

    if ($baseDir) {
        $runtimePath = Join-Path $baseDir "runtimes/$archFolder/native"
        if (Test-Path $runtimePath) {
            Write-Verbose -Message "Adding $runtimePath to PATH"
            $env:PATH = "$runtimePath;" + $env:PATH
        }
    }
}

$BinaryDev = if ($Development) {
    @(
        foreach ($BinaryModule in $BinaryModules) {
            if ($PSEdition -eq 'Core') {
                $Variable = Resolve-Path (Join-Path (Join-Path $DevelopmentPath $DevelopmentFolderCore) $BinaryModule)
            } else {
                $Variable = Resolve-Path (Join-Path (Join-Path $DevelopmentPath $DevelopmentFolderDefault) $BinaryModule)
            }
            $Variable
            Write-Verbose "Development mode: Using binaries from $Variable"
        }
    )
} else {
    @()
}

$ImportedBinaryModules = @()
$FoundErrors = @(
    if ($Development) {
        foreach ($BinaryModule in $BinaryDev) {
            try {
                $binaryModulePath = (Resolve-Path -LiteralPath $BinaryModule).ProviderPath
                if ($PSEdition -eq 'Core') {
                    Import-PSWriteOfficeDevelopmentBinaryModule -Path $binaryModulePath
                } else {
                    Import-Module -Name $BinaryModule -Force -ErrorAction Stop
                }
            } catch {
                Write-Warning "Failed to import module $($BinaryModule): $($_.Exception.Message)"
                $true
            }
        }
    } else {
        foreach ($BinaryModule in $BinaryModules) {
            try {
                if ($Framework -and $PSEdition -eq 'Core') {
                    $importedModule = Import-Module -Name "$PSScriptRoot\Lib\$Framework\$BinaryModule" -Force -PassThru -ErrorAction Stop
                    if ($importedModule) {
                        $ImportedBinaryModules += $importedModule
                    }
                }
                if ($FrameworkNet -and $PSEdition -ne 'Core') {
                    $importedModule = Import-Module -Name "$PSScriptRoot\Lib\$FrameworkNet\$BinaryModule" -Force -PassThru -ErrorAction Stop
                    if ($importedModule) {
                        $ImportedBinaryModules += $importedModule
                    }
                }
            } catch {
                Write-Warning "Failed to import module $($BinaryModule): $($_.Exception.Message)"
                $true
            }
        }
    }
    #Dot source the files
    foreach ($Import in @($Classes + $Enums + $Private + $Public)) {
        try {
            . $Import.Fullname
        } catch {
            Write-Error -Message "Failed to import functions from $($import.Fullname): $_"
            $true
        }
    }
)

if ($FoundErrors.Count -gt 0) {
    $ModuleName = (Get-ChildItem $PSScriptRoot\*.psd1).BaseName
    Write-Warning "Importing module $ModuleName failed. Fix errors before continuing."
    throw "Importing module $ModuleName failed. Fix errors before continuing."
    #break
}

$binaryAliases = @{
    'Compare-OfficeExcelSheet'          = 'Compare-OfficeExcelRange'
    'ConvertFrom-MarkdownHtml'          = 'ConvertFrom-OfficeMarkdownHtml'
    'ConvertFrom-PdfHtml'               = 'ConvertFrom-OfficePdfHtml'
    'ConvertFrom-Rtf'                   = 'ConvertFrom-OfficeRtf'
    'ConvertFrom-WordHtml'              = 'ConvertFrom-OfficeWordHtml'
    'ConvertFrom-WordMarkdown'          = 'ConvertFrom-OfficeWordMarkdown'
    'ConvertTo-ExcelHtml'               = 'ConvertTo-OfficeExcelHtml'
    'ConvertTo-MarkdownHtml'            = 'ConvertTo-OfficeMarkdownHtml'
    'ConvertTo-PdfExcel'                = 'ConvertTo-OfficePdfExcel'
    'ConvertTo-PdfHtml'                 = 'ConvertTo-OfficePdfHtml'
    'ConvertTo-PdfPowerPoint'           = 'ConvertTo-OfficePdfPowerPoint'
    'ConvertTo-PdfWord'                 = 'ConvertTo-OfficePdfWord'
    'ConvertTo-PowerPointHtml'          = 'ConvertTo-OfficePowerPointHtml'
    'ConvertTo-Rtf'                     = 'ConvertTo-OfficeRtf'
    'ConvertTo-VisioPng'                = 'ConvertTo-OfficeVisioPng'
    'ConvertTo-VisioSvg'                = 'ConvertTo-OfficeVisioSvg'
    'ConvertTo-WordHtml'                = 'ConvertTo-OfficeWordHtml'
    'ConvertTo-WordMarkdown'            = 'ConvertTo-OfficeWordMarkdown'
    'Edit-ExcelRow'                     = 'Edit-OfficeExcelRow'
    'ExcelAccessibility'                = 'Test-OfficeExcelAccessibility'
    'ExcelActiveSheet'                  = 'Set-OfficeExcelActiveSheet'
    'ExcelAutoFilter'                   = 'Add-OfficeExcelAutoFilter'
    'ExcelAutoFilterClear'              = 'Clear-OfficeExcelAutoFilter'
    'ExcelAutoFilterSet'                = 'Set-OfficeExcelAutoFilter'
    'ExcelAutoFit'                      = 'Invoke-OfficeExcelAutoFit'
    'ExcelCell'                         = 'Set-OfficeExcelCell'
    'ExcelChart'                        = 'Add-OfficeExcelChart'
    'ExcelChartAxis'                    = 'Set-OfficeExcelChartAxis'
    'ExcelChartPoint'                   = 'Set-OfficeExcelChartPoint'
    'ExcelChartSeries'                  = 'Set-OfficeExcelChartSeries'
    'ExcelChartTrendline'               = 'Set-OfficeExcelChartTrendline'
    'ExcelColumn'                       = 'Set-OfficeExcelColumn'
    'ExcelColumnGroup'                  = 'Set-OfficeExcelColumnGroup'
    'ExcelColumnStyle'                  = 'Set-OfficeExcelColumnStyleByHeader'
    'ExcelColumnStyleByHeader'          = 'Set-OfficeExcelColumnStyleByHeader'
    'ExcelComment'                      = 'Add-OfficeExcelComment'
    'ExcelCommentAudit'                 = 'Get-OfficeExcelCommentAudit'
    'ExcelCommentClear'                 = 'Clear-OfficeExcelComment'
    'ExcelCommentRemove'                = 'Remove-OfficeExcelComment'
    'ExcelComments'                     = 'Get-OfficeExcelComment'
    'ExcelCommentsAudit'                = 'Get-OfficeExcelCommentAudit'
    'ExcelCommentUpdate'                = 'Update-OfficeExcelComment'
    'ExcelCompare'                      = 'Compare-OfficeExcelRange'
    'ExcelConditionalColorScale'        = 'Add-OfficeExcelConditionalColorScale'
    'ExcelConditionalDataBar'           = 'Add-OfficeExcelConditionalDataBar'
    'ExcelConditionalFormatting'        = 'Get-OfficeExcelConditionalFormatting'
    'ExcelConditionalFormattingClear'   = 'Clear-OfficeExcelConditionalFormatting'
    'ExcelConditionalIconSet'           = 'Add-OfficeExcelConditionalIconSet'
    'ExcelConditionalRule'              = 'Add-OfficeExcelConditionalRule'
    'ExcelConnectionMetadata'           = 'Add-OfficeExcelPackageMetadata'
    'ExcelCsvImport'                    = 'Import-OfficeExcelDelimitedText'
    'ExcelDashboard'                    = 'New-OfficeExcelDashboard'
    'ExcelDashboardChart'               = 'Add-OfficeExcelDashboardChart'
    'ExcelDataModel'                    = 'Get-OfficeExcelDataModel'
    'ExcelDataSet'                      = 'Add-OfficeExcelDataSet'
    'ExcelDataValidation'               = 'Get-OfficeExcelDataValidation'
    'ExcelDataValidationClear'          = 'Clear-OfficeExcelDataValidation'
    'ExcelDataValidationMessage'        = 'Set-OfficeExcelDataValidationMessage'
    'ExcelDateSystem'                   = 'Set-OfficeExcelDateSystem'
    'ExcelDelimitedImport'              = 'Import-OfficeExcelDelimitedText'
    'ExcelDoctor'                       = 'Test-OfficeExcelWorkbook'
    'ExcelExecutionPolicy'              = 'Set-OfficeExcelExecutionPolicy'
    'ExcelExport'                       = 'Export-OfficeExcel'
    'ExcelFormula'                      = 'Set-OfficeExcelFormula'
    'ExcelFormulaAnalysis'              = 'Get-OfficeExcelFormulaAnalysis'
    'ExcelFormulaAudit'                 = 'Get-OfficeExcelFormulaAnalysis'
    'ExcelFreeze'                       = 'Set-OfficeExcelFreeze'
    'ExcelGridlines'                    = 'Set-OfficeExcelGridlines'
    'ExcelHeaderFooter'                 = 'Set-OfficeExcelHeaderFooter'
    'ExcelHyperlink'                    = 'Set-OfficeExcelHyperlink'
    'ExcelHyperlinkHost'                = 'Set-OfficeExcelHostHyperlink'
    'ExcelHyperlinkSmart'               = 'Set-OfficeExcelSmartHyperlink'
    'ExcelImage'                        = 'Add-OfficeExcelImage'
    'ExcelImageFromUrl'                 = 'Add-OfficeExcelImageFromUrl'
    'ExcelImport'                       = 'Import-OfficeExcel'
    'ExcelInternalLinks'                = 'Set-OfficeExcelInternalLinks'
    'ExcelInternalLinksByHeader'        = 'Set-OfficeExcelInternalLinksByHeader'
    'ExcelMargins'                      = 'Set-OfficeExcelMargins'
    'ExcelNamedRange'                   = 'Set-OfficeExcelNamedRange'
    'ExcelNamedRangeRemove'             = 'Remove-OfficeExcelNamedRange'
    'ExcelNamedRangeRename'             = 'Rename-OfficeExcelNamedRange'
    'ExcelNew'                          = 'New-OfficeExcel'
    'ExcelNumberFormatPreset'           = 'Get-OfficeExcelNumberFormatPreset'
    'ExcelOrientation'                  = 'Set-OfficeExcelOrientation'
    'ExcelPackageCopy'                  = 'Copy-OfficeExcelWorkbook'
    'ExcelPackageMetadata'              = 'Add-OfficeExcelPackageMetadata'
    'ExcelPageBreak'                    = 'Add-OfficeExcelPageBreak'
    'ExcelPageBreakClear'               = 'Clear-OfficeExcelPageBreak'
    'ExcelPageBreaks'                   = 'Get-OfficeExcelPageBreak'
    'ExcelPageSetup'                    = 'Set-OfficeExcelPageSetup'
    'ExcelPivotTable'                   = 'Add-OfficeExcelPivotTable'
    'ExcelPivotTables'                  = 'Get-OfficeExcelPivotTable'
    'ExcelPowerQuery'                   = 'Get-OfficeExcelDataModel'
    'ExcelPowerQueryMetadata'           = 'Add-OfficeExcelPowerQueryMetadata'
    'ExcelPreflight'                    = 'Get-OfficeExcelPreflight'
    'ExcelPrintArea'                    = 'Set-OfficeExcelPrintArea'
    'ExcelPrintLayout'                  = 'Set-OfficeExcelPrintLayout'
    'ExcelPrintTitles'                  = 'Set-OfficeExcelPrintTitles'
    'ExcelProtect'                      = 'Protect-OfficeExcelSheet'
    'ExcelQueryMetadata'                = 'Add-OfficeExcelPowerQueryMetadata'
    'ExcelRangeClear'                   = 'Clear-OfficeExcelRange'
    'ExcelRefreshOnOpen'                = 'Set-OfficeExcelRefreshOnOpen'
    'ExcelRepair'                       = 'Repair-OfficeExcelWorkbook'
    'ExcelReportCallout'                = 'Add-OfficeExcelReportCallout'
    'ExcelReportKpiRow'                 = 'Add-OfficeExcelReportKpiRow'
    'ExcelReportLegend'                 = 'Add-OfficeExcelReportLegend'
    'ExcelReportParagraph'              = 'Add-OfficeExcelReportParagraph'
    'ExcelReportSection'                = 'Add-OfficeExcelReportSection'
    'ExcelReportSheet'                  = 'Add-OfficeExcelReportSheet'
    'ExcelReportSpacer'                 = 'Add-OfficeExcelReportSpacer'
    'ExcelReportTable'                  = 'Add-OfficeExcelReportTable'
    'ExcelReportTitle'                  = 'Add-OfficeExcelReportTitle'
    'ExcelRichText'                     = 'Set-OfficeExcelRichText'
    'ExcelRichTextRuns'                 = 'Get-OfficeExcelRichText'
    'ExcelRow'                          = 'Set-OfficeExcelRow'
    'ExcelRowEdit'                      = 'Edit-OfficeExcelRow'
    'ExcelRowGroup'                     = 'Set-OfficeExcelRowGroup'
    'ExcelRuntimePreflight'             = 'Get-OfficeExcelRuntimePreflight'
    'ExcelSheet'                        = 'Add-OfficeExcelSheet'
    'ExcelSheetCopy'                    = 'Copy-OfficeExcelSheet'
    'ExcelSheetJoin'                    = 'Join-OfficeExcelSheet'
    'ExcelSheetMerge'                   = 'Join-OfficeExcelSheet'
    'ExcelSheetOrder'                   = 'Move-OfficeExcelSheet'
    'ExcelSheetTabColor'                = 'Set-OfficeExcelSheetTabColor'
    'ExcelSheetView'                    = 'Set-OfficeExcelWorksheetView'
    'ExcelSheetVisibility'              = 'Set-OfficeExcelSheetVisibility'
    'ExcelSlicer'                       = 'Add-OfficeExcelSlicer'
    'ExcelSort'                         = 'Invoke-OfficeExcelSort'
    'ExcelSparkline'                    = 'Add-OfficeExcelSparkline'
    'ExcelStreamingContract'            = 'Get-OfficeExcelStreamingContract'
    'ExcelSubtotals'                    = 'Add-OfficeExcelSubtotalSummary'
    'ExcelSubtotalSummary'              = 'Add-OfficeExcelSubtotalSummary'
    'ExcelSummary'                      = 'Get-OfficeExcelSummary'
    'ExcelTable'                        = 'Add-OfficeExcelTable'
    'ExcelTableOfContents'              = 'Add-OfficeExcelTableOfContents'
    'ExcelTableStyle'                   = 'Get-OfficeExcelTableStyle'
    'ExcelTemplate'                     = 'Invoke-OfficeExcelTemplate'
    'ExcelTemplateApply'                = 'Invoke-OfficeExcelTemplate'
    'ExcelTemplateBinding'              = 'Test-OfficeExcelTemplateBinding'
    'ExcelTemplateMarkers'              = 'Get-OfficeExcelTemplateMarker'
    'ExcelTemplateOptionalRow'          = 'Invoke-OfficeExcelTemplateOptionalRow'
    'ExcelTemplateOptionalRows'         = 'Invoke-OfficeExcelTemplateOptionalRow'
    'ExcelTemplateRow'                  = 'Invoke-OfficeExcelTemplateRow'
    'ExcelTemplateRows'                 = 'Invoke-OfficeExcelTemplateRow'
    'ExcelTemplateSheet'                = 'Invoke-OfficeExcelTemplateSheet'
    'ExcelTemplateSheets'               = 'Invoke-OfficeExcelTemplateSheet'
    'ExcelTemplateValidate'             = 'Test-OfficeExcelTemplateBinding'
    'ExcelTextRun'                      = 'New-OfficeTextRun'
    'ExcelTheme'                        = 'Set-OfficeExcelTheme'
    'ExcelThreadedComment'              = 'Add-OfficeExcelThreadedComment'
    'ExcelTimeline'                     = 'Add-OfficeExcelTimeline'
    'ExcelUnprotect'                    = 'Unprotect-OfficeExcelSheet'
    'ExcelUrlLinks'                     = 'Set-OfficeExcelUrlLinks'
    'ExcelUrlLinksByHeader'             = 'Set-OfficeExcelUrlLinksByHeader'
    'ExcelValidationCustomFormula'      = 'Add-OfficeExcelValidationCustomFormula'
    'ExcelValidationDate'               = 'Add-OfficeExcelValidationDate'
    'ExcelValidationDecimal'            = 'Add-OfficeExcelValidationDecimal'
    'ExcelValidationList'               = 'Add-OfficeExcelValidationList'
    'ExcelValidationTextLength'         = 'Add-OfficeExcelValidationTextLength'
    'ExcelValidationTime'               = 'Add-OfficeExcelValidationTime'
    'ExcelValidationWholeNumber'        = 'Add-OfficeExcelValidationWholeNumber'
    'ExcelVisual'                       = 'Add-OfficeExcelVisual'
    'ExcelWorkbookCompare'              = 'Compare-OfficeExcelWorkbook'
    'ExcelWorkbookCopy'                 = 'Copy-OfficeExcelWorkbook'
    'ExcelWorkbookDoctor'               = 'Test-OfficeExcelWorkbook'
    'ExcelWorkbookJoin'                 = 'Join-OfficeExcelWorkbook'
    'ExcelWorkbookMerge'                = 'Join-OfficeExcelWorkbook'
    'ExcelWorkbookProtect'              = 'Protect-OfficeExcelWorkbook'
    'ExcelWorkbookRepair'               = 'Repair-OfficeExcelWorkbook'
    'ExcelWorkbookUnprotect'            = 'Unprotect-OfficeExcelWorkbook'
    'ExcelWorksheetView'                = 'Get-OfficeExcelWorksheetView'
    'ExcelWriteReservation'             = 'Get-OfficeExcelWriteReservation'
    'ExcelWriteReservationClear'        = 'Clear-OfficeExcelWriteReservation'
    'ExcelWriteReservationSet'          = 'Set-OfficeExcelWriteReservation'
    'Export-OfficeDocumentAsset'        = 'Get-OfficeDocumentAsset'
    'Export-VisioStencilPreviewGallery' = 'Export-OfficeVisioStencilPreviewGallery'
    'Find-VisioStencil'                 = 'Find-OfficeVisioStencil'
    'Get-OfficeReaderCapability'        = 'Get-OfficeDocumentCapability'
    'Import-VisioStencil'               = 'Import-OfficeVisioStencil'
    'MarkdownCallout'                   = 'Add-OfficeMarkdownCallout'
    'MarkdownCode'                      = 'Add-OfficeMarkdownCode'
    'MarkdownDefinitionList'            = 'Add-OfficeMarkdownDefinitionList'
    'MarkdownDetails'                   = 'Add-OfficeMarkdownDetails'
    'MarkdownFrontMatter'               = 'Add-OfficeMarkdownFrontMatter'
    'MarkdownHeading'                   = 'Add-OfficeMarkdownHeading'
    'MarkdownHorizontalRule'            = 'Add-OfficeMarkdownHorizontalRule'
    'MarkdownHr'                        = 'Add-OfficeMarkdownHorizontalRule'
    'MarkdownImage'                     = 'Add-OfficeMarkdownImage'
    'MarkdownList'                      = 'Add-OfficeMarkdownList'
    'MarkdownNew'                       = 'New-OfficeMarkdown'
    'MarkdownParagraph'                 = 'Add-OfficeMarkdownParagraph'
    'MarkdownQuote'                     = 'Add-OfficeMarkdownQuote'
    'MarkdownTable'                     = 'Add-OfficeMarkdownTable'
    'MarkdownTableOfContents'           = 'Add-OfficeMarkdownTableOfContents'
    'MarkdownTaskList'                  = 'Add-OfficeMarkdownTaskList'
    'MarkdownToc'                       = 'Add-OfficeMarkdownTableOfContents'
    'Merge-OfficeExcelSheet'            = 'Join-OfficeExcelSheet'
    'Merge-OfficeExcelWorkbook'         = 'Join-OfficeExcelWorkbook'
    'Merge-OfficeWordDocument'          = 'Join-OfficeWordDocument'
    'New-VisioGallery'                  = 'New-OfficeVisioGallery'
    'OfficeExcel'                       = 'New-OfficeExcel'
    'OfficeMarkdown'                    = 'New-OfficeMarkdown'
    'OfficeOpenDocument'                = 'New-OfficeOpenDocument'
    'OfficePdf'                         = 'New-OfficePdf'
    'OfficePowerPoint'                  = 'New-OfficePowerPoint'
    'OfficeRtf'                         = 'New-OfficeRtf'
    'OfficeVisio'                       = 'New-OfficeVisio'
    'OfficeVisual'                      = 'ConvertTo-OfficeVisual'
    'OfficeWord'                        = 'New-OfficeWord'
    'OpenDocumentNew'                   = 'New-OfficeOpenDocument'
    'PdfAttachment'                     = 'Add-OfficePdfAttachment'
    'PdfBackground'                     = 'Set-OfficePdfBackground'
    'PdfBackgroundImage'                = 'Set-OfficePdfBackgroundImage'
    'PdfBackgroundShape'                = 'Add-OfficePdfBackgroundShape'
    'PdfBookmark'                       = 'Add-OfficePdfBookmark'
    'PdfCanvasStamp'                    = 'Add-OfficePdfCanvas'
    'PdfCanvasText'                     = 'Add-OfficePdfCanvasText'
    'PdfCompliance'                     = 'Set-OfficePdfCompliance'
    'PdfElectronicInvoice'              = 'Set-OfficePdfElectronicInvoice'
    'PdfFooter'                         = 'Set-OfficePdfFooter'
    'PdfFormField'                      = 'Add-OfficePdfFormField'
    'PdfHeader'                         = 'Set-OfficePdfHeader'
    'PdfHeading'                        = 'Add-OfficePdfHeading'
    'PdfHorizontalRule'                 = 'Add-OfficePdfHorizontalRule'
    'PdfHr'                             = 'Add-OfficePdfHorizontalRule'
    'PdfImage'                          = 'Add-OfficePdfImage'
    'PdfList'                           = 'Add-OfficePdfList'
    'PdfMetadata'                       = 'Set-OfficePdfMetadata'
    'PdfNativeTextRun'                  = 'ConvertTo-OfficePdfTextRun'
    'PdfNew'                            = 'New-OfficePdf'
    'PdfPageBorder'                     = 'Set-OfficePdfPageBorder'
    'PdfPageBreak'                      = 'Add-OfficePdfPageBreak'
    'PdfPageOverlay'                    = 'Add-OfficePdfPageOverlay'
    'PdfPageSetup'                      = 'Set-OfficePdfPageSetup'
    'PdfPanel'                          = 'Add-OfficePdfPanel'
    'PdfParagraph'                      = 'Add-OfficePdfParagraph'
    'PdfRow'                            = 'Add-OfficePdfRow'
    'PdfSpace'                          = 'Add-OfficePdfSpacer'
    'PdfSpacer'                         = 'Add-OfficePdfSpacer'
    'PdfStamp'                          = 'Add-OfficePdfStamp'
    'PdfTable'                          = 'Add-OfficePdfTable'
    'PdfTableCell'                      = 'New-OfficePdfTableCell'
    'PdfTableCellCheckBox'              = 'New-OfficePdfTableCellCheckBox'
    'PdfTableCellField'                 = 'New-OfficePdfTableCellField'
    'PdfTableCellImage'                 = 'New-OfficePdfTableCellImage'
    'PdfText'                           = 'Add-OfficePdfText'
    'PdfTextRun'                        = 'New-OfficeTextRun'
    'PdfTheme'                          = 'Set-OfficePdfTheme'
    'PdfVisual'                         = 'Add-OfficePdfVisual'
    'PdfWatermark'                      = 'Add-OfficePdfWatermark'
    'PowerPointNew'                     = 'New-OfficePowerPoint'
    'PowerPointTextRun'                 = 'New-OfficeTextRun'
    'PptArrange'                        = 'Set-OfficePowerPointShapeLayout'
    'PptBackground'                     = 'Set-OfficePowerPointBackground'
    'PptBullets'                        = 'Add-OfficePowerPointBullets'
    'PptChart'                          = 'Add-OfficePowerPointChart'
    'PptDeckPlan'                       = 'New-OfficePowerPointDeckPlan'
    'PptDesignerDeck'                   = 'Add-OfficePowerPointDesignerDeck'
    'PptImage'                          = 'Add-OfficePowerPointImage'
    'PptLayoutBox'                      = 'Get-OfficePowerPointLayoutBox'
    'PptLayoutPlaceholderBounds'        = 'Set-OfficePowerPointLayoutPlaceholderBounds'
    'PptLayoutPlaceholderMargins'       = 'Set-OfficePowerPointLayoutPlaceholderTextMargins'
    'PptLayoutPlaceholders'             = 'Get-OfficePowerPointLayoutPlaceholder'
    'PptLayoutPlaceholderTextStyle'     = 'Set-OfficePowerPointLayoutPlaceholderTextStyle'
    'PptNew'                            = 'New-OfficePowerPoint'
    'PptNotes'                          = 'Set-OfficePowerPointNotes'
    'PptPlaceholderText'                = 'Set-OfficePowerPointPlaceholderText'
    'PptPlanCapability'                 = 'Add-OfficePowerPointPlanCapability'
    'PptPlanCardGrid'                   = 'Add-OfficePowerPointPlanCardGrid'
    'PptPlanCaseStudy'                  = 'Add-OfficePowerPointPlanCaseStudy'
    'PptPlanCoverage'                   = 'Add-OfficePowerPointPlanCoverage'
    'PptPlanLogoWall'                   = 'Add-OfficePowerPointPlanLogoWall'
    'PptPlanProcess'                    = 'Add-OfficePowerPointPlanProcess'
    'PptPlanSection'                    = 'Add-OfficePowerPointPlanSection'
    'PptSection'                        = 'Add-OfficePowerPointSection'
    'PptShape'                          = 'Add-OfficePowerPointShape'
    'PptShapeLayout'                    = 'Set-OfficePowerPointShapeLayout'
    'PptSlide'                          = 'Add-OfficePowerPointSlide'
    'PptSlideLayout'                    = 'Set-OfficePowerPointSlideLayout'
    'PptSlideSize'                      = 'Set-OfficePowerPointSlideSize'
    'PptTable'                          = 'Add-OfficePowerPointTable'
    'PptTextBox'                        = 'Add-OfficePowerPointTextBox'
    'PptTextRun'                        = 'New-OfficeTextRun'
    'PptTheme'                          = 'Get-OfficePowerPointTheme'
    'PptThemeColor'                     = 'Set-OfficePowerPointThemeColor'
    'PptThemeFonts'                     = 'Set-OfficePowerPointThemeFonts'
    'PptThemeName'                      = 'Set-OfficePowerPointThemeName'
    'PptTitle'                          = 'Set-OfficePowerPointSlideTitle'
    'PptTransition'                     = 'Set-OfficePowerPointSlideTransition'
    'PptVisual'                         = 'Add-OfficePowerPointVisual'
    'Read-OfficeDocument'               = 'Get-OfficeDocument'
    'Read-OfficeDocumentAsset'          = 'Get-OfficeDocumentAsset'
    'Read-OfficeDocumentChunk'          = 'Get-OfficeDocumentChunk'
    'Read-OfficeDocumentTable'          = 'Get-OfficeDocumentTable'
    'Read-OfficeDocumentVisual'         = 'Get-OfficeDocumentVisual'
    'Replace-OfficeExcelText'           = 'Update-OfficeExcelText'
    'Replace-OfficePowerPointText'      = 'Update-OfficePowerPointText'
    'Replace-OfficeRtfText'             = 'Update-OfficeRtfText'
    'Replace-OfficeWordText'            = 'Update-OfficeWordText'
    'RtfNew'                            = 'New-OfficeRtf'
    'RtfOpen'                           = 'Get-OfficeRtf'
    'RtfText'                           = 'Update-OfficeRtfText'
    'Set-OfficeExcelSheetOrder'         = 'Move-OfficeExcelSheet'
    'TextRun'                           = 'New-OfficeTextRun'
    'VisioArrange'                      = 'Set-OfficeVisioShapeLayout'
    'VisioConnector'                    = 'Add-OfficeVisioConnector'
    'VisioContainer'                    = 'Add-OfficeVisioContainer'
    'VisioDiamond'                      = 'Add-OfficeVisioDiamond'
    'VisioEllipse'                      = 'Add-OfficeVisioEllipse'
    'VisioInfo'                         = 'Get-OfficeVisioInfo'
    'VisioLayout'                       = 'Set-OfficeVisioShapeLayout'
    'VisioNew'                          = 'New-OfficeVisio'
    'VisioOpen'                         = 'Get-OfficeVisio'
    'VisioPage'                         = 'Add-OfficeVisioPage'
    'VisioRect'                         = 'Add-OfficeVisioRectangle'
    'VisioRectangle'                    = 'Add-OfficeVisioRectangle'
    'VisioSave'                         = 'Save-OfficeVisio'
    'VisioStencil'                      = 'Add-OfficeVisioStencilShape'
    'VisioStencilCatalog'               = 'Get-OfficeVisioStencilCatalog'
    'VisioStencilImport'                = 'Import-OfficeVisioStencil'
    'VisioText'                         = 'Add-OfficeVisioTextBox'
    'VisioTextBox'                      = 'Add-OfficeVisioTextBox'
    'WordBold'                          = 'Add-OfficeWordText'
    'WordBookmark'                      = 'Add-OfficeWordBookmark'
    'WordBreak'                         = 'Add-OfficeWordBreak'
    'WordChart'                         = 'Add-OfficeWordChart'
    'WordCheckBox'                      = 'Add-OfficeWordCheckBox'
    'WordCheckBoxes'                    = 'Get-OfficeWordCheckBox'
    'WordComboBox'                      = 'Add-OfficeWordComboBox'
    'WordComboBoxes'                    = 'Get-OfficeWordComboBox'
    'WordContentControl'                = 'Add-OfficeWordContentControl'
    'WordContentControls'               = 'Get-OfficeWordContentControl'
    'WordCoverPage'                     = 'Add-OfficeWordCoverPage'
    'WordDatePicker'                    = 'Add-OfficeWordDatePicker'
    'WordDatePickers'                   = 'Get-OfficeWordDatePicker'
    'WordDocumentJoin'                  = 'Join-OfficeWordDocument'
    'WordDropDownList'                  = 'Add-OfficeWordDropDownList'
    'WordDropDownLists'                 = 'Get-OfficeWordDropDownList'
    'WordEndnote'                       = 'Add-OfficeWordEndnote'
    'WordEndnotes'                      = 'Get-OfficeWordEndnote'
    'WordEquation'                      = 'Add-OfficeWordEquation'
    'WordField'                         = 'Add-OfficeWordField'
    'WordFooter'                        = 'Add-OfficeWordFooter'
    'WordFootnote'                      = 'Add-OfficeWordFootnote'
    'WordFootnotes'                     = 'Get-OfficeWordFootnote'
    'WordHeader'                        = 'Add-OfficeWordHeader'
    'WordHyperlink'                     = 'Add-OfficeWordHyperlink'
    'WordImage'                         = 'Add-OfficeWordImage'
    'WordImages'                        = 'Get-OfficeWordImage'
    'WordImageStyle'                    = 'Set-OfficeWordImage'
    'WordItalic'                        = 'Add-OfficeWordText'
    'WordList'                          = 'Add-OfficeWordList'
    'WordListItem'                      = 'Add-OfficeWordListItem'
    'WordNew'                           = 'New-OfficeWord'
    'WordPageNumber'                    = 'Add-OfficeWordPageNumber'
    'WordPageSetup'                     = 'Set-OfficeWordPageSetup'
    'WordParagraph'                     = 'Add-OfficeWordParagraph'
    'WordParagraphStyle'                = 'Set-OfficeWordParagraphStyle'
    'WordPictureControl'                = 'Add-OfficeWordPictureControl'
    'WordPictureControls'               = 'Get-OfficeWordPictureControl'
    'WordRepeatingSection'              = 'Add-OfficeWordRepeatingSection'
    'WordRepeatingSections'             = 'Get-OfficeWordRepeatingSection'
    'WordSection'                       = 'Add-OfficeWordSection'
    'WordShape'                         = 'Add-OfficeWordShape'
    'WordShapes'                        = 'Get-OfficeWordShape'
    'WordShapeStyle'                    = 'Set-OfficeWordShape'
    'WordStatistics'                    = 'Get-OfficeWordStatistics'
    'WordTable'                         = 'Add-OfficeWordTable'
    'WordTableCell'                     = 'Add-OfficeWordTableCell'
    'WordTableCells'                    = 'Get-OfficeWordTableCell'
    'WordTableCellSpec'                 = 'New-OfficeWordTableCell'
    'WordTableCellStyle'                = 'Set-OfficeWordTableCell'
    'WordTableCondition'                = 'Add-OfficeWordTableCondition'
    'WordTableOfContents'               = 'Add-OfficeWordTableOfContents'
    'WordTabStop'                       = 'Add-OfficeWordTabStop'
    'WordText'                          = 'Add-OfficeWordText'
    'WordTextBox'                       = 'Add-OfficeWordTextBox'
    'WordTextRun'                       = 'New-OfficeTextRun'
    'WordTextStyle'                     = 'Set-OfficeWordTextStyle'
    'WordVisual'                        = 'Add-OfficeWordVisual'
    'WordWatermark'                     = 'Add-OfficeWordWatermark'
}

foreach ($binaryAlias in $binaryAliases.GetEnumerator()) {
    Set-Alias -Name $binaryAlias.Key -Value $binaryAlias.Value -Scope Local -Force -ErrorAction Stop
}

Export-ModuleMember -Alias '*' -Cmdlet '*'
