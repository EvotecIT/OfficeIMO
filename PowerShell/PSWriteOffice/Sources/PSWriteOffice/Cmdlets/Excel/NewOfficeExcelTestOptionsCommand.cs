using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Security;
using System.Threading.Tasks;
using OfficeIMO;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Creates Excel checks for Test-OfficeDocument: capability preflight and package validation.</summary>
/// <example>
/// <summary>Check that formulas can be calculated</summary>
/// <code>Test-OfficeDocument -Path .\Finance\Budget-2027.xlsx -Options (New-OfficeExcelTestOptions -Capability CalculateFormulas)</code>
/// <para>Tests whether OfficeIMO can calculate the formulas in a workbook.</para>
/// </example>
/// <example>
/// <summary>Preflight templates before a report run</summary>
/// <code>$checks = New-OfficeExcelTestOptions -Capability BindTemplate, ExportPdfReport -IncludeRepairHints
/// Get-ChildItem -Path .\Templates -Filter *.xltx |
///     Test-OfficeDocument -Options $checks</code>
/// <para>Checks each template for template binding and PDF export and includes repair hints for anything that fails.</para>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeExcelTestOptions", DefaultParameterSetName = "Default")]
public sealed class NewOfficeExcelTestOptionsCommand : PSCmdlet {
    /// <summary>Capabilities to evaluate or enforce. Defaults to all capabilities.</summary>
    [Parameter(ParameterSetName = "Default")]
    public ExcelPreflightCapability[]? Capability { get; set; }

    /// <summary>Skips native Open XML schema validation while retaining other workbook checks.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter NoOpenXmlValidation { get; set; }

    /// <summary>Include discovered feature rows in object output.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter IncludeFeatures { get; set; }

    /// <summary>Include actionable OfficeIMO repair hints for blocked capabilities.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter IncludeRepairHints { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() => WriteObject(new ExcelTestOptions { Doctor = new ExcelWorkbookDoctorOptions { ValidateOpenXml = !NoOpenXmlValidation }, Capability = Capability, IncludeFeatures = IncludeFeatures, IncludeRepairHints = IncludeRepairHints });

}
