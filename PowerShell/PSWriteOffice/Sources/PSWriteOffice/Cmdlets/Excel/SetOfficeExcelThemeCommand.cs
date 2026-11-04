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

/// <summary>Applies an OfficeTheme or theme XML to a workbook.</summary>
/// <example>
/// <summary>Apply a built-in theme</summary>
/// <code>Open-OfficeExcel -Path .\Reports\Q3-Sales.xlsx {
///     ExcelTheme Minimal
/// }</code>
/// <para>Restyles the workbook's cell styles, table styles and charts with the Minimal theme.</para>
/// </example>
/// <example>
/// <summary>Apply the corporate theme from a template</summary>
/// <code>Open-OfficeExcel -Path .\Reports\Legacy-Report.xlsx {
///     ExcelTheme .\Templates\Northwind.xltx
/// }</code>
/// <para>Takes the colours and fonts from the company template so an older workbook matches current branding.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelTheme", DefaultParameterSetName = "Name")]
[Alias("ExcelTheme")]
public sealed class SetOfficeExcelThemeCommand : PSCmdlet {
    /// <summary>Caller-owned workbook; omission uses the enclosing Excel content block.</summary>
    [Parameter(ValueFromPipeline = true)]
    public ExcelDocument? Document { get; set; }

    /// <summary>Built-in Office theme, native theme object or a document whose visual theme is read.</summary>
    [Parameter(ParameterSetName = "Theme", Mandatory = true, Position = 0)]
    [OfficeThemeTransformation]
    public OfficeThemeInput Theme { get; set; } = null!;

    /// <summary>Optional workbook theme name to apply after writing the theme.</summary>
    [Parameter(ParameterSetName = "Theme")]
    [Parameter(ParameterSetName = "Xml")]
    [Parameter(ParameterSetName = "XmlPath")]
    [Parameter(ParameterSetName = "Name", Mandatory = true)]
    public string? Name { get; set; }

    /// <summary>Emits the native result after the operation.</summary>
    [Parameter()]
    public SwitchParameter PassThru { get; set; }

    /// <summary>Theme XML to write to the workbook theme part.</summary>
    [Parameter(ParameterSetName = "Xml", Mandatory = true)]
    public string Xml { get; set; } = null!;

    /// <summary>Literal path to an XML definition.</summary>
    [Parameter(ParameterSetName = "XmlPath", Mandatory = true)]
    public string XmlPath { get; set; } = null!;

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var document = ExcelCommandTargets.Document(Document);
        if (ParameterSetName == "Theme") document.ApplyTheme(Theme!.Resolve(OfficeTheme.FromDocument));
        else if (ParameterSetName != "Name") document.SetWorkbookThemeXml(Xml ?? File.ReadAllText(OfficePath.ResolveInput(SessionState, XmlPath!, literal: true)[0]));
        if (Name != null) document.SetWorkbookThemeName(Name);
        if (PassThru) WriteObject(document.GetWorkbookTheme(includeXml: true));
    }

}
