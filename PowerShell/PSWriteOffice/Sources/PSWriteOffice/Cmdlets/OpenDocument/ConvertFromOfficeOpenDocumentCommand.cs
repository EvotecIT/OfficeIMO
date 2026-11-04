using System;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Excel.OpenDocument;
using OfficeIMO.OpenDocument;
using OfficeIMO.PowerPoint.OpenDocument;
using OfficeIMO.Word.OpenDocument;

namespace PSWriteOffice.Cmdlets.OpenDocument;

/// <summary>Converts native ODT, ODS, or ODP content to Word, Excel, or PowerPoint with fidelity evidence.</summary>
/// <example>
///   <summary>Convert an ODS spreadsheet to Excel.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$options = New-OfficeExcelOpenDocumentOptions -IncludeBasicStyles -MaximumExpandedCells 250000
/// ConvertFrom-OfficeOpenDocument -Path .\Status.ods -OutputPath .\Status.xlsx -ExcelOptions $options</code>
/// </example>
[Cmdlet(VerbsData.ConvertFrom, "OfficeOpenDocument", SupportsShouldProcess = true)]
public sealed class ConvertFromOfficeOpenDocumentCommand : PSCmdlet
{
    /// <summary>Path to an ODT, ODS, or ODP file.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Path { get; set; } = string.Empty;

    /// <summary>Destination DOCX, XLSX, or PPTX path.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string OutputPath { get; set; } = string.Empty;

    /// <summary>Optional ODT-to-Word conversion settings.</summary>
    [Parameter]
    public WordOpenDocumentConversionOptions? WordOptions { get; set; }

    /// <summary>Optional ODS-to-Excel conversion settings.</summary>
    [Parameter]
    public ExcelOpenDocumentConversionOptions? ExcelOptions { get; set; }

    /// <summary>Optional ODP-to-PowerPoint conversion settings.</summary>
    [Parameter]
    public PowerPointOpenDocumentConversionOptions? PowerPointOptions { get; set; }

    /// <summary>Throw when the conversion approximates, skips, or cannot map a feature.</summary>
    [Parameter]
    public SwitchParameter FailOnLoss { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var input = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
        var output = SessionState.Path.GetUnresolvedProviderPathFromPSPath(OutputPath);
        var source = OdfDocument.Load(input);
        OpenDocumentCommandUtilities.ValidateOfficeExtension(output, source.Kind, nameof(OutputPath));
        if (!ShouldProcess(output, "Convert OpenDocument to Office document")) return;
        Directory.CreateDirectory(System.IO.Path.GetDirectoryName(output) ?? SessionState.Path.CurrentFileSystemLocation.Path);
        switch (source)
        {
            case OdtDocument text:
                OfficeConversionReport wordResult;
                using (var converted = text.ToWord(WordOptions, out wordResult))
                {
                    if (FailOnLoss.IsPresent) wordResult.ThrowIfLoss();
                    converted.Save(output);
                }
                WriteObject(wordResult);
                break;
            case OdsDocument spreadsheet:
                OfficeConversionReport excelResult;
                using (var converted = spreadsheet.ToExcel(ExcelOptions, out excelResult))
                {
                    if (FailOnLoss.IsPresent) excelResult.ThrowIfLoss();
                    converted.Save(output);
                }
                WriteObject(excelResult);
                break;
            case OdpPresentation presentation:
                OfficeConversionReport presentationResult;
                using (var converted = presentation.ToPowerPoint(PowerPointOptions, out presentationResult))
                {
                    if (FailOnLoss.IsPresent) presentationResult.ThrowIfLoss();
                    converted.Save(output);
                }
                WriteObject(presentationResult);
                break;
            default:
                throw new InvalidOperationException("Unsupported OpenDocument kind.");
        }
    }

}
