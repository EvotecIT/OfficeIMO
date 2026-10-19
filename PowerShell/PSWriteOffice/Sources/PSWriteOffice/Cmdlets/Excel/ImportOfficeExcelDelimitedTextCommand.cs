#pragma warning disable CS1591
using System;
using System.Globalization;
using System.IO;
using System.Management.Automation;
using OfficeIMO.CSV;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Csv;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Imports normalized CSV/TSV text into an Excel workbook through OfficeIMO.</summary>
/// <example>
///   <summary>Import a semicolon-delimited regional export as a typed worksheet table.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$result = Import-OfficeExcelDelimitedText -Path .\Report.xlsx `
///     -SourcePath .\sales-pl.csv `
///     -Delimiter ';' `
///     -CultureName 'pl-PL' `
///     -SheetName Sales `
///     -PassThru
/// $result | Format-List SheetName,Range,RowCount,ColumnCount,Delimiter</code>
///   <para>Normalizes delimited text through OfficeIMO, performs culture-aware value conversion, and writes the result as an Excel table unless -NoTable is used.</para>
/// </example>
[Cmdlet(VerbsData.Import, "OfficeExcelDelimitedText", DefaultParameterSetName = ParameterSetPath, SupportsShouldProcess = true, ConfirmImpact = ConfirmImpact.Medium)]
[Alias("ExcelDelimitedImport", "ExcelCsvImport")]
[OutputType(typeof(PSObject))]
public sealed class ImportOfficeExcelDelimitedTextCommand : PSCmdlet {
    private const string ParameterSetPath = "Path";
    private const string ParameterSetDocument = "Document";

    /// <summary>Workbook path.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Workbook document.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Delimited text source path.</summary>
    [Parameter(Mandatory = true)]
    public string SourcePath { get; set; } = string.Empty;

    /// <summary>Delimiter character. When omitted, it is detected.</summary>
    [Parameter]
    public char? Delimiter { get; set; }
    /// <summary>Worksheet name to create or inspect.</summary>
    [Parameter]
    public string? SheetName { get; set; }
    /// <summary>Culture name for number and date conversion.</summary>
    [Parameter]
    public string? CultureName { get; set; }
    /// <summary>Treat the first row as data.</summary>
    [Parameter]
    public SwitchParameter NoHeader { get; set; }
    /// <summary>Number of parsed delimited records to skip before header discovery or data import.</summary>
    [Parameter]
    [ValidateRange(0, int.MaxValue)]
    public int SkipRows { get; set; }
    /// <summary>Import rows without creating an Excel table.</summary>
    [Parameter]
    public SwitchParameter NoTable { get; set; }
    /// <summary>Keep imported values as text.</summary>
    [Parameter]
    public SwitchParameter NoTypeConversion { get; set; }
    /// <summary>Emit a result object.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    protected override void ProcessRecord() {
        var source = SessionState.Path.GetUnresolvedProviderPathFromPSPath(SourcePath);
        if (!File.Exists(source)) {
            throw new FileNotFoundException($"Delimited text file '{source}' was not found.", source);
        }

        var target = ResolveTargetPath();
        if (!ShouldProcess(target ?? "Excel document", "Import delimited text into Excel workbook")) {
            return;
        }

        using var workbook = ResolveWorkbook(target);
        var culture = string.IsNullOrWhiteSpace(CultureName) ? CultureInfo.InvariantCulture : CultureInfo.GetCultureInfo(CultureName!);
        var loadOptions = new CsvLoadOptions {
            DetectDelimiter = !Delimiter.HasValue,
            HasHeaderRow = !NoHeader.IsPresent,
            SkipInitialRecords = SkipRows,
            Culture = culture
        };
        if (Delimiter.HasValue) {
            loadOptions.Delimiter = Delimiter.Value;
        }

        var result = workbook.Document.ImportCsvFile(source, new ExcelCsvImportOptions {
            SheetName = string.IsNullOrWhiteSpace(SheetName) ? "Import" : SheetName!,
            IncludeHeaders = !NoHeader.IsPresent,
            CreateTable = !NoTable.IsPresent,
            LoadOptions = loadOptions,
            ReaderOptions = new CsvDataReaderOptions {
                InferSchema = !NoTypeConversion.IsPresent
            }
        });

        workbook.SaveIfOwned();
        if (PassThru.IsPresent) {
            var output = new PSObject();
            output.Properties.Add(new PSNoteProperty("SheetName", result.SheetName));
            var rowCount = 0;
            var columnCount = 0;
            if (!string.IsNullOrWhiteSpace(result.Range)) {
                var (firstRow, firstColumn, lastRow, lastColumn) = A1.ParseRange(result.Range);
                rowCount = Math.Max(0, lastRow - firstRow + (NoHeader.IsPresent ? 1 : 0));
                columnCount = lastColumn - firstColumn + 1;
            }
            output.Properties.Add(new PSNoteProperty("TableName", result.TableName));
            output.Properties.Add(new PSNoteProperty("Range", result.Range));
            output.Properties.Add(new PSNoteProperty("RowCount", rowCount));
            output.Properties.Add(new PSNoteProperty("ColumnCount", columnCount));
            output.Properties.Add(new PSNoteProperty("Delimiter", result.Delimiter.ToString()));
            output.Properties.Add(new PSNoteProperty("Warnings", Array.Empty<string>()));
            WriteObject(output);
        }
    }

    private string? ResolveTargetPath() {
        if (!string.Equals(ParameterSetName, ParameterSetPath, System.StringComparison.OrdinalIgnoreCase)) {
            return Document.FilePath;
        }

        return SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
    }

    private ExcelWorkbookCommandScope ResolveWorkbook(string? targetPath) {
        if (!string.Equals(ParameterSetName, ParameterSetPath, System.StringComparison.OrdinalIgnoreCase)) {
            return new ExcelWorkbookCommandScope(Document, ownsDocument: false);
        }

        if (string.IsNullOrWhiteSpace(targetPath)) {
            throw new PSArgumentException("Specify a workbook path.", nameof(Path));
        }

        var resolvedPath = targetPath!;
        var directory = System.IO.Path.GetDirectoryName(resolvedPath);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory)) {
            Directory.CreateDirectory(directory);
        }

        var document = File.Exists(resolvedPath)
            ? ExcelDocumentService.LoadDocument(resolvedPath, readOnly: false, autoSave: false)
            : ExcelDocumentService.CreateDocument(resolvedPath, autoSave: false);

        return new ExcelWorkbookCommandScope(document, ownsDocument: true);
    }
}
#pragma warning restore CS1591