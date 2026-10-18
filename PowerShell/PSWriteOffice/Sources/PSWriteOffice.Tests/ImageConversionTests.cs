using System.Management.Automation;
using System.Management.Automation.Runspaces;
using OfficeIMO.Drawing;
using OfficeIMO.CSV;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Ocr;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using PSWriteOffice.Cmdlets.Ocr;
using PSWriteOffice.Services.Ocr;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class ImageConversionTests
{
    [Fact]
    public async Task ReviewedImageExportsTypedCellsAndReusesRecognition()
    {
        string directory = NewDirectory();
        try
        {
            int calls = 0;
            var engine = new DelegateOcrEngine("table-contract", (_, _) => { calls++; return Task.FromResult(TableResult()); });
            var review = await Source().PrepareSearchableOcrAsync(engine);
            var options = new PdfTablesToExcelOptions { UseFirstRowAsHeader = true };
            string output = Path.Combine(directory, "ledger.xlsx");
            var result = OfficeImageExporter.Export(review, output, false, options, null, default);
            Assert.Same(review, result.Review);
            var report = Assert.IsType<PdfExcelTableImportReport>(result.ConversionReport);
            Assert.True(Assert.Single(report.Entries).FirstRowUsedAsHeader);
            using (var excel = ExcelDocument.Load(output))
            {
                var sheet = Assert.Single(excel.Sheets);
                Assert.Equal("Credit", sheet.CellAt(3, 1).GetValue<string>());
                Assert.Equal(-13.5, sheet.CellAt(3, 2).GetValue<double>());
            }
            OfficeImageExporter.Export(review, Path.Combine(directory, "ledger.csv"), false, options, null, default);
            Assert.Contains("Credit,'-13.50", File.ReadAllText(Path.Combine(directory, "ledger.csv")));
            var csvOptions = new OfficeIMO.CSV.CsvSaveOptions { FormulaInjectionPolicy = OfficeIMO.CSV.CsvFormulaInjectionPolicy.Preserve };
            OfficeImageExporter.Export(review, Path.Combine(directory, "exact.csv"), false, options, null, default, csvOptions: csvOptions);
            Assert.Contains("Credit,-13.50", File.ReadAllText(Path.Combine(directory, "exact.csv")));
            Assert.True(csvOptions.IncludeHeader);
            OfficeImageExporter.Export(review, Path.Combine(directory, "ledger.json"), false, options, null, default);
            Assert.Equal(1, calls);
        }
        finally { Directory.Delete(directory, recursive: true); }
    }

    [Fact]
    public async Task CsvDefaultsEscapeAndCompleteCallerOptionsRetainTheirPolicy()
    {
        string directory = NewDirectory();
        try
        {
            var review = await Source().PrepareSearchableOcrAsync(new DelegateOcrEngine("formula-table", (_, _) => Task.FromResult(new OcrResult
            {
                Spans = new[] {
                    Word("Value", 20, 20, 100), Word("Amount", 200, 20, 45),
                    Word("=SUM(1+1)", 20, 40, 100), Word("1", 200, 40, 40),
                    Word("+cmd", 20, 60, 100), Word("2", 200, 60, 40),
                    Word("-13.50", 20, 80, 100), Word("3", 200, 80, 40),
                    Word("@name", 20, 100, 100), Word("4", 200, 100, 40)
                }
            })));
            var cases = new (string Name, CsvSaveOptions? Options, string Delimiter, bool Escape)[] {
                ("default", null, ",", true),
                ("custom-default", new CsvSaveOptions { Delimiter = ';' }, ";", false),
                ("custom-escape", new CsvSaveOptions { Delimiter = ';', FormulaInjectionPolicy = CsvFormulaInjectionPolicy.Escape }, ";", true),
                ("custom-preserve", new CsvSaveOptions { Delimiter = ';', FormulaInjectionPolicy = CsvFormulaInjectionPolicy.Preserve }, ";", false)
            };
            foreach (var item in cases)
            {
                string path = Path.Combine(directory, item.Name + ".csv");
                OfficeImageExporter.Export(review, path, false, new(), null, default, csvOptions: item.Options);
                string[] lines = File.ReadAllLines(path);
                Assert.Equal(new[] { "Value" + item.Delimiter + "Amount" }
                    .Concat(new[] { "=SUM(1+1)", "+cmd", "-13.50", "@name" }
                        .Select((value, index) => (item.Escape ? "'" : "") + value + item.Delimiter + (index + 1))), lines);
                if (item.Options != null)
                {
                    Assert.True(item.Options.IncludeHeader);
                    Assert.Equal(item.Escape ? CsvFormulaInjectionPolicy.Escape : CsvFormulaInjectionPolicy.Preserve, item.Options.FormulaInjectionPolicy);
                }
            }
        }
        finally { Directory.Delete(directory, recursive: true); }
    }

    [Fact]
    public async Task OutputFailuresAndCancellationPreserveExistingDestination()
    {
        string directory = NewDirectory();
        try
        {
            string output = Path.Combine(directory, "existing.xlsx");
            File.WriteAllText(output, "valuable destination");
            var review = await Source().PrepareSearchableOcrAsync(new DelegateOcrEngine("table", (_, _) => Task.FromResult(TableResult())));
            Assert.Throws<IOException>(() => OfficeImageExporter.Export(review, output, false, new(), null, default));
            Assert.Equal("valuable destination", File.ReadAllText(output));
            using var canceled = new CancellationTokenSource();
            canceled.Cancel();
            Assert.ThrowsAny<OperationCanceledException>(() => OfficeImageExporter.Export(review, output, true, new(), null, canceled.Token));
            var empty = await Source().PrepareSearchableOcrAsync(new DelegateOcrEngine("empty", (_, _) => Task.FromResult(new OcrResult())));
            Assert.Throws<InvalidOperationException>(() => OfficeImageExporter.Export(empty, output, true, new(), null, default));
            var textOnly = await Source().PrepareSearchableOcrAsync(new DelegateOcrEngine("text", (_, _) => Task.FromResult(new OcrResult {
                Spans = new[] { Word("paragraph", 20, 20, 65) }
            })));
            Assert.Throws<InvalidOperationException>(() => OfficeImageExporter.Export(textOnly, output, true, new(), null, default));
            Assert.Equal("valuable destination", File.ReadAllText(output));
            Assert.Equal(new[] { output }, Directory.GetFiles(directory));
        }
        finally { Directory.Delete(directory, recursive: true); }
    }

    [Fact]
    public async Task WhatIfCreatesNoDestinationAndDoesNotChangeCallerOptions()
    {
        string directory = NewDirectory();
        try
        {
            var review = await Source().PrepareSearchableOcrAsync(new DelegateOcrEngine("table", (_, _) => Task.FromResult(TableResult())));
            using var runspace = RunspaceFactory.CreateRunspace(InitialSessionState.CreateDefault());
            runspace.InitialSessionState.Commands.Add(new SessionStateCmdletEntry("ConvertFrom-OfficeImage", typeof(ConvertFromOfficeImageCommand), null));
            runspace.Open();
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            var options = new PdfTablesToExcelOptions();
            shell.AddCommand("ConvertFrom-OfficeImage").AddParameter("InputObject", review)
                .AddParameter("OutputPath", Path.Combine(directory, "preview.xlsx"))
                .AddParameter("ExcelOptions", options).AddParameter("FirstRowIsHeader").AddParameter("WhatIf");
            shell.Invoke();
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            Assert.Empty(Directory.GetFiles(directory));
            Assert.False(options.UseFirstRowAsHeader);
            shell.Commands.Clear();
            shell.AddCommand("ConvertFrom-OfficeImage").AddParameter("InputObject", review)
                .AddParameter("OutputPath", Path.Combine(directory, "actual.xlsx"))
                .AddParameter("ExcelOptions", options).AddParameter("FirstRowIsHeader").AddParameter("PassThru");
            var result = Assert.IsType<OfficeImageConversionResult>(Assert.Single(shell.Invoke()).BaseObject);
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            Assert.True(result.File.Exists);
            Assert.False(options.UseFirstRowAsHeader);
        }
        finally { Directory.Delete(directory, recursive: true); }
    }

    private static PdfImageDocumentSource Source() => new(OfficeRasterImageEncoder.Encode(
        new OfficeRasterImage(800, 400, OfficeColor.White), OfficeImageExportFormat.Png));

    private static OcrResult TableResult() => new()
    {
        Spans = new[] {
            Word("Description", 20, 20, 75), Word("Amount", 200, 20, 45),
            Word("Service", 20, 40, 45), Word("119.88", 200, 40, 40),
            Word("Credit", 20, 60, 40), Word("-13.50", 200, 60, 40),
            Word("Plan", 20, 80, 30), Word("89.00", 200, 80, 40)
        }
    };

    private static OcrTextSpan Word(string text, double x, double y, double width) => new()
    {
        Text = text, Confidence = 1, Level = OcrTextSpanLevel.Word, CoordinateUnit = OcrCoordinateUnit.Points,
        Region = new OcrRegion { X = x, Y = y, Width = width, Height = 10 }
    };

    private static string NewDirectory()
    {
        string path = Path.Combine(Path.GetTempPath(), "PSWriteOffice-image-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(path);
        return path;
    }
}
