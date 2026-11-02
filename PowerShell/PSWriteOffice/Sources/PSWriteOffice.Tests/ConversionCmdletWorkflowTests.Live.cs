using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Pdf;
using OfficeIMO.Word;

namespace PSWriteOffice.Tests;

public sealed partial class ConversionCmdletWorkflowTests {
    [Fact]
    public void LiveDocumentConversionUsesCurrentEditsWithoutSavingOrClosingTheSource() {
        string root = Directory.CreateTempSubdirectory("officeimo-conversion-live-").FullName;
        try {
            string source = Path.Combine(root, "source.docx"), outputPath = Path.Combine(root, "review.pdf");
            using var document = WordDocument.Create(source);
            document.AddParagraph("Saved source");
            document.Save();
            byte[] saved = File.ReadAllBytes(source);
            var paragraph = document.AddParagraph("Current unsaved evidence");
            using var runspace = OpenRunspace();
            runspace.SessionStateProxy.SetVariable("document", document);
            runspace.SessionStateProxy.SetVariable("outputPath", outputPath);
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddScript("""
                Convert-OfficeDocument -Document $document -To Pdf -OutputPath $outputPath -WhatIf -ErrorAction Stop
                if (Test-Path -LiteralPath $outputPath) { throw 'WhatIf published a conversion' }
                $document | Convert-OfficeDocument -To Html -SummaryVariable htmlReport -ErrorAction Stop
                Convert-OfficeDocument -Document $document -To Pdf -OutputPath $outputPath -SummaryVariable pdfReport -ErrorAction Stop
                $pdfReport
                """);
            var output = shell.Invoke();
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            Assert.Equal(3, output.Count);
            Assert.Contains("Current unsaved evidence", Assert.IsType<string>(output[0].BaseObject));
            Assert.Equal(outputPath, Assert.IsType<OfficeConversionReport>(output[1].BaseObject).OutputPath);
            Assert.Same(output[1].BaseObject, output[2].BaseObject);
            using var pdf = PdfDocument.Load(outputPath);
            Assert.Contains("Current unsaved evidence", pdf.Read().Text);
            Assert.Equal(saved, File.ReadAllBytes(source));
            Assert.Equal(source, document.FilePath);
            Assert.Contains(paragraph, document.Paragraphs);
            paragraph.Text = "Still editable";
            Assert.Equal("Still editable", paragraph.Text);
        } finally { Directory.Delete(root, true); }
    }
    [Theory]
    [InlineData("Excel", "Pdf")]
    [InlineData("PowerPoint", "Html")]
    [InlineData("Pdf", "Html")]
    [InlineData("Markdown", "Html")]
    [InlineData("Html", "Markdown")]
    [InlineData("Rtf", "Pdf")]
    [InlineData("Odf", "Pdf")]
    [InlineData("Epub", "Pdf")]
    [InlineData("Visio", "Pdf")]
    public void LiveNativeFamiliesUseTheirOwnSerializerAndTheSharedConversionRoute(string kind, string target) {
        object document = LiveSource(kind);
        string root = Directory.CreateTempSubdirectory("officeimo-live-family-").FullName;
        try {
            using var runspace = OpenRunspace();
            runspace.SessionStateProxy.SetVariable("document", document);
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddCommand("Convert-OfficeDocument").AddParameter("Document", document).AddParameter("To", target);
            bool binary = target == "Pdf";
            string path = Path.Combine(root, "converted.pdf");
            if (binary) shell.AddParameter("OutputPath", path);
            var result = shell.Invoke();
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            if (binary) {
                Assert.IsType<OfficeConversionReport>(Assert.Single(result).BaseObject);
                using var pdf = PdfDocument.Load(path);
                Assert.Contains("Live source evidence", pdf.Read().Text);
            } else {
                Assert.Contains("Live source evidence", Assert.IsType<string>(Assert.Single(result).BaseObject));
            }
        } finally {
            (document as IDisposable)?.Dispose();
            Directory.Delete(root, true);
        }
    }

    [Theory]
    [InlineData("-To Pdf", "binary target")]
    [InlineData("-To Html -AsString -OutputPath $outputPath", "AsString")]
    [InlineData("-To Html -OutputPath $outputPath", "same target format")]
    [InlineData("-To Html -From Markdown", "native format")]
    [InlineData("-To Pdf -OutputPath $outputPath -NoClobber -AutoRename", "cannot be combined")]
    public void InvalidLiveConversionOptionsLeaveAnExistingOutputUnchanged(string arguments, string message) {
        string root = Directory.CreateTempSubdirectory("officeimo-live-invalid-").FullName;
        try {
            string path = Path.Combine(root, "output.pdf");
            byte[] original = [1, 2, 3, 4];
            File.WriteAllBytes(path, original);
            using var document = WordDocument.Create();
            document.AddParagraph("Retained live content");
            using var runspace = OpenRunspace();
            runspace.SessionStateProxy.SetVariable("document", document);
            runspace.SessionStateProxy.SetVariable("outputPath", path);
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddScript("Convert-OfficeDocument -Document $document " + arguments + " -ErrorAction Stop");
            Exception? error = Record.Exception(() => shell.Invoke());
            Assert.True(error != null || shell.HadErrors);
            string detail = error?.ToString() ?? string.Join(Environment.NewLine, shell.Streams.Error);
            Assert.Contains(message, detail, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(original, File.ReadAllBytes(path));
            Assert.Single(document.Paragraphs);
        } finally { Directory.Delete(root, true); }
    }

    [Theory]
    [InlineData("Word")]
    [InlineData("Excel")]
    [InlineData("PowerPoint")]
    public void LiveConversionPreservesTheCallersNativeEncryptionPolicy(string kind) {
        const string password = "synthetic-live-conversion-password";
        object original = kind == "Word" ? WordDocument.Create() : LiveSource(kind);
        if (original is WordDocument word) word.AddParagraph("Live source evidence");
        object? live = null;
        string root = Directory.CreateTempSubdirectory("officeimo-live-security-").FullName;
        try {
            byte[] encrypted = NativeBytes(original, password);
            using var input = new MemoryStream(encrypted);
            live = kind switch {
                "Word" => WordDocument.Load(input, new WordLoadOptions { Password = password }),
                "Excel" => OfficeIMO.Excel.ExcelDocument.Load(input, new OfficeIMO.Excel.ExcelLoadOptions { Password = password }),
                _ => OfficeIMO.PowerPoint.PowerPointDocument.Load(input, new OfficeIMO.PowerPoint.PowerPointLoadOptions { Password = password })
            };
            using var runspace = OpenRunspace();
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            string path = Path.Combine(root, "converted.pdf");
            shell.AddCommand("Convert-OfficeDocument").AddParameter("Document", live).AddParameter("To", "Pdf").AddParameter("OutputPath", path);
            var result = shell.Invoke();
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            Assert.IsType<OfficeConversionReport>(Assert.Single(result).BaseObject);
            using var converted = PdfDocument.Load(path);
            Assert.Contains("Live source evidence", converted.Read().Text);
            byte[] retained = NativeBytes(live);
            Assert.Equal(new byte[] { 0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1 }, retained.Take(8));
            using var nextInput = new MemoryStream(retained);
            using IDisposable reopened = kind switch {
                "Word" => WordDocument.Load(nextInput, new WordLoadOptions { Password = password }),
                "Excel" => OfficeIMO.Excel.ExcelDocument.Load(nextInput, new OfficeIMO.Excel.ExcelLoadOptions { Password = password }),
                _ => OfficeIMO.PowerPoint.PowerPointDocument.Load(nextInput, new OfficeIMO.PowerPoint.PowerPointLoadOptions { Password = password })
            };
        } finally {
            (live as IDisposable)?.Dispose();
            (original as IDisposable)?.Dispose();
            Directory.Delete(root, true);
        }
    }

    private static byte[] NativeBytes(object document, string? password = null) => document switch {
        WordDocument word => word.ToBytes(password == null ? null : new WordSaveOptions { Password = password }),
        OfficeIMO.Excel.ExcelDocument excel => excel.ToBytes(password == null ? null : new OfficeIMO.Excel.ExcelSaveOptions { Password = password }),
        OfficeIMO.PowerPoint.PowerPointDocument slides => slides.ToBytes(password == null ? null : new OfficeIMO.PowerPoint.PowerPointSaveOptions { Password = password }),
        _ => throw new ArgumentException("Choose a native Office fixture.", nameof(document))
    };

    private static object LiveSource(string kind) {
        const string text = "Live source evidence";
        switch (kind) {
            case "Excel":
                var excel = OfficeIMO.Excel.ExcelDocument.Create();
                excel.AddSheet("Evidence").CellValue(1, 1, text);
                return excel;
            case "PowerPoint":
                var slides = OfficeIMO.PowerPoint.PowerPointDocument.Create();
                slides.AddSlide().AddTextBox(text);
                return slides;
            case "Pdf":
                var pdf = PdfDocument.Create();
                pdf.Compose(compose => compose.Content(content => content.Paragraph(p => p.Text(text))));
                return pdf;
            case "Markdown": return OfficeIMO.Markdown.MarkdownDocument.Create().H1(text);
            case "Html": return OfficeIMO.Html.HtmlConversionDocument.Parse("<h1>" + text + "</h1>");
            case "Rtf":
                var rtf = OfficeIMO.Rtf.RtfDocument.Create();
                rtf.AddParagraph(text);
                return rtf;
            case "Odf":
                var odf = OfficeIMO.OpenDocument.OdgDocument.Create();
                odf.AddPage("Evidence").Shapes.AddTextBox(OfficeIMO.OpenDocument.OdfRect.FromCentimeters(1, 1, 10, 2), text);
                return odf;
            case "Epub":
                var epub = OfficeIMO.Epub.EpubPublication.Create("Evidence");
                epub.AddChapter("opening", "EPUB/text/opening.xhtml", "Opening", "<p>" + text + "</p>");
                return epub;
            case "Visio":
                var diagram = OfficeIMO.Visio.VisioDocument.Create();
                diagram.AddPage("Evidence").AddRectangle(4, 4, 4, 1, text, OfficeIMO.Visio.VisioMeasurementUnit.Inches);
                return diagram;
            default: throw new ArgumentException("Unknown fixture kind.", nameof(kind));
        }
    }

}
