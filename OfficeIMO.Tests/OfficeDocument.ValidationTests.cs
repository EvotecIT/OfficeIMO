using System.IO.Compression;
using OfficeIMO.Excel;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.Tests;

public sealed class OfficeDocumentValidationTests {
    [Theory]
    [InlineData("source.docx")]
    [InlineData("source.xlsx")]
    [InlineData("source.pptx")]
    [InlineData("source.pdf")]
    public void NativeFileStructureChecksRetainTheirEvidenceAndLeaveTheFileUnchanged(string name) {
        string root = Directory.CreateTempSubdirectory("officeimo-validation-file-").FullName;
        try {
            string path = Path.Combine(root, name);
            switch (System.IO.Path.GetExtension(name)) {
                case ".docx":
                    using (var word = WordDocument.Create()) { word.AddParagraph("Structure"); word.Save(path); }
                    break;
                case ".xlsx":
                    using (var book = ExcelDocument.Create()) { book.AddSheet("Data").Cell(1, 1).SetValue("Structure"); book.Save(path); }
                    break;
                case ".pptx":
                    using (var presentation = PowerPointDocument.Create()) { presentation.AddSlide().AddTextBox("Structure"); presentation.Save(path); }
                    break;
                case ".pdf":
                    using (var pdf = PdfDocument.Create()) {
                        pdf.Compose(builder => builder.Content(content => content.AddParagraph("Structure")));
                        pdf.Save(path, new PdfSaveOptions { AllowLoss = true });
                    }
                    break;
            }
            byte[] before = File.ReadAllBytes(path);
            var report = OfficeDocument.Validate(path);
            Assert.NotEmpty(report.NativeReports);
            Assert.DoesNotContain(report.Diagnostics, finding => finding.Code == "Validation.NotApplicable");
            Assert.Contains(report.Diagnostics, finding => finding.Code == "Validation.Source" && finding.Location == path);
            Assert.Equal(before, File.ReadAllBytes(path));
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void LiveWorkbookCapabilityValidationSeesUnsavedFormulasWithoutSavingOrTakingOwnership() {
        string root = Directory.CreateTempSubdirectory("officeimo-validation-").FullName;
        try {
            string path = Path.Combine(root, "book.xlsx");
            using var book = ExcelDocument.Create();
            var sheet = book.AddSheet("Data");
            sheet.Cell(1, 1).SetValue(7);
            book.Save(path);
            byte[] before = File.ReadAllBytes(path);
            sheet.Cell(1, 2).SetFormula("UNSUPPORTED(A1)");
            var options = new OfficeExcelValidationOptions {
                Doctor = new ExcelWorkbookDoctorOptions { ValidateOpenXml = false },
                Capability = new[] { ExcelPreflightCapability.CalculateFormulas },
                IncludeFeatures = true, IncludeRepairHints = true
            };
            var report = OfficeDocument.Validate(book, options: new OfficeDocumentValidationOptions { Excel = options });
            Assert.False(report.IsValid);
            Assert.Contains(report.Diagnostics, finding => finding.Code == "Excel.Capability.CalculateFormulas" && finding.Severity == OfficeDiagnosticSeverity.Error);
            Assert.Contains(report.NativeReports, item => item is ExcelFeatureReport);
            Assert.Contains("Excel.Capability.CalculateFormulas", report.ToMarkdown());
            Assert.Throws<InvalidOperationException>(() => report.ThrowIfFailed());
            Assert.Equal("UNSUPPORTED(A1)", sheet.GetFormulaText(1, 2));
            Assert.Equal(before, File.ReadAllBytes(path));
            sheet.Cell(2, 1).SetValue("Still caller-owned");
            options.Doctor.RepairDefinedNames = true;
            Assert.Throws<ArgumentException>(() => OfficeDocument.Validate(book, options: new OfficeDocumentValidationOptions { Excel = options }));
            Assert.True(options.Doctor.RepairDefinedNames);
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void UntrustedSecurityRetainsNativeInventoryAndNeverChangesThePackageOrSuppliedPolicy() {
        string root = Directory.CreateTempSubdirectory("officeimo-validation-security-").FullName;
        try {
            string path = Path.Combine(root, "upload.docm");
            using (var stream = File.Create(path))
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Create)) {
                using (var payload = archive.CreateEntry("word/vbaProject.bin").Open()) payload.WriteByte(1);
                using var writer = new StreamWriter(archive.CreateEntry("word/_rels/document.xml.rels").Open());
                writer.Write("<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink' Target='https://example.org' TargetMode='External'/></Relationships>");
            }
            byte[] before = File.ReadAllBytes(path);
            var policy = OfficePackageSecurityOptions.SecureDefaults;
            var report = OfficeDocument.Validate(path, new[] { OfficeDocumentCheck.Security },
                new OfficeDocumentValidationOptions { Untrusted = true, Security = policy });
            Assert.False(report.IsValid);
            var native = Assert.Single(report.NativeReports.OfType<OfficePackageSecurityReport>());
            Assert.Equal(1, native.MacroPartCount);
            Assert.Equal(1, native.ExternalRelationshipCount);
            Assert.Contains(report.Diagnostics, finding => finding.Code == "Security.Macros" && finding.Severity == OfficeDiagnosticSeverity.Error);
            Assert.Equal(OfficePackageContentPolicy.Allow, policy.Macros);
            Assert.Equal(before, File.ReadAllBytes(path));
            string unknown = Path.Combine(root, "unknown.bin");
            File.WriteAllText(unknown, "not an Office container");
            var oversized = OfficeDocument.Validate(unknown, new[] { OfficeDocumentCheck.Security },
                new OfficeDocumentValidationOptions { Security = new OfficePackageSecurityOptions { MaxPackageBytes = 1 } });
            Assert.False(oversized.IsValid);
            Assert.Contains(oversized.Diagnostics, finding => finding.Code == "Security.PackageSize");
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void BorrowedWordSchemaUsesTheLiveModelAndUnavailableChecksExplainTheirScope() {
        using var document = WordDocument.Create();
        document.AddParagraph("Unsaved content");
        var report = OfficeDocument.Validate(document, new[] { OfficeDocumentCheck.Schema, OfficeDocumentCheck.Security, OfficeDocumentCheck.Accessibility });
        Assert.Contains(report.NativeReports, evidence => evidence is IReadOnlyList<OfficeOpenXmlValidationError>);
        Assert.Contains(report.Diagnostics, finding => finding.Code == "Validation.NotApplicable" && finding.Location == "Security");
        Assert.Contains(report.Diagnostics, finding => finding.Code == "Validation.NotApplicable" && finding.Location == "Accessibility");
        Assert.Equal("Unsaved content", document.Paragraphs[0].Text);
        document.AddParagraph("The caller can continue editing");
    }

    [Fact]
    public void PdfChecksRetainPreflightComplianceAndRewriteMismatchEvidence() {
        using var reference = PdfDocument.Create();
        reference.Compose(builder => builder.Content(content => content.AddParagraph("Reference")));
        using var difference = PdfDocument.Create();
        difference.Compose(builder => builder.Content(content => content.AddParagraph("Different")));
        var rewrite = new PdfRewritePreservationOptions().RequireTextMarkers("Reference");
        var report = OfficeDocument.Validate(reference, new[] { OfficeDocumentCheck.Structure, OfficeDocumentCheck.Compliance },
            new OfficeDocumentValidationOptions { Pdf = new OfficePdfValidationOptions {
                Profile = PdfComplianceProfile.PdfA2B, DifferenceDocument = difference, Rewrite = rewrite
            } });
        Assert.False(report.IsValid);
        Assert.Contains(report.NativeReports, item => item is PdfDocumentPreflight);
        Assert.Contains(report.NativeReports, item => item is PdfComplianceReadinessReport);
        Assert.False(Assert.Single(report.NativeReports.OfType<PdfRewritePreservationReport>()).IsPreserved);
        Assert.Contains(report.Diagnostics, finding => finding.Code.StartsWith("Pdf.Preservation.") && finding.Severity == OfficeDiagnosticSeverity.Error);
        Assert.Single(rewrite.RequiredTextMarkers);
    }
}
