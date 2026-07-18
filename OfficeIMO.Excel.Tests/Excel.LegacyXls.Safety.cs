using OfficeIMO.Drawing.Internal;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Model;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public async Task FormatApi_ToXlsxToXlsAndLoadAsyncStream_RoundTrips() {
            string path = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(path);
            document.AddWorksheet("Data").CellValue(1, 1, "Explicit format API");

            byte[] xlsx = document.ToBytes();
            Assert.Equal(0x50, xlsx[0]);
            Assert.Equal(0x4b, xlsx[1]);

            byte[] xls = document.ToBytes(OfficeIMO.Excel.ExcelFileFormat.Xls);
            Assert.Equal(0xd0, xls[0]);
            Assert.Equal(0xcf, xls[1]);

            using var stream = new MemoryStream(xls);
            using ExcelDocument loaded = await ExcelDocument.LoadAsync(stream);
            Assert.Equal(ExcelFileFormat.Xls, loaded.SourceFormat);
            Assert.True(loaded.Sheets[0].TryGetCellText(1, 1, out string? value));
            Assert.Equal("Explicit format API", value);
        }

        [Fact]
        public void LegacyXls_LoadResult_CachesCompactAndAdvancedReports() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateMinimalWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);
            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound));

            Assert.Equal(1, result.Summary.WorksheetCount);
            Assert.False(result.Summary.HasImportErrors);
            Assert.Same(result.Summary, result.Summary);
            Assert.Same(result.ImportReport, result.CreateAdvancedImportReport());
        }

        [Fact]
        public void LegacyXls_LoadResult_CountsPreservedOnlyRecordsAsConversionLoss() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateUnsupportedFeatureWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);
            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound));
            Assert.NotEmpty(result.PreservedFeatures);

            result.Workbook.MutableUnsupportedFeatures.Clear();
            Assert.Empty(result.UnsupportedFeatures);

            Assert.True(result.HasUnsupportedFeatures);
            Assert.True(result.HasConversionLoss);
            Assert.True(result.Summary.HasConversionLoss);
            Assert.Throws<InvalidOperationException>(() => result.EnsureNoUnsupportedFeatures());
            Assert.Throws<InvalidOperationException>(() => result.EnsureNoConversionLoss());
        }

        [Fact]
        public void LegacyXls_FeatureReport_IncludesPreservedOnlyRecords() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateUnsupportedFeatureWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);
            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound));
            ExcelDocument document = result.Document;
            Assert.NotEmpty(document.LegacyXlsPreservedFeatures);

            typeof(ExcelDocument)
                .GetField("_legacyXlsUnsupportedFeatures", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!
                .SetValue(document, Array.Empty<LegacyXlsUnsupportedFeature>());

            ExcelFeatureReport report = document.InspectFeatures();
            ExcelFeatureFinding finding = Assert.Single(
                report.PreservedFeatures,
                feature => feature.Name == "Legacy XLS preserved records");
            Assert.Equal(document.LegacyXlsPreservedFeatures.Count, finding.Count);
            Assert.All(document.LegacyXlsPreservedFeatures, preserved =>
                Assert.Contains(finding.Details, detail => detail.Contains(preserved.Code, StringComparison.Ordinal)));
        }

        [Fact]
        public void Load_WhenInputIsLegacyWord_ReportsFormatMismatch() {
            string path = Path.Combine(
                AppContext.BaseDirectory,
                "Documents",
                "LegacyDocCorpus",
                "ComSimpleParagraphs.doc");

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() => ExcelDocument.Load(path));

            Assert.Contains("legacy Word document", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Load_WhenInputIsLegacyPowerPoint_ReportsFormatMismatch() {
            string path = Path.Combine(
                AppContext.BaseDirectory,
                "Documents",
                "LegacyPptCorpus",
                "BasicPowerPoint.ppt");

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() => ExcelDocument.Load(path));

            Assert.Contains("legacy PowerPoint presentation", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void FormatApi_UsesCanonicalExcelNamingWithoutLegacyAliases() {
            Type documentType = typeof(ExcelDocument);
            Type importOptionsType = typeof(LegacyXlsImportOptions);

            Assert.NotNull(documentType.GetMethod(nameof(ExcelDocument.ToBytes),
                new[] { typeof(ExcelFileFormat), typeof(ExcelSaveOptions) }));
            Assert.NotNull(documentType.GetMethod(nameof(ExcelDocument.ToStream),
                new[] { typeof(ExcelFileFormat), typeof(ExcelSaveOptions) }));
            Assert.DoesNotContain(documentType.GetMethods(), method => method.Name is "ToXlsx" or "ToXls" or "ToXlsxStream" or "ToXlsStream");
            Assert.DoesNotContain(documentType.GetMethods(), method => method.Name is "Save" or "SaveAsync" &&
                method.GetParameters().Any(parameter => parameter.ParameterType == typeof(bool)));
            Assert.DoesNotContain(documentType.GetMethods(), method => method.Name == "Open");
            Assert.Contains(documentType.GetMethods(), method => method.Name == nameof(ExcelDocument.OpenInApplication));
            Assert.DoesNotContain(documentType.GetMethods(), method => method.Name == "Close");
            Assert.Contains(documentType.GetMethods(), method => method.Name == nameof(ExcelDocument.SaveCopy));
            Assert.Null(documentType.GetProperty("WasLoadedFromLegacyXls"));
            Assert.Null(importOptionsType.GetProperty("MaxWorkbookStreamBytes"));
            Assert.Null(importOptionsType.GetProperty("ReportUnsupportedRecords"));
            Assert.NotNull(importOptionsType.GetProperty(nameof(LegacyXlsImportOptions.MaxInputBytes)));
            Assert.NotNull(importOptionsType.GetProperty(nameof(LegacyXlsImportOptions.ReportUnsupportedContent)));
        }
    }
}
