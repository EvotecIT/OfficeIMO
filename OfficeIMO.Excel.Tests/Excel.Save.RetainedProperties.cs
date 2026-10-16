using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData(false, false)]
        [InlineData(true, false)]
        [InlineData(false, true)]
        [InlineData(true, true)]
        public void Save_retained_core_properties_preserve_subsequent_edits(bool standardWriter, bool fileSave) {
            using var stream = new MemoryStream();
            using var document = ExcelDocument.Create(stream);
            document.AddWorksheet("Report").CellValue(1, 1, "Data");
            ExcelBuiltinDocumentProperties properties = document.BuiltinDocumentProperties;
            properties.Title = "Before save";
            var options = new ExcelSaveOptions { DisableFastPackageWriter = standardWriter };
            SaveRetainedMetadataFixture(document, stream, options, fileSave);

            properties.Title = "After save";
            Assert.Equal("After save", document.BuiltinDocumentProperties.Title);
            SaveRetainedMetadataFixture(document, stream, options, fileSave);

            using var saved = SpreadsheetDocument.Open(stream, false);
            Assert.Equal("After save", saved.PackageProperties.Title);
        }

        [Theory]
        [InlineData(false, false)]
        [InlineData(true, false)]
        [InlineData(false, true)]
        [InlineData(true, true)]
        public void Save_retained_application_properties_preserve_subsequent_edits(bool standardWriter, bool fileSave) {
            using var stream = new MemoryStream();
            using var document = ExcelDocument.Create(stream);
            document.AddWorksheet("Report").CellValue(1, 1, "Data");
            ExcelApplicationProperties properties = document.ApplicationProperties;
            properties.Company = "Before save";
            var options = new ExcelSaveOptions { DisableFastPackageWriter = standardWriter };
            SaveRetainedMetadataFixture(document, stream, options, fileSave);

            properties.Company = "After save";
            properties.Manager = "Report manager";
            properties.ApplicationName = "Metadata lifetime test";
            Assert.Equal("After save", document.ApplicationProperties.Company);
            SaveRetainedMetadataFixture(document, stream, options, fileSave);

            using var saved = SpreadsheetDocument.Open(stream, false);
            var metadata = saved.ExtendedFilePropertiesPart!.Properties!;
            Assert.Equal("After save", metadata.Company!.Text);
            Assert.Equal("Report manager", metadata.Manager!.Text);
            Assert.Equal("Metadata lifetime test", metadata.Application!.Text);
        }

        [Theory]
        [InlineData(false, false)]
        [InlineData(true, false)]
        [InlineData(false, true)]
        [InlineData(true, true)]
        public void Save_retained_custom_property_preserves_subsequent_edits(bool standardWriter, bool fileSave) {
            using var stream = new MemoryStream();
            using var document = ExcelDocument.Create(stream);
            document.AddWorksheet("Report").CellValue(1, 1, "Data");
            var property = new ExcelCustomProperty("Before save");
            document.SetCustomDocumentProperty("Review", property);
            var options = new ExcelSaveOptions { DisableFastPackageWriter = standardWriter };
            SaveRetainedMetadataFixture(document, stream, options, fileSave);

            property.Value = "After save";
            SaveRetainedMetadataFixture(document, stream, options, fileSave);

            using (var saved = SpreadsheetDocument.Open(stream, false)) {
                CustomDocumentProperty metadata = saved.CustomFilePropertiesPart!.Properties!
                    .Elements<CustomDocumentProperty>().Single(p => p.Name == "Review");
                Assert.Equal("After save", metadata.VTLPWSTR!.Text);
            }
            property.Value = true;
            property.PropertyType = ExcelCustomPropertyType.YesNo;
            SaveRetainedMetadataFixture(document, stream, options, fileSave);
            using var final = SpreadsheetDocument.Open(stream, false);
            Assert.Equal("true", final.CustomFilePropertiesPart!.Properties!
                .Elements<CustomDocumentProperty>().Single(p => p.Name == "Review").VTBool!.Text);
        }

        [Fact]
        public void Save_does_not_reconnect_a_replaced_custom_property() {
            using var stream = new MemoryStream();
            using var document = ExcelDocument.Create(stream);
            document.AddWorksheet("Report");
            var removed = new ExcelCustomProperty("Original");
            document.SetCustomDocumentProperty("Review", removed);
            var options = new ExcelSaveOptions();
            SaveRetainedMetadataFixture(document, stream, options, fileSave: true);
            document.RemoveCustomDocumentProperty("Review");
            document.SetCustomDocumentProperty("Review", "Replacement");
            SaveRetainedMetadataFixture(document, stream, options, fileSave: true);
            removed.Value = "Must stay detached";
            SaveRetainedMetadataFixture(document, stream, options, fileSave: true);

            using var saved = SpreadsheetDocument.Open(stream, false);
            Assert.Equal("Replacement", saved.CustomFilePropertiesPart!.Properties!
                .Elements<CustomDocumentProperty>().Single(p => p.Name == "Review").VTLPWSTR!.Text);
        }

        [Fact]
        public void Retained_property_wrappers_reject_access_after_document_disposal() {
            var document = ExcelDocument.Create();
            ExcelBuiltinDocumentProperties core = document.BuiltinDocumentProperties;
            ExcelApplicationProperties application = document.ApplicationProperties;
            document.Dispose();

            Assert.Throws<ObjectDisposedException>(() => core.Title);
            Assert.Throws<ObjectDisposedException>(() => core.Title = "Closed");
            Assert.Throws<ObjectDisposedException>(() => application.Company);
            Assert.Throws<ObjectDisposedException>(() => application.Company = "Closed");
        }

        private static void SaveRetainedMetadataFixture(ExcelDocument document, MemoryStream stream, ExcelSaveOptions options, bool fileSave) {
            if (!fileSave) {
                document.Save(stream, options);
                return;
            }
            string path = Path.Combine(Path.GetTempPath(), "OfficeIMO-RetainedMetadata-" + Guid.NewGuid().ToString("N") + ".xlsx");
            try {
                document.Save(path, options);
                stream.SetLength(0);
                using var saved = File.OpenRead(path);
                saved.CopyTo(stream);
                stream.Position = 0;
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }
    }
}
