using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        [Trait("Category", "OfficeInteroperability")]
        public void OfficeInteroperability_CorpusManifest_TracksAndLoadsExcelArtifacts() {
            OfficeInteroperabilityCorpusManifest manifest = OfficeInteroperabilityCorpusManifestLoader.Load();
            IReadOnlyList<string> errors = OfficeInteroperabilityCorpusManifestLoader.Validate(manifest);
            Assert.True(errors.Count == 0, string.Join(Environment.NewLine, errors));

            OfficeInteroperabilityCorpusCollection[] collections = manifest.Collections
                .Where(collection => collection.Format is "xls" or "xlsb")
                .ToArray();
            Assert.Equal(5, collections.Length);
            Assert.Equal(29, collections.Sum(collection => collection.Artifacts.Count));
            Assert.Contains(collections, collection => collection.Format == "xlsb" && collection.Artifacts.Count == 5);
            Assert.Contains(collections, collection => collection.Role == "diagnostic");

            foreach (OfficeInteroperabilityCorpusCollection collection in collections.Where(item => item.Role == "compatibility")) {
                foreach (OfficeInteroperabilityCorpusArtifact artifact in collection.Artifacts) {
                    string path = OfficeInteroperabilityCorpusManifestLoader.ResolveArtifactPath(collection, artifact);
                    if (collection.Format == "xlsb") {
                        using ExcelDocument document = ExcelDocument.Load(path);
                        Assert.Equal(ExcelFileFormat.Xlsb, document.SourceFormat);
                        Assert.NotEmpty(document.Sheets);
                    } else {
                        using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(path, new LegacyXlsImportOptions {
                            ReportUnsupportedContent = true
                        });
                        Assert.False(result.HasImportErrors, $"{collection.Id}/{artifact.File}");
                        Assert.True(result.HasDocument, $"{collection.Id}/{artifact.File}");
                    }
                }
            }
        }
    }
}
