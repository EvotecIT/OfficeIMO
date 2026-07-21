using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    [LegacyXlsComFact]
    [Trait("Category", "MicrosoftOfficeInteroperability")]
    public void OfficeInteroperability_CompatibilityCorpusSourcesAndConversionsOpenInDesktopExcelWhenRequested() {
        OfficeInteroperabilityCorpusManifest manifest = OfficeInteroperabilityCorpusManifestLoader.Load();
        IReadOnlyList<string> errors = OfficeInteroperabilityCorpusManifestLoader.Validate(manifest);
        Assert.True(errors.Count == 0, string.Join(Environment.NewLine, errors));

        OfficeInteroperabilityCorpusCollection[] collections = manifest.Collections
            .Where(collection => collection.Role == "compatibility")
            .Where(collection => collection.Format is "xls" or "xlsb")
            .Where(collection => collection.Oracles.Contains("microsoft-office-open"))
            .ToArray();
        Assert.NotEmpty(collections);

        var sourcePaths = new List<string>();
        var convertedPaths = new List<string>();
        string outputDirectory = Path.Combine(
            _directoryWithFiles,
            "OfficeInteroperabilityExcelCom",
            GetCurrentTargetFrameworkLabel());
        Directory.CreateDirectory(outputDirectory);

        foreach (OfficeInteroperabilityCorpusCollection collection in collections) {
            int index = 0;
            foreach (OfficeInteroperabilityCorpusArtifact artifact in collection.Artifacts) {
                index++;
                string source = OfficeInteroperabilityCorpusManifestLoader.ResolveArtifactPath(collection, artifact);
                sourcePaths.Add(source);
                string destination = Path.Combine(
                    outputDirectory,
                    $"{collection.Id}-{index:D3}-{Path.GetFileNameWithoutExtension(artifact.File)}.xlsx");
                ExcelDocument.Convert(source, destination, new ExcelDocumentConversionOptions {
                    CompatibilityMode = OfficeCompatibilityMode.BestEffort,
                    LossPolicy = ExcelConversionLossPolicy.Allow,
                    FileConflictPolicy = ExcelConversionFileConflictPolicy.Replace
                });
                convertedPaths.Add(destination);
            }
        }

        string generatedModern = Path.Combine(outputDirectory, "officeimo-modern-source.xlsx");
        using (ExcelDocument workbook = ExcelDocument.Create(generatedModern)) {
            ExcelSheet sheet = workbook.AddWorksheet("Compatibility");
            sheet.CellValue(1, 1, "OfficeIMO modern to binary");
            sheet.CellValue(2, 1, 42);
            workbook.Save();
        }
        foreach (string extension in new[] { ".xls", ".xlsb" }) {
            string generatedBinary = Path.Combine(outputDirectory, "officeimo-modern-to-binary" + extension);
            ExcelDocument.Convert(generatedModern, generatedBinary, new ExcelDocumentConversionOptions {
                FileConflictPolicy = ExcelConversionFileConflictPolicy.Replace
            }).RequireNoLoss();
            convertedPaths.Add(generatedBinary);
        }

        AssertWorkbooksOpenViaExcelComWhenAvailable(
            sourcePaths,
            "One or more declared Excel corpus sources failed to open in desktop Excel.");
        AssertWorkbooksOpenViaExcelComWhenAvailable(
            convertedPaths,
            "One or more OfficeIMO corpus conversions failed to open in desktop Excel.");
    }
}
