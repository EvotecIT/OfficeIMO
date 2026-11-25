using Xunit;

namespace OfficeIMO.Tests {
    /// <summary>
    /// The Open XML SDK leaves OfficeIMO.Excel one feature area at a time. Its types are allowed in the bridge folder
    /// (OfficeIMO.Excel/OpenXmlSdk) and in the files that still build on the SDK model; that number may only shrink, and
    /// the engine's own Open XML layer never references the SDK.
    /// </summary>
    public class ExcelOpenXmlSdkBoundary {
        // Lower this when a feature area moves off the SDK; it must never rise.
        private const int MaximumFilesReferencingTheSdkOutsideTheBridge = 580;

        private static readonly string[] OwnedOpenXmlFiles = {
            "ExcelDeferredPackageParts.cs",
            "ExcelDeferredWorksheetWriter.cs",
            "ExcelDocument.DeferredLoad.cs",
            "ExcelDocument.DeferredSave.cs",
            "ExcelOwnedStyles.cs",
            "ExcelSheet.Deferred.cs",
            "ExcelSheet.DeferredEdits.cs",
        };

        [Fact]
        public void FilesReferencingTheSdkOutsideTheBridgeDoNotGrow() {
            List<string> files = FilesReferencingTheSdk().ToList();
            Assert.True(files.Count <= MaximumFilesReferencingTheSdkOutsideTheBridge,
                files.Count + " files outside OfficeIMO.Excel/OpenXmlSdk reference DocumentFormat.OpenXml; at most "
                + MaximumFilesReferencingTheSdkOutsideTheBridge + " may.");
        }

        [Fact]
        public void TheOwnedOpenXmlLayerDoesNotReferenceTheSdk() {
            string root = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.Excel");
            foreach (string file in OwnedOpenXmlFiles) {
                string text = File.ReadAllText(Path.Combine(root, file));
                Assert.False(text.Contains("DocumentFormat.OpenXml", StringComparison.Ordinal), file + " references the Open XML SDK.");
            }
        }

        private static IEnumerable<string> FilesReferencingTheSdk() {
            string root = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.Excel");
            string bridge = Path.Combine(root, "OpenXmlSdk") + Path.DirectorySeparatorChar;
            foreach (string file in Directory.EnumerateFiles(root, "*.cs", SearchOption.AllDirectories)) {
                string relative = Path.GetRelativePath(root, file);
                string first = relative.Split(Path.DirectorySeparatorChar)[0];
                if (first is "bin" or "obj" || file.StartsWith(bridge, StringComparison.Ordinal)) continue;
                if (File.ReadAllText(file).Contains("DocumentFormat.OpenXml", StringComparison.Ordinal)) yield return relative;
            }
        }
    }
}
