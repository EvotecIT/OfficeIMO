#if NETFRAMEWORK
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Save_to_target_without_delete_sharing_preserves_target_and_retained_edits() {
            string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO-LockedSave-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            try {
                string targetPath = Path.Combine(directory, "target.xlsx");
                using (var existing = ExcelDocument.Create()) {
                    existing.AddWorksheet("Original target").CellValue(1, 1, "Target before failure");
                    existing.Save(targetPath);
                }
                byte[] originalTarget = File.ReadAllBytes(targetPath);

                using var document = ExcelDocument.Create();
                ExcelSheet retainedSheet = document.AddWorksheet("Report");
                retainedSheet.CellValue(1, 1, "Live before failure");
                var retainedProperty = new ExcelCustomProperty("Property before failure");
                document.SetCustomDocumentProperty("Review", retainedProperty);

                // Windows permits ordinary access but rejects replacing or moving this open target.
                using (var targetHandle = new FileStream(targetPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) {
                    Assert.Throws<IOException>(() => document.Save(targetPath));
                    Assert.Equal(originalTarget, File.ReadAllBytes(targetPath));
                }

                retainedSheet.CellValue(1, 1, "Live after failure");
                retainedProperty.Value = "Property after failure";
                Assert.Same(retainedProperty, document.CustomDocumentProperties["Review"]);
                document.Save(targetPath);

                using var reopened = ExcelDocument.Load(targetPath);
                Assert.Equal("Live after failure", reopened["Report"].CellAt(1, 1).GetValue<string>());
                Assert.Equal("Property after failure", reopened.CustomDocumentProperties["Review"].Value);
            } finally {
                Directory.Delete(directory, recursive: true);
            }
        }
    }
}
#endif
