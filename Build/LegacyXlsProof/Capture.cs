using System.Runtime.CompilerServices;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void LegacyXls_CaptureStreamingOutput() {
            string folder = Path.Combine(LegacyXlsProof.Output, "test-files", nameof(LegacyXls_CaptureStreamingOutput));
            Directory.CreateDirectory(folder);
            foreach (bool sharedStrings in new[] { false, true }) {
                using var output = new MemoryStream();
                using (var writer = ExcelDocument.CreateWriter(ExcelFileFormat.Xls, output,
                    new ExcelWriterOptions { UseSharedStrings = sharedStrings }, OfficeTheme.Modern.WithName("Legacy XLS proof"), leaveOpen: true)) {
                    writer.WriteTable("Data", new[] { 1, 2 }, new[] { "Text", "Number", "Date" }, static (row, index) => {
                        row.Write("Value " + index.ToString(System.Globalization.CultureInfo.InvariantCulture));
                        row.Write(index * 3);
                        row.Write(new DateTime(2025, 1, index));
                    });
                }
                File.WriteAllBytes(Path.Combine(folder, sharedStrings ? "shared.xls" : "inline.xls"), output.ToArray());
            }
        }

        [Fact]
        public void LegacyXls_CaptureFixtureProjections() {
            string assets = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Documents");
            foreach (string corpus in new[] { "LegacyXlsCorpus", "LegacyXlsDiagnosticCorpus", "SpreadsheetDateSerialCorpus" }) {
                string source = Path.Combine(assets, corpus);
                foreach (string path in Directory.EnumerateFiles(source, "*.xls", SearchOption.AllDirectories).OrderBy(path => path, StringComparer.Ordinal)) {
                    string relative = Path.GetRelativePath(assets, path);
                    string output = Path.Combine(LegacyXlsProof.Output, "fixtures", relative);
                    Directory.CreateDirectory(Path.GetDirectoryName(output)!);
                    using LegacyXlsLoadResult imported = ExcelDocument.LoadLegacyXlsWithReport(path);
                    if (imported.HasDocument) {
                        File.WriteAllBytes(output + ".xlsx", imported.Document.ToBytes(new ExcelSaveOptions {
                            AllowLoss = true
                        }));
                    } else {
                        File.WriteAllText(output + ".failure.txt", imported.ProjectionException?.GetType().FullName ?? "Projection unavailable");
                    }
                    File.WriteAllLines(output + ".diagnostics.txt", imported.NativeDiagnostics.Select(diagnostic =>
                        $"{diagnostic.Severity} {diagnostic.Code} {diagnostic.Message}"));
                }
            }
        }
    }

    internal static class LegacyXlsProof {
        private static readonly Dictionary<(string Test, string Extension), int> OutputCounts = new();

        internal static string Output => Environment.GetEnvironmentVariable("OFFICEIMO_LEGACY_XLS_PROOF_OUTPUT")
            ?? throw new InvalidOperationException("The comparison runner must specify the evidence directory.");

        internal static void Delete(string path, [CallerMemberName] string test = "") {
            if (!File.Exists(path)) return;
            string extension = Path.GetExtension(path);
            if (extension is ".xls" or ".xlsx" or ".xlsm") {
                string directory = Path.Combine(Output, "test-files", test);
                Directory.CreateDirectory(directory);
                lock (OutputCounts) {
                    var key = (test, extension);
                    OutputCounts.TryGetValue(key, out int index);
                    string stem = Path.Combine(directory, (index + 1).ToString("000", System.Globalization.CultureInfo.InvariantCulture));
                    File.Copy(path, stem + extension);
                    if (extension == ".xls" && test is ("LegacyXls_NativeSave_WritesCalculationSettings" or "LegacyXls_NativeSave_WritesWorkbookOptions")) {
                        using ExcelDocument document = ExcelDocument.Load(path);
                        File.WriteAllBytes(stem + ".projected.xlsx", document.ToBytes(new ExcelSaveOptions { AllowLoss = true }));
                    }
                    OutputCounts[key] = index + 1;
                }
            }
            File.Delete(path);
        }
    }
}
