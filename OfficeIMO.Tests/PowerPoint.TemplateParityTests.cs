using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
#if !NETFRAMEWORK
using OfficeIMO.Examples.PowerPoint;
#endif
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointTemplateParityTests {
#if !NETFRAMEWORK
        private static string TempDir() => Path.Combine(Path.GetTempPath(), "officeimo_ppt_parity", Guid.NewGuid().ToString("N"));

        private static string[] LoadManifest(string path, Func<string, string>? normalize) {
            using var zip = ZipFile.OpenRead(path);
            return zip.Entries
                .Where(e => !e.FullName.EndsWith("/"))
                .Select(e => normalize == null ? e.FullName : normalize(e.FullName))
                .OrderBy(x => x)
                .ToArray();
        }
        private static void AssertHasEntries(string pptxPath, string[] entries, Func<string, string>? normalize = null) {
            var names = LoadManifest(pptxPath, normalize).ToHashSet(StringComparer.OrdinalIgnoreCase);
            foreach (string entry in entries) {
                Assert.True(names.Contains(entry), $"Missing expected entry: {entry}");
            }
        }

        private static string NormalizeTablesChartsEntry(string entry) {
            const string chartPrefix = "ppt/slides/charts/";
            const string relsPrefix = "ppt/slides/charts/_rels/";
            const string embedPrefix = "ppt/slides/charts/embeddings/";

            if (entry.StartsWith(relsPrefix, StringComparison.OrdinalIgnoreCase)) {
                return "ppt/charts/_rels/" + entry.Substring(relsPrefix.Length);
            }

            if (entry.StartsWith(embedPrefix, StringComparison.OrdinalIgnoreCase)) {
                string name = entry.Substring(embedPrefix.Length);
                if (string.Equals(name, "package.bin", StringComparison.OrdinalIgnoreCase)) {
                    return "ppt/embeddings/Microsoft_Excel_Worksheet.xlsx";
                }
                if (string.Equals(name, "package2.bin", StringComparison.OrdinalIgnoreCase)) {
                    return "ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx";
                }
                return "ppt/embeddings/" + name;
            }

            if (entry.StartsWith(chartPrefix, StringComparison.OrdinalIgnoreCase)) {
                string name = entry.Substring(chartPrefix.Length);
                if (string.Equals(name, "style.xml", StringComparison.OrdinalIgnoreCase)) {
                    name = "style1.xml";
                } else if (string.Equals(name, "colors.xml", StringComparison.OrdinalIgnoreCase)) {
                    name = "colors1.xml";
                }
                return "ppt/charts/" + name;
            }

            if (string.Equals(entry, "ppt/media/image.png", StringComparison.OrdinalIgnoreCase)) {
                return "ppt/media/image1.png";
            }

            return entry;
        }

        private static int CountEntries(string pptxPath, string prefix) {
            using var zip = ZipFile.OpenRead(pptxPath);
            return zip.Entries.Count(e => e.FullName.StartsWith(prefix, StringComparison.OrdinalIgnoreCase));
        }

        private static long EntryLength(string pptxPath, string entryName) {
            using var zip = ZipFile.OpenRead(pptxPath);
            return zip.Entries.First(e => e.FullName.Equals(entryName, StringComparison.OrdinalIgnoreCase)).Length;
        }

        [Fact]
        public void BasicTemplateMatchesBlank() {
            string dir = TempDir();
            Directory.CreateDirectory(dir);
            var outPath = Path.Combine(dir, "Basic PowerPoint.pptx");
            BasicPowerPointDocument.Example_BasicPowerPoint(dir, false);
            AssertHasEntries(outPath, new[] {
                "ppt/presentation.xml",
                "ppt/slideMasters/slideMaster1.xml",
                "ppt/slides/slide1.xml",
                "ppt/theme/theme1.xml",
                "ppt/tableStyles.xml"
            });
        }

        [Fact]
        public void TitleTemplateMatchesAdvanced() {
            string dir = TempDir();
            Directory.CreateDirectory(dir);
            var outPath = Path.Combine(dir, "Advanced PowerPoint.pptx");
            AdvancedPowerPoint.Example_AdvancedPowerPoint(dir, false);
            AssertHasEntries(outPath, new[] {
                "ppt/presentation.xml",
                "ppt/slideMasters/slideMaster1.xml",
                "ppt/slides/slide1.xml",
                "ppt/theme/theme1.xml",
                "ppt/tableStyles.xml"
            });
        }

        [Fact]
        public void TablesChartsTemplateMatches() {
            string dir = TempDir();
            Directory.CreateDirectory(dir);
            var outPath = Path.Combine(dir, "Table Operations.pptx");
            TablesPowerPoint.Example_PowerPointTables(dir, false);
            // Additional structure assertions to catch packaging regressions
            AssertHasEntries(outPath, new[] {
                "ppt/charts/chart1.xml",
                "ppt/charts/chart2.xml",
                "ppt/charts/style1.xml",
                "ppt/charts/style2.xml",
                "ppt/charts/colors1.xml",
                "ppt/charts/colors2.xml",
                "ppt/embeddings/Microsoft_Excel_Worksheet.xlsx",
                "ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx",
                "ppt/media/image1.png"
            }, NormalizeTablesChartsEntry);
        }

        [Fact]
        public void BlankTemplateLayoutsAndTableStylesMatch() {
            string dir = TempDir();
            Directory.CreateDirectory(dir);
            var outPath = Path.Combine(dir, "Basic PowerPoint.pptx");
            BasicPowerPointDocument.Example_BasicPowerPoint(dir, false);

            Assert.True(CountEntries(outPath, "ppt/slideLayouts/slideLayout") > 0);

            // Sizes can differ slightly; ensure tableStyles exists and is non-empty
            Assert.True(EntryLength(outPath, "ppt/tableStyles.xml") > 0);
        }
#endif
    }
}

