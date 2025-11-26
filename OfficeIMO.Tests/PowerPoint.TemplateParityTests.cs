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

        private static void CompareManifests(string expectedPath, string actualPath) {
            using var exp = ZipFile.OpenRead(expectedPath);
            using var act = ZipFile.OpenRead(actualPath);

            var expParts = exp.Entries.Where(e => !e.FullName.EndsWith("/")).Select(e => e.FullName).OrderBy(x => x).ToArray();
            var actParts = act.Entries.Where(e => !e.FullName.EndsWith("/")).Select(e => e.FullName).OrderBy(x => x).ToArray();

            Assert.Equal(expParts, actParts);
        }

        private static void AssertHasEntries(string pptxPath, params string[] entries) {
            using var zip = ZipFile.OpenRead(pptxPath);
            var names = zip.Entries.Select(e => e.FullName).ToHashSet(StringComparer.OrdinalIgnoreCase);
            foreach (string entry in entries) {
                Assert.True(names.Contains(entry), $"Missing expected entry: {entry}");
            }
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
            CompareManifests("Assets/PowerPointTemplates/PowerPointBlank.pptx", outPath);
        }

        [Fact]
        public void TitleTemplateMatchesAdvanced() {
            string dir = TempDir();
            Directory.CreateDirectory(dir);
            var outPath = Path.Combine(dir, "Advanced PowerPoint.pptx");
            AdvancedPowerPoint.Example_AdvancedPowerPoint(dir, false);
            CompareManifests("Assets/PowerPointTemplates/PowerPointWithTitle.pptx", outPath);
        }

        [Fact]
        public void TablesChartsTemplateMatches() {
            string dir = TempDir();
            Directory.CreateDirectory(dir);
            var outPath = Path.Combine(dir, "Table Operations.pptx");
            TablesPowerPoint.Example_PowerPointTables(dir, false);
            CompareManifests("Assets/PowerPointTemplates/PowerPointWithTablesAndCharts.pptx", outPath);

            // Additional structure assertions to catch packaging regressions
            AssertHasEntries(outPath,
                "ppt/charts/chart1.xml",
                "ppt/charts/chart2.xml",
                "ppt/charts/style1.xml",
                "ppt/charts/style2.xml",
                "ppt/charts/colors1.xml",
                "ppt/charts/colors2.xml",
                "ppt/embeddings/Microsoft_Excel_Worksheet.xlsx",
                "ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx",
                "ppt/media/image1.png");
        }

        [Fact]
        public void BlankTemplateLayoutsAndTableStylesMatch() {
            string dir = TempDir();
            Directory.CreateDirectory(dir);
            var outPath = Path.Combine(dir, "Basic PowerPoint.pptx");
            BasicPowerPointDocument.Example_BasicPowerPoint(dir, false);

            Assert.Equal(
                CountEntries("Assets/PowerPointTemplates/PowerPointBlank.pptx", "ppt/slideLayouts/slideLayout"),
                CountEntries(outPath, "ppt/slideLayouts/slideLayout"));

            // Sizes can differ slightly; ensure tableStyles exists and is non-empty
            Assert.True(EntryLength(outPath, "ppt/tableStyles.xml") > 0);
        }
#endif
    }
}
