using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using OfficeIMO.Examples.PowerPoint;
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
        }
#endif
    }
}
