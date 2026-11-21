using OfficeIMO;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public class OfficeAccessibilityReportTests {
    [Fact]
    public void JsonCarriesEveryFindingFieldAndSaveJsonWritesTheSameText() {
        var report = new OfficeAccessibilityReport("PowerPoint", OfficeAccessibilityPolicyProfile.Strict, new List<OfficeAccessibilityFinding> {
            new(OfficeAccessibilitySeverity.Error, "accessibility-missing-document-title", "No \"title\".", "Document"),
            new(OfficeAccessibilitySeverity.Warning, "accessibility-low-contrast", "Line\nbreak", "Slide 2", 1, 7U, "Chart 1",
                isHeuristic: true, measuredValue: 2.3456D, requiredValue: 4.5D)
        });

        string json = report.ToJson(indented: false);

        Assert.Equal("{\"schemaVersion\":2,\"format\":\"PowerPoint\",\"profile\":\"Strict\",\"errorCount\":1,\"warningCount\":1,\"isSuccessful\":false,\"findings\":[" +
            "{\"severity\":\"Error\",\"code\":\"accessibility-missing-document-title\",\"message\":\"No \\\"title\\\".\",\"location\":\"Document\",\"locationIndex\":null,\"objectId\":null,\"objectName\":null,\"isHeuristic\":false,\"measuredValue\":null,\"requiredValue\":null}," +
            "{\"severity\":\"Warning\",\"code\":\"accessibility-low-contrast\",\"message\":\"Line\\nbreak\",\"location\":\"Slide 2\",\"locationIndex\":1,\"objectId\":7,\"objectName\":\"Chart 1\",\"isHeuristic\":true,\"measuredValue\":2.346,\"requiredValue\":4.5}" +
            "]}", json);
        OfficeAccessibilityException failure = Assert.Throws<OfficeAccessibilityException>(() => report.EnsureCompliant());
        Assert.Same(report, failure.Report);

        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"), "accessibility.json");
        try {
            report.SaveJson(path);
            Assert.Equal(report.ToJson(), File.ReadAllText(path));
        } finally {
            Directory.Delete(Path.GetDirectoryName(path)!, recursive: true);
        }
    }
}
