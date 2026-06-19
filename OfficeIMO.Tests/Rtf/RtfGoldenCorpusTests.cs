using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Diagnostics;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfGoldenCorpusTests {
    [Fact]
    public void GoldenCorpusFixturesParseWithoutErrors() {
        string corpusPath = Path.Combine(AppContext.BaseDirectory, "Documents", "RtfCorpus");
        string[] files = Directory.GetFiles(corpusPath, "*.rtf", SearchOption.AllDirectories);

        Assert.NotEmpty(files);

        foreach (string file in files) {
            RtfReadResult result = RtfDocument.Load(file);

            Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == RtfDiagnosticSeverity.Error);
            Assert.NotEmpty(result.ToRtfLossless());
        }
    }
}
