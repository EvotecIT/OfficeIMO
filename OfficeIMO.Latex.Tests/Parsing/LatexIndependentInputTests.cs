using System.Security.Cryptography;

namespace OfficeIMO.Latex.Tests;

public sealed class LatexIndependentInputTests {
    [Theory]
    [InlineData("small2e.tex", "6995024E85F537D32EEF704A81D7F53B84AA06CB23361559D3FEA3B6C5653627", 0)]
    [InlineData("sample2e.tex", "F135855F870C31F1101001BDB75B11E94A23C2605F7AB5DFFA1299D53B1977CC", 1)]
    public void LaTeXProject_inputs_keep_original_bytes_and_typed_semantic_ownership(string file, string hash, int notes) {
        string path = Path.Combine(AppContext.BaseDirectory, "Fixtures", "latex2e", file);
        byte[] bytes = File.ReadAllBytes(path);
        using (SHA256 algorithm = SHA256.Create())
            Assert.Equal(hash, BitConverter.ToString(algorithm.ComputeHash(bytes)).Replace("-", string.Empty));
        LatexDocument document = LatexDocument.Load(path);
        Assert.True(document.SyntaxTree.IsLossless);
        Assert.Equal(File.ReadAllText(path), document.ToLatex());
        Assert.Equal(document.Source.Text, string.Concat(document.Tokens.Select(static token => token.Text)));
        Assert.Equal(notes, document.Footnotes.Count);
        Assert.DoesNotContain(document.Diagnostics, static diagnostic => diagnostic.Severity == LatexDiagnosticSeverity.Error);
        if (notes > 0) {
            Assert.Equal("This is an example of a footnote.", Assert.Single(document.Footnotes).Content);
            Assert.Contains(document.Paragraphs, static paragraph => paragraph.Content.Contains(@"Footnotes\footnote{This is an example of a footnote.}"));
            Assert.Equal(new[] { 3, 2 }, document.Lists.Select(static list => list.Items.Count));
        }
    }
}
