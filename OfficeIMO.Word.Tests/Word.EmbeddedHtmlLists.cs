using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    private const string MixedListHtml = """
        <html><body>
        <ul style="list-style-type:disc">
          <li value="1" class="bullet" data-note="value=8">Bullet</li>
          <li VALUE='2'>Second<ol start="4"><li value="4">Nested ordered</li></ol></li>
        </ul>
        <ol><li value="7">Ordered</li></ol>
        <script>const template = '<ul><li value="9">raw</li></ul>';</script>
        </body></html>
        """;

    [Fact]
    public void AddEmbeddedFragment_PreservesListKindsWhenUnorderedItemsDeclareValues() {
        using WordDocument document = WordDocument.Create();

        WordEmbeddedDocument embedded = document.AddEmbeddedFragment(
            MixedListHtml,
            WordAlternativeFormatImportPartType.Html);

        string storedHtml = Assert.IsType<string>(embedded.GetHtml());
        Assert.DoesNotContain("value=\"1\"", storedHtml, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("VALUE='2'", storedHtml, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("data-note=\"value=8\"", storedHtml, StringComparison.Ordinal);
        Assert.Contains("<ol start=\"4\"><li value=\"4\">Nested ordered", storedHtml, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<ol><li value=\"7\">Ordered", storedHtml, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("'<ul><li value=\"9\">raw</li></ul>'", storedHtml, StringComparison.Ordinal);
    }

    [Fact]
    public void AddEmbeddedFragmentAfter_AppliesTheSameListNormalization() {
        using WordDocument document = WordDocument.Create();
        WordParagraph anchor = document.AddParagraph("Before");

        WordEmbeddedDocument embedded = document.AddEmbeddedFragmentAfter(anchor, MixedListHtml);

        string storedHtml = Assert.IsType<string>(embedded.GetHtml());
        Assert.DoesNotContain("value=\"1\"", storedHtml, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<li value=\"4\">Nested ordered", storedHtml, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<li value=\"7\">Ordered", storedHtml, StringComparison.OrdinalIgnoreCase);
    }
}
