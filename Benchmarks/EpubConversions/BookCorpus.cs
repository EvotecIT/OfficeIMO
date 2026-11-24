using System.Globalization;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Epub;

namespace OfficeIMO.Benchmarks.EpubConversions;

public enum BookScale {
    Small,
    Large,
    ImageHeavy
}

internal static class BookCorpus {
    internal const string Title = "Field guide";
    private const string Stylesheet = "@font-face{font-family:FieldGuide;src:url('../fonts/field-guide.ttf') format('truetype')}" +
        "h1{font:22pt FieldGuide,sans-serif} p{margin-bottom:9pt} img{width:140px;height:105px;margin-right:6px}" +
        "img.mark{width:16px;height:16px}";

    internal static (int Chapters, int ImagesPerChapter) Size(BookScale scale) => scale switch {
        BookScale.Small => (10, 0),
        BookScale.Large => (500, 0),
        BookScale.ImageHeavy => (50, 4),
        _ => throw new ArgumentOutOfRangeException(nameof(scale))
    };

    internal static byte[] CreateEpub(BookScale scale) {
        var (chapters, imagesPerChapter) = Size(scale);
        var book = EpubPublication.Create(Title, "en", "urn:officeimo:benchmark:field-guide");
        var css = new StringBuilder(Stylesheet);
        for (int rule = 0; rule < 128; rule++) css.Append(".notation-").Append(rule).Append("{font-style:italic;margin-left:1em}");
        book.AddStylesheet("book-style", "EPUB/styles/book.css", css.ToString());
        book.AddResource("typeface", "EPUB/fonts/field-guide.ttf", "font/ttf",
            File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "OfficeIMOBaselineSans-Regular.ttf")));
        book.AddResource("mark", "EPUB/art/mark.svg", "image/svg+xml", Encoding.UTF8.GetBytes(
            "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='16' viewBox='0 0 16 16'><circle cx='8' cy='8' r='4' fill='#356772'/></svg>"));
        for (int chapter = 1; chapter <= chapters; chapter++) {
            string label = ChapterLabel(chapter);
            var body = new StringBuilder("<h1>").Append(label)
                .Append("</h1><img class='mark' src='../art/mark.svg' alt='Field guide mark' />");
            for (int paragraph = 0; paragraph < 5; paragraph++) {
                body.Append("<p>Field observations retain their order through the print edition. The published layout keeps the heading, spacing and illustrations together.</p>");
            }
            for (int image = 0; image < imagesPerChapter; image++) {
                int index = (chapter - 1) * imagesPerChapter + image;
                string name = "figure-" + index.ToString(CultureInfo.InvariantCulture);
                book.AddResource(name, "EPUB/art/" + name + ".png", "image/png", CreateImage(index));
                body.Append("<img src='../art/").Append(name).Append(".png' alt='Field illustration' />");
            }
            book.AddChapter("chapter-" + chapter, "EPUB/text/chapter-" + chapter + ".xhtml", label, body.ToString(), ["book-style"]);
        }
        return book.Write().Bytes;
    }

    internal static string ChapterLabel(int chapter) => "Field notes " + chapter.ToString("D4", CultureInfo.InvariantCulture);

    internal static byte[] CreateImage(int index) {
        var bitmap = new OfficeRasterImage(240, 180, OfficeColor.FromRgb((byte)index, 126, 138));
        for (int y = 12; y < 168; y++)
            for (int x = 12; x < 228; x++)
                bitmap.SetPixel(x, y, OfficeColor.FromRgb((byte)((x + index) % 256), (byte)y, 78));
        return OfficeRasterImageEncoder.Encode(bitmap, OfficeImageExportFormat.Png);
    }
}
