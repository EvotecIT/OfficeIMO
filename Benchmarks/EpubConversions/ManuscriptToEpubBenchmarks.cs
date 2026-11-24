using System.Text;
using BenchmarkDotNet.Attributes;
using OfficeIMO.Epub;
using OfficeIMO.Word;
using OfficeIMO.Workflows;

namespace OfficeIMO.Benchmarks.EpubConversions;

[MemoryDiagnoser]
[BenchmarkCategory("Performance")]
public class ManuscriptToEpubBenchmarks {
    private const string Paragraph = "Field observations retain their order. This chapter records dates, places and illustrated details from the published manuscript.";
    private byte[] _source = [];

    [Params(BookScale.Small, BookScale.Large, BookScale.ImageHeavy)]
    public BookScale Scale { get; set; }

    [Params(OfficeDocumentFormat.Word, OfficeDocumentFormat.Markdown, OfficeDocumentFormat.Html)]
    public OfficeDocumentFormat Format { get; set; }

    [GlobalSetup(Target = nameof(Catalog))]
    public void SetupCatalog() {
        OfficeWorkflow.RegisterConversions();
        Prepare();
        using var output = new MemoryStream();
        ConvertCatalog(output);
        Verify(output.ToArray());
    }

    [GlobalSetup(Target = nameof(ImportAndWrite))]
    public void SetupImportAndWrite() {
        Prepare();
        Verify(ConvertImporter());
    }

    [Benchmark(Baseline = true)]
    public int ImportAndWrite() => ConvertImporter().Length;

    [Benchmark]
    public long Catalog() => ConvertCatalog(Stream.Null);

    private long ConvertCatalog(Stream output) {
        using var input = new MemoryStream(_source, writable: false);
        return OfficeConverter.Convert(input, Format, output, OfficeDocumentFormat.Epub).ItemsWritten;
    }

    private byte[] ConvertImporter() {
        var options = new EpubManuscriptOptions { Title = BookCorpus.Title };
        EpubManuscriptResult imported = BookManuscriptImporter.ImportBytesAsync(_source, Format switch {
            OfficeDocumentFormat.Word => ".docx",
            OfficeDocumentFormat.Markdown => ".md",
            _ => ".html"
        }, options).GetAwaiter().GetResult();
        return imported.RequireValue().Write().Bytes;
    }

    private void Prepare() {
        var (chapters, imagesPerChapter) = BookCorpus.Size(Scale);
        if (Format == OfficeDocumentFormat.Word) {
            using var word = WordDocument.Create();
            word.Properties.Title = BookCorpus.Title;
            for (int chapter = 1; chapter <= chapters; chapter++) {
                word.AddParagraph(BookCorpus.ChapterLabel(chapter)).SetStyle(WordParagraphStyles.Heading1);
                for (int paragraph = 0; paragraph < 5; paragraph++) word.AddParagraph(Paragraph);
                for (int image = 0; image < imagesPerChapter; image++) {
                    int index = (chapter - 1) * imagesPerChapter + image;
                    using var bytes = new MemoryStream(BookCorpus.CreateImage(index), writable: false);
                    word.AddParagraph().AddImage(bytes, $"figure-{index}.png", OfficeLength.Pixels(140), OfficeLength.Pixels(105),
                        new WordImageOptions { AltText = "Field illustration" });
                }
            }
            _source = word.ToBytes();
            return;
        }

        var text = new StringBuilder();
        bool markdown = Format == OfficeDocumentFormat.Markdown;
        if (markdown) text.Append("---\ntitle: ").Append(BookCorpus.Title).Append("\n---\n\n");
        else text.Append("<!doctype html><html lang='en'><title>").Append(BookCorpus.Title).Append("</title><body>");
        for (int chapter = 1; chapter <= chapters; chapter++) {
            if (markdown) text.Append("# ").Append(BookCorpus.ChapterLabel(chapter)).Append("\n\n");
            else text.Append("<h1>").Append(BookCorpus.ChapterLabel(chapter)).Append("</h1>");
            for (int paragraph = 0; paragraph < 5; paragraph++) {
                if (markdown) text.Append(Paragraph).Append("\n\n");
                else text.Append("<p>").Append(Paragraph).Append("</p>");
            }
            for (int image = 0; image < imagesPerChapter; image++) {
                int index = (chapter - 1) * imagesPerChapter + image;
                string data = Convert.ToBase64String(BookCorpus.CreateImage(index));
                if (markdown) text.Append("![Field illustration](data:image/png;base64,").Append(data).Append(")\n\n");
                else text.Append("<img alt='Field illustration' src='data:image/png;base64,").Append(data).Append("'>");
            }
        }
        if (!markdown) text.Append("</body></html>");
        _source = Encoding.UTF8.GetBytes(text.ToString());
    }

    private void Verify(byte[] bytes) {
        var (chapters, imagesPerChapter) = BookCorpus.Size(Scale);
        using var input = new MemoryStream(bytes, writable: false);
        EpubDocument book = EpubDocument.Load(input);
        if (book.Title != BookCorpus.Title || book.Chapters.Count != chapters
            || !book.Chapters[0].Text.Contains(BookCorpus.ChapterLabel(1), StringComparison.Ordinal)
            || !book.Chapters[chapters - 1].Text.Contains(BookCorpus.ChapterLabel(chapters), StringComparison.Ordinal)
            || book.Resources.Count(resource => resource.MediaType == "image/png") != chapters * imagesPerChapter)
            throw new InvalidDataException("The EPUB did not retain the manuscript title, chapter order, and images.");
    }
}
