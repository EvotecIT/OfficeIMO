using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Word;
using OfficeIMO.Word;
using Xunit;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public sealed class ReaderWordTextProjectionTests {
    private const string StreamedCapability = "officeimo.reader.word.streamed-text";

    // A text box streams too: its text is read inline as the model reads it.
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TextProjectionReturnsOnlyTextChunksWithTheTextOfTheLeanRead(bool withTextBox) {
        byte[] package = CreateDocument(withTextBox);
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddWordHandler().Build();

        OfficeDocumentReadResult lean = reader.ReadDocument(package, "report.docx",
            new ReaderOptions { Projection = ReaderProjection.TextAndMetadata });
        OfficeDocumentReadResult text = reader.ReadDocument(package, "report.docx",
            new ReaderOptions { Projection = ReaderProjection.Text });

        Assert.NotEmpty(lean.Tables);
        Assert.Equal(lean.GetText(), text.GetText());
        Assert.NotEmpty(text.Chunks);
        Assert.All(text.Chunks, chunk => {
            Assert.False(string.IsNullOrEmpty(chunk.Text));
            Assert.Null(chunk.Tables);
            Assert.Null(chunk.Markdown);
        });
        Assert.Empty(text.Blocks);
        Assert.Empty(text.Tables);
        Assert.Empty(text.Pages);
        Assert.Empty(text.Links);
        Assert.Empty(text.Metadata);
        Assert.Null(text.Markdown);
        Assert.Contains("officeimo.reader.projection.text", text.CapabilitiesUsed);
        Assert.Contains(StreamedCapability, text.CapabilitiesUsed);
        Assert.Equal(withTextBox, text.GetText().Contains("Boxed note"));
        Assert.DoesNotContain(StreamedCapability, lean.CapabilitiesUsed);
    }

    // Table shapes whose streamed text must still match the model: a grid wider than every row, a grid with no
    // rows, and a header row whose revised properties record it as an ordinary row.
    [Theory]
    [InlineData("<w:tblGrid><w:gridCol/><w:gridCol/><w:gridCol/></w:tblGrid>"
        + "<w:tr><w:tc><w:p><w:r><w:t>North</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>1,200</w:t></w:r></w:p></w:tc></w:tr>", true)]
    [InlineData("<w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid>", true)]
    [InlineData("<w:tblGrid><w:gridCol/></w:tblGrid>"
        + "<w:tr><w:trPr><w:tblHeader/><w:trPrChange w:id=\"1\" w:author=\"A\"><w:trPr><w:tblHeader w:val=\"false\"/></w:trPr></w:trPrChange></w:trPr>"
        + "<w:tc><w:p><w:r><w:t>Region</w:t></w:r></w:p></w:tc></w:tr>"
        + "<w:tr><w:tc><w:p><w:r><w:t>North</w:t></w:r></w:p></w:tc></w:tr>", false)]
    public void TextProjectionMatchesTheLeanReadForTableShapes(string tableContent, bool streams) {
        byte[] package = CreateTableDocument(tableContent);
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddWordHandler().Build();

        OfficeDocumentReadResult lean = reader.ReadDocument(package, "table.docx",
            new ReaderOptions { Projection = ReaderProjection.TextAndMetadata, MaxTableRows = 1 });
        OfficeDocumentReadResult text = reader.ReadDocument(package, "table.docx",
            new ReaderOptions { Projection = ReaderProjection.Text, MaxTableRows = 1 });

        Assert.Equal(lean.GetText(), text.GetText());
        Assert.Equal(streams, text.CapabilitiesUsed.Contains(StreamedCapability));
    }

    [Theory]
    [InlineData("<w:p><w:r><w:t>before<!-- note -->after</w:t></w:r></w:p>")]
    [InlineData("<w:p><w:r><w:t xml:space=\"preserve\">before <?pi x?>after</w:t></w:r></w:p>")]
    [InlineData("<w:p><w:r><w:t xml:space=\"preserve\">Outer </w:t></w:r><w:sdt><w:sdtContent><w:r><w:t>kept</w:t></w:r>"
        + "<w:sdt><w:sdtPr><w14:checkbox/></w:sdtPr><w:sdtContent><w:r><w:t>box</w:t></w:r></w:sdtContent></w:sdt></w:sdtContent></w:sdt></w:p>")]
    [InlineData("<w:p><w:sdt><w:sdtPr><w14:checkbox/></w:sdtPr><w:sdtContent><w:r><w:t>box</w:t></w:r></w:sdtContent></w:sdt><w:r><w:t>after</w:t></w:r></w:p>")]
    public void TextProjectionMatchesTheLeanReadForParagraphMarkup(string paragraph) {
        byte[] package = CreateDocumentWithParagraph(paragraph);
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddWordHandler().Build();

        OfficeDocumentReadResult lean = reader.ReadDocument(package, "markup.docx", new ReaderOptions { Projection = ReaderProjection.TextAndMetadata });
        OfficeDocumentReadResult text = reader.ReadDocument(package, "markup.docx", new ReaderOptions { Projection = ReaderProjection.Text });

        Assert.Equal(lean.GetText(), text.GetText());
        Assert.Contains(StreamedCapability, text.CapabilitiesUsed);
    }

    // Replaces the paragraph that holds "Marker" with raw markup, so comments and instructions stay where they are.
    private static byte[] CreateDocumentWithParagraph(string paragraph) {
        using var stream = new MemoryStream();
        using (WordDocument document = WordDocument.Create(stream)) {
            document.AddParagraph("Marker");
            document.Save();
        }
        using (var archive = new System.IO.Compression.ZipArchive(stream, System.IO.Compression.ZipArchiveMode.Update, leaveOpen: true)) {
            System.IO.Compression.ZipArchiveEntry entry = archive.GetEntry("word/document.xml")!;
            string xml;
            using (var readerStream = new StreamReader(entry.Open())) xml = readerStream.ReadToEnd();
            xml = System.Text.RegularExpressions.Regex.Replace(xml, "<w:p\\b(?:(?!</w:p>).)*?Marker.*?</w:p>", paragraph.Replace("$", "$$"),
                System.Text.RegularExpressions.RegexOptions.Singleline);
            if (!xml.Contains("xmlns:w14=", StringComparison.Ordinal)) {
                xml = xml.Replace("<w:document ", "<w:document xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" ", StringComparison.Ordinal);
            }
            entry.Delete();
            using var writer = new StreamWriter(archive.CreateEntry("word/document.xml").Open(), new System.Text.UTF8Encoding(false));
            writer.Write(xml);
        }
        return stream.ToArray();
    }

    // Headers, footers, fields, footnotes, revisions and inline content controls: the streamed pass reads them with
    // the model's rules, so the text matches the model read and the pass does not decline.
    [Fact]
    public void TextProjectionStreamsSectionsHeadersFieldsAndNotesLikeTheLeanRead() {
        byte[] package = CreateRichDocument();
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddWordHandler().Build();

        OfficeDocumentReadResult lean = reader.ReadDocument(package, "rich.docx",
            new ReaderOptions { Projection = ReaderProjection.TextAndMetadata });
        OfficeDocumentReadResult text = reader.ReadDocument(package, "rich.docx",
            new ReaderOptions { Projection = ReaderProjection.Text });

        Assert.Contains(StreamedCapability, text.CapabilitiesUsed);
        Assert.Equal(lean.GetText(), text.GetText());
        Assert.Contains("Running header", text.GetText());
        Assert.Contains("A footnote", text.GetText());
        Assert.DoesNotContain("PAGE", text.GetText());
        Assert.DoesNotContain("deleted", text.GetText());
        Assert.Contains("field result", text.GetText());
        Assert.DoesNotContain("DATE", text.GetText());
    }

    // A table in a later section, after a section with a header, is read once and in source order.
    [Fact]
    public void LeanReadKeepsTablesOfLaterSectionsInOrder() {
        byte[] package = CreateRichDocument();
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddWordHandler().Build();

        string text = reader.ReadDocument(package, "rich.docx",
            new ReaderOptions { Projection = ReaderProjection.TextAndMetadata }).GetText();

        int intro = text.IndexOf("Second section", StringComparison.Ordinal);
        int table = text.IndexOf("North\t1,200", StringComparison.Ordinal);
        Assert.True(intro >= 0 && table > intro, text);
        Assert.Equal(table, text.LastIndexOf("North\t1,200", StringComparison.Ordinal));
    }

    private static byte[] CreateRichDocument() {
        using var stream = new MemoryStream();
        using (WordDocument document = WordDocument.Create(stream)) {
            document.PageSetup.Header.AddParagraph("Running header");
            WordParagraph footer = document.PageSetup.Footer.AddParagraph("Page ");
            footer.AddField(WordFieldType.Page);
            WordParagraph body = document.AddParagraph("Body with a note");
            body.AddFootNote("A footnote");
            document.AddParagraph("Field: ").AddField(WordFieldType.Author);
            document.AddSection();
            document.AddParagraph("Second section");
            WordTable table = document.AddTable(2, 2);
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Region";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "Total";
            table.Rows[1].Cells[0].Paragraphs[0].Text = "North";
            table.Rows[1].Cells[1].Paragraphs[0].Text = "1,200";
            document.Save();
        }
        using (WordprocessingDocument package = WordprocessingDocument.Open(stream, true)) {
            W.Body body = package.MainDocumentPart!.Document!.Body!;
            W.Paragraph paragraph = body.Elements<W.Paragraph>().First();
            paragraph.Append(new W.DeletedRun(new W.Run(new W.DeletedText("deleted text"))) { Id = "1", Author = "A" });
            paragraph.Append(new W.SdtRun(new W.SdtContentRun(new W.Run(new W.Text(" in a control")))));
            // A content control holding a field: its runs follow the field state, so only the result shows.
            paragraph.Append(new W.SdtRun(new W.SdtContentRun(
                new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.Begin }),
                new W.Run(new W.FieldCode(" DATE ")),
                new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.Separate }),
                new W.Run(new W.Text(" field result")),
                new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.End }))));
        }
        return stream.ToArray();
    }

    private static byte[] CreateTableDocument(string tableContent) {
        using var stream = new MemoryStream();
        using (WordDocument document = WordDocument.Create(stream)) {
            document.AddParagraph("Before");
            document.Save();
        }
        using (WordprocessingDocument package = WordprocessingDocument.Open(stream, true)) {
            W.Body body = package.MainDocumentPart!.Document!.Body!;
            var table = new W.Table("<w:tbl xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" + tableContent + "</w:tbl>");
            body.InsertAfter(table, body.Elements<W.Paragraph>().First());
            body.InsertAfter(new W.Paragraph(new W.Run(new W.Text("After"))), table);
        }
        return stream.ToArray();
    }

    // Paragraphs, list items, a hyperlink, a line break and tab, and a table with a repeated header row.
    private static byte[] CreateDocument(bool withTextBox) {
        using var stream = new MemoryStream();
        using (WordDocument document = WordDocument.Create(stream)) {
            document.AddParagraph("Quarterly review");
            WordParagraph mixed = document.AddParagraph("Owner:");
            mixed.AddText("\tPat").AddBreak();
            mixed.AddText("second line");
            document.AddParagraph("See ").AddHyperLink("the dashboard", new Uri("https://example.com/"));
            WordList list = document.AddList(WordListStyle.Bulleted);
            list.AddItem("First item");
            list.AddItem("Second item");
            WordTable table = document.AddTable(3, 2);
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Region";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "Total";
            table.Rows[1].Cells[0].Paragraphs[0].Text = "North";
            table.Rows[1].Cells[1].Paragraphs[0].Text = "1,200";
            table.Rows[2].Cells[0].Paragraphs[0].Text = "South";
            table.RepeatHeaderRowAtTheTopOfEachPage = true;
            document.AddParagraph("Closing note");
            if (withTextBox) {
                document.AddTextBox("Boxed note");
            }
            document.Save();
        }
        return stream.ToArray();
    }
}
