using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void WordTemplate_DictionaryBindingHandlesSplitRunsNestedPathsAndDocumentStories() {
            string filePath = Path.Combine(_directoryWithFiles, "WordTemplateDictionary.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                Body body = document._document.MainDocumentPart!.Document.Body!;
                body.Append(new Paragraph(
                    new Run(new RunProperties(new Bold()), new Text("Hello {{Customer")),
                    new Run(new RunProperties(new Italic()), new Text(".Name}}!"))));
                document.AddHeadersAndFooters();
                document.Header.Default.AddParagraph("Order {{Order.Id}}");

                var values = new Dictionary<string, object?> {
                    ["Customer"] = new Dictionary<string, object?> { ["Name"] = "Ada" },
                    ["Order"] = new Dictionary<string, object?> { ["Id"] = 42 }
                };

                WordTemplateResult result = WordTemplate.Apply(document, values).EnsureComplete();

                Assert.Equal(2, result.PlaceholderCount);
                Assert.Equal(2, result.ReplacedPlaceholderCount);
                Assert.Contains("Hello Ada!", body.InnerText, StringComparison.Ordinal);
                Assert.Contains("Order 42", document.Header.Default.Paragraphs.Single().Text, StringComparison.Ordinal);
                Assert.True(body.Descendants<Run>().First().RunProperties?.Bold != null);
                Assert.True(body.Descendants<Run>().Skip(1).First().RunProperties?.Italic != null);
                document.Save();
            }

            using WordDocument reopened = WordDocument.Load(filePath);
            Assert.Contains("Hello Ada!", reopened._document.MainDocumentPart!.Document.Body!.InnerText, StringComparison.Ordinal);
            Assert.Empty(new OpenXmlValidator().Validate(reopened._wordprocessingDocument));
        }

        [Fact]
        public void WordTemplate_PocoBindingExpandsNestedBlocksAndEvaluatesPerItemConditions() {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph("{{#each Lines}}");
            document.AddParagraph("{{Product}}: {{Amount}}");
            document.AddParagraph("{{#Preferred}}");
            document.AddParagraph("Preferred customer price");
            document.AddParagraph("{{/Preferred}}");
            document.AddParagraph("{{#each Tags}}");
            document.AddParagraph("{{this}} for {{Product}}");
            document.AddParagraph("{{/each Tags}}");
            document.AddParagraph("{{/each Lines}}");

            var model = new InvoiceTemplateModel {
                Lines = new[] {
                    new InvoiceTemplateLine { Product = "Support", Amount = 25.5m, Preferred = true, Tags = new[] { "remote", "priority" } },
                    new InvoiceTemplateLine { Product = "Training", Amount = 80m, Preferred = false, Tags = new[] { "onsite" } }
                }
            };

            WordTemplateResult result = WordTemplate.Apply(document, model).EnsureComplete();
            string text = document._document.MainDocumentPart!.Document.Body!.InnerText;

            Assert.Equal(5, result.RepeatedBlockCount);
            Assert.Equal(2, result.ConditionalBlockCount);
            Assert.Equal(10, result.ReplacedPlaceholderCount);
            Assert.Contains("Support: 25.5", text, StringComparison.Ordinal);
            Assert.Contains("Training: 80", text, StringComparison.Ordinal);
            Assert.Contains("priority for Support", text, StringComparison.Ordinal);
            Assert.Contains("onsite for Training", text, StringComparison.Ordinal);
            Assert.Equal(1, document.Paragraphs.Count(paragraph => paragraph.Text == "Preferred customer price"));
            Assert.DoesNotContain("{{", text, StringComparison.Ordinal);
        }

        [Fact]
        public void WordTemplate_EmbedsImagesAndCreatesHyperlinksAtInlinePlaceholders() {
            string filePath = Path.Combine(_directoryWithFiles, "WordTemplateRichValues.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Logo {{Logo}} and {{Portal}}.");
                byte[] png = OfficeIMO.Drawing.OfficePngWriter.Encode(
                    new OfficeIMO.Drawing.OfficeRasterImage(1, 1, OfficeIMO.Drawing.OfficeColor.White));
                var values = new Dictionary<string, object?> {
                    ["Logo"] = new WordTemplateImage(png, "logo.png", width: 12, height: 12, description: "Company logo"),
                    ["Portal"] = new WordTemplateHyperlink("customer portal", new Uri("https://example.com/customer"))
                };

                WordTemplateResult result = WordTemplate.Apply(document, values).EnsureComplete();
                Paragraph paragraph = document._document.MainDocumentPart!.Document.Body!.Elements<Paragraph>().Single();

                Assert.Equal(2, result.ReplacedPlaceholderCount);
                Assert.Single(paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>());
                Assert.Equal("customer portal", Assert.Single(paragraph.Descendants<Hyperlink>()).InnerText);
                Assert.DoesNotContain("{{", paragraph.InnerText, StringComparison.Ordinal);
                document.Save();
            }

            using WordDocument reopened = WordDocument.Load(filePath);
            Assert.Single(reopened.Images);
            Assert.Equal(new Uri("https://example.com/customer"), Assert.Single(reopened.HyperLinks).Uri);
            Assert.Empty(new OpenXmlValidator().Validate(reopened._wordprocessingDocument));
        }

        [Fact]
        public void WordTemplate_PreservesMissingPlaceholdersAndReportsThemByDefault() {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph("Hello {{Name}} from {{City}}");

            WordTemplateResult result = WordTemplate.Apply(
                document,
                new Dictionary<string, object?> { ["Name"] = "Ada" });

            Assert.False(result.IsComplete);
            Assert.Equal(new[] { "City" }, result.MissingValueNames);
            Assert.Equal("Hello Ada from {{City}}", document.Paragraphs.Single().Text);
            Assert.Throws<InvalidOperationException>(() => result.EnsureComplete());
        }

        [Fact]
        public void WordTemplate_RejectsRichPlaceholdersNestedInsideExistingHyperlinks() {
            using WordDocument document = WordDocument.Create();
            WordParagraph paragraph = document.AddParagraph("Visit ");
            paragraph.AddHyperLink("{{Portal}}", new Uri("https://example.com/old"));

            var values = new Dictionary<string, object?> {
                ["Portal"] = new WordTemplateHyperlink("new portal", new Uri("https://example.com/new"))
            };

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
                () => WordTemplate.Apply(document, values));

            Assert.Contains("direct paragraph text runs", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void WordTemplate_IDictionaryBindingExpandsBlocksExposedByIncludedCondition() {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph("{{#Include}}");
            document.AddParagraph("{{#each Items}}");
            document.AddParagraph("Item {{this}}");
            document.AddParagraph("{{/each Items}}");
            document.AddParagraph("{{/Include}}");
            IDictionary<string, object?> values = new Dictionary<string, object?> {
                ["Include"] = true,
                ["Items"] = new[] { "one", "two" }
            };

            WordTemplateResult result = WordTemplate.Apply(document, values).EnsureComplete();

            Assert.Equal(2, result.RepeatedBlockCount);
            Assert.Equal(1, result.ConditionalBlockCount);
            Assert.Equal(new[] { "Item one", "Item two" }, document.Paragraphs.Select(paragraph => paragraph.Text));
        }

        [Fact]
        public void WordTemplate_RepeatingAuthoredContentReassignsDocumentUniqueIdentifiers() {
            string filePath = Path.Combine(_directoryWithFiles, "WordTemplateRepeatedIdentifiers.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("{{#each Items}}");
                WordParagraph repeated = document.AddParagraph("{{Name}}");
                repeated._paragraph.ParagraphId = "00000001";
                repeated._paragraph.TextId = "00000002";
                var bookmarkStart = new BookmarkStart { Id = "5", Name = "Target" };
                ParagraphProperties? paragraphProperties = repeated._paragraph.GetFirstChild<ParagraphProperties>();
                if (paragraphProperties != null) repeated._paragraph.InsertAfter(bookmarkStart, paragraphProperties);
                else repeated._paragraph.PrependChild(bookmarkStart);
                repeated._paragraph.Append(new BookmarkEnd { Id = "5" });
                byte[] png = OfficeIMO.Drawing.OfficePngWriter.Encode(
                    new OfficeIMO.Drawing.OfficeRasterImage(1, 1, OfficeIMO.Drawing.OfficeColor.White));
                using (var image = new MemoryStream(png, writable: false)) {
                    repeated.AddImage(image, "marker.png", 8, 8);
                }
                document.AddParagraph("{{/each Items}}");
                WordParagraph externalLink = document.AddParagraph("Jump to ");
                externalLink._paragraph.Append(new Hyperlink(
                    new Run(new Text("target"))) { Anchor = "Target" });

                WordTemplate.Apply(document, new Dictionary<string, object?> {
                    ["Items"] = new object[] {
                        new Dictionary<string, object?> { ["Name"] = "One" },
                        new Dictionary<string, object?> { ["Name"] = "Two" }
                    }
                }).EnsureComplete();

                Body body = document._document.MainDocumentPart!.Document.Body!;
                Assert.Equal(2, body.Descendants<BookmarkStart>().Select(bookmark => bookmark.Id!.Value).Distinct().Count());
                Assert.Equal(2, body.Descendants<BookmarkStart>().Select(bookmark => bookmark.Name!.Value).Distinct(StringComparer.OrdinalIgnoreCase).Count());
                Assert.Contains(body.Descendants<BookmarkStart>(), bookmark => bookmark.Name?.Value == "Target");
                Assert.Equal("Target", body.Descendants<Hyperlink>().Single().Anchor?.Value);
                Assert.Equal(4, body.Descendants<DW.DocProperties>().Concat<OpenXmlElement>(
                    body.Descendants<DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties>())
                    .Select(element => element switch {
                        DW.DocProperties properties => properties.Id!.Value,
                        DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties properties => properties.Id!.Value,
                        _ => 0U
                    }).Distinct().Count());
                Assert.Equal(2, body.Descendants<Paragraph>().Select(paragraph => paragraph.ParagraphId?.Value).Where(value => value != null).Distinct().Count());
                document.Save();
            }

            using WordDocument reopened = WordDocument.Load(filePath);
            Assert.Empty(new OpenXmlValidator().Validate(reopened._wordprocessingDocument));
        }

        [Fact]
        public void WordTemplate_BindsNestedBlocksAndRichValuesInsideTextBoxes() {
            using WordDocument document = WordDocument.Create();
            WordTextBox content = document.AddTextBox("{{#Include}}");
            WordParagraph inside = content.Paragraphs[0].AddParagraph("Inside {{Name}}");
            inside.AddParagraph("{{/Include}}");
            WordTextBox link = document.AddTextBox("{{Portal}}");

            WordTemplate.Apply(document, new Dictionary<string, object?> {
                ["Include"] = true,
                ["Name"] = "Ada",
                ["Portal"] = new WordTemplateHyperlink("customer portal", new Uri("https://example.com/customer"))
            }).EnsureComplete();

            Assert.Contains(content.Paragraphs, paragraph => paragraph.Text == "Inside Ada");
            Assert.DoesNotContain("{{", string.Concat(content.Paragraphs.Select(paragraph => paragraph.Text)), StringComparison.Ordinal);
            Assert.Equal("customer portal", Assert.Single(link.Paragraphs
                .SelectMany(paragraph => paragraph._paragraph.Descendants<Hyperlink>())
                .Distinct()).InnerText);
            Assert.Empty(new OpenXmlValidator().Validate(document._wordprocessingDocument));
        }

        private sealed class InvoiceTemplateModel {
            public IReadOnlyList<InvoiceTemplateLine> Lines { get; set; } = Array.Empty<InvoiceTemplateLine>();
        }

        private sealed class InvoiceTemplateLine {
            public string Product { get; set; } = string.Empty;
            public decimal Amount { get; set; }
            public bool Preferred { get; set; }
            public IReadOnlyList<string> Tags { get; set; } = Array.Empty<string>();
        }
    }
}
