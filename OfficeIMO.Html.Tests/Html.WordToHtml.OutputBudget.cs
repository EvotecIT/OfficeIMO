using OfficeIMO.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.IO;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class HtmlWordToHtml {
        [Fact]
        public void Test_WordToHtml_OutputBudgetExcludesOmittedRevisionText() {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph("Visible content").AddInsertedText(new string('x', 16_384), "Reviewer");
            var options = new WordToHtmlOptions { MaxOutputCharacters = 4096 };

            HtmlTextConversionResult result = document.ToHtmlResult(options);

            Assert.True(result.Succeeded);
            Assert.Contains("Visible content", result.RequireValue(), StringComparison.Ordinal);
            Assert.DoesNotContain(new string('x', 256), result.RequireValue(), StringComparison.Ordinal);
            Assert.Contains(result.Report.Diagnostics, diagnostic =>
                diagnostic.Code == "TrackedRevisionTextOmitted" &&
                diagnostic.LossKind == HtmlConversionLossKind.Omission);
        }

        [Fact]
        public void Test_WordToHtml_OutputBudgetStopsEmptyParagraphsBeforeDomConstruction() {
            using WordDocument document = WordDocument.Create();
            for (int index = 0; index < 1024; index++) {
                document.AddParagraph();
            }
            var options = new WordToHtmlOptions {
                IncludeDefaultCss = false,
                MaxOutputCharacters = 512
            };

            HtmlConversionLimitException exception = Assert.Throws<HtmlConversionLimitException>(
                () => document.ToHtmlResult(options));

            Assert.Equal("WordHtmlOutputLimitExceeded", exception.Code);
            Assert.Equal("GeneratedElement:p", exception.LimitSource);
            Assert.True(exception.Actual > exception.Limit);
        }

        [Fact]
        public void Test_WordToHtml_OutputBudgetCombinesContentAndElementConstruction() {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph(new string('x', 512));
            for (int index = 0; index < 500; index++) {
                document.AddParagraph();
            }
            var options = new WordToHtmlOptions {
                IncludeDefaultCss = false,
                MaxOutputCharacters = 4096
            };

            HtmlConversionLimitException exception = Assert.Throws<HtmlConversionLimitException>(
                () => document.ToHtmlResult(options));

            Assert.Equal("WordHtmlOutputLimitExceeded", exception.Code);
            Assert.StartsWith("GeneratedElement:", exception.LimitSource, StringComparison.Ordinal);
            Assert.True(exception.Actual > exception.Limit);
        }

        [Fact]
        public void Test_WordToHtml_OutputBudgetCountsHtmlEscapingBeforeDomConstruction() {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph(new string('&', 1024));
            var options = new WordToHtmlOptions {
                IncludeDefaultCss = false,
                MaxOutputCharacters = 4500
            };

            HtmlConversionLimitException exception = Assert.Throws<HtmlConversionLimitException>(
                () => document.ToHtmlResult(options));

            Assert.Equal("WordHtmlOutputLimitExceeded", exception.Code);
            Assert.Contains("document.xml", exception.LimitSource, StringComparison.OrdinalIgnoreCase);
            Assert.True(exception.Actual > exception.Limit);
        }

        [Fact]
        public void Test_WordToHtml_OutputBudgetCombinesEarlyImageAndLaterElements() {
            using WordDocument document = WordDocument.Create();
            string assetPath = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "OfficeIMO.png");
            document.AddParagraph().AddImage(assetPath, 20, 20);
            for (int index = 0; index < 800; index++) {
                document.AddParagraph();
            }
            long imageBytes = new FileInfo(assetPath).Length;
            long imageDataUriCharacters = "data:image/png;base64,".Length + ((imageBytes + 2L) / 3L) * 4L;
            var options = new WordToHtmlOptions {
                IncludeDefaultCss = false,
                MaxOutputCharacters = imageDataUriCharacters + 4096L
            };

            HtmlConversionLimitException exception = Assert.Throws<HtmlConversionLimitException>(
                () => document.ToHtmlResult(options));

            Assert.Equal("WordHtmlOutputLimitExceeded", exception.Code);
            Assert.StartsWith("GeneratedElement:", exception.LimitSource, StringComparison.Ordinal);
            Assert.True(exception.Actual > exception.Limit);
        }

        [Fact]
        public void Test_WordToHtml_OutputBudgetStopsEmptyDropDownOptionsBeforeDomConstruction() {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph().AddDropDownList(
                Enumerable.Range(0, 256).Select(index => "Item" + index).ToArray());
            var options = new WordToHtmlOptions {
                IncludeDefaultCss = false,
                MaxOutputCharacters = 4096
            };

            HtmlConversionLimitException exception = Assert.Throws<HtmlConversionLimitException>(
                () => document.ToHtmlResult(options));

            Assert.Equal("WordHtmlOutputLimitExceeded", exception.Code);
            Assert.Equal("GeneratedElement:option", exception.LimitSource);
            Assert.True(exception.Actual > exception.Limit);
        }

        [Fact]
        public void Test_WordToHtml_OutputBudgetUsesSerializedVoidElementSize() {
            using WordDocument document = WordDocument.Create();
            WordParagraph paragraph = document.AddParagraph();
            for (int index = 0; index < 200; index++) paragraph.AddBreak();
            var unboundedOptions = new WordToHtmlOptions { IncludeDefaultCss = false };
            string expected = document.ToHtmlResult(unboundedOptions).RequireValue();

            HtmlTextConversionResult bounded = document.ToHtmlResult(new WordToHtmlOptions {
                IncludeDefaultCss = false,
                MaxOutputCharacters = expected.Length
            });

            Assert.True(bounded.Succeeded);
            Assert.Equal(expected, bounded.RequireValue());
        }

        [Fact]
        public void Test_WordToHtml_OutputBudgetExcludesFieldInstructionsThatAreNotRendered() {
            using WordDocument document = WordDocument.Create();
            string largeInstruction = " QUOTE \"" + new string('x', 8192) + "\" ";
            document.AddParagraph()._paragraph.Append(
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode(largeInstruction) { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new Run(new Text("Complex cached result")),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
            document.AddParagraph()._paragraph.Append(
                new SimpleField(new Run(new Text("Simple cached result"))) { Instruction = largeInstruction });

            HtmlTextConversionResult result = document.ToHtmlResult(new WordToHtmlOptions {
                IncludeDefaultCss = false,
                MaxOutputCharacters = 4096
            });

            Assert.True(result.Succeeded);
            Assert.Contains("Complex cached result", result.RequireValue(), StringComparison.Ordinal);
            Assert.DoesNotContain(new string('x', 128), result.RequireValue(), StringComparison.Ordinal);
        }

        [Fact]
        public void Test_WordToHtml_OutputBudgetExcludesDisabledRunFontAttributes() {
            using WordDocument document = WordDocument.Create();
            WordParagraph paragraph = document.AddParagraph();
            paragraph._paragraph.Append(new Run(
                new RunProperties(new RunFonts { Ascii = new string('x', 8192) }),
                new Text("Visible content")));

            HtmlTextConversionResult result = document.ToHtmlResult(new WordToHtmlOptions {
                IncludeDefaultCss = false,
                IncludeFontStyles = false,
                MaxOutputCharacters = 4096
            });

            Assert.True(result.Succeeded);
            Assert.Contains("Visible content", result.RequireValue(), StringComparison.Ordinal);
            Assert.DoesNotContain("font-family", result.RequireValue(), StringComparison.Ordinal);
        }

        [Fact]
        public void Test_WordToHtml_OutputBudgetCountsSharedHeaderForEveryExportedSection() {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph("Body");
            WordSection firstSection = document.Sections[0];
            firstSection.GetOrCreateHeader(HeaderFooterValues.Default)
                .AddParagraph(new string('H', 1024));
            string relationshipId = firstSection._sectionProperties
                .GetFirstChild<HeaderReference>()!.Id!;
            for (int index = 0; index < 8; index++) {
                WordSection section = document.AddSection(SectionMarkValues.NextPage);
                section._sectionProperties.InsertAt(new HeaderReference {
                    Type = HeaderFooterValues.Default,
                    Id = relationshipId
                }, 0);
                section.Header.Default = firstSection.Header.Default;
                section.AddParagraph("Section " + index);
            }

            HtmlConversionLimitException exception = Assert.Throws<HtmlConversionLimitException>(() =>
                document.ToHtmlResult(new WordToHtmlOptions {
                    ExportHeadersAndFooters = true,
                    IncludeDefaultCss = false,
                    MaxOutputCharacters = 5_000
                }));

            Assert.Equal("WordHtmlOutputLimitExceeded", exception.Code);
            Assert.StartsWith("HeaderFooter:header:default:section-", exception.LimitSource, StringComparison.Ordinal);
            Assert.True(exception.Actual > exception.Limit);
        }
    }
}
