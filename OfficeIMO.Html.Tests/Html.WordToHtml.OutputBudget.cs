using OfficeIMO.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.IO;
using Xunit;
using M = DocumentFormat.OpenXml.Math;

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
        public void Test_WordToHtml_OutputBudgetReservesGeneratedRunAttributesBeforeDomAssignment() {
            using WordDocument document = WordDocument.Create();
            WordParagraph paragraph = document.AddParagraph();
            for (int index = 0; index < 64; index++) {
                paragraph.AddText("x").SetFontSize(12);
            }
            string expected = document.ToHtmlResult(new WordToHtmlOptions {
                IncludeDefaultCss = false
            }).RequireValue();

            HtmlConversionLimitException exception = Assert.Throws<HtmlConversionLimitException>(() =>
                document.ToHtmlResult(new WordToHtmlOptions {
                    IncludeDefaultCss = false,
                    MaxOutputCharacters = expected.Length - 1
                }));

            Assert.Equal("WordHtmlOutputLimitExceeded", exception.Code);
            Assert.Equal("RunFormatting:font-size", exception.LimitSource);
            Assert.True(exception.Actual > exception.Limit);
        }

        [Fact]
        public void Test_WordToHtml_OutputBudgetExcludesOpenXmlOnlyTextAttributes() {
            using WordDocument document = WordDocument.Create();
            for (int index = 0; index < 256; index++) {
                document.AddParagraph()._paragraph.Append(
                    new Run(new Text(" value ") { Space = SpaceProcessingModeValues.Preserve }));
            }
            var unboundedOptions = new WordToHtmlOptions { IncludeDefaultCss = false };
            string expected = document.ToHtmlResult(unboundedOptions).RequireValue();

            HtmlTextConversionResult bounded = document.ToHtmlResult(new WordToHtmlOptions {
                IncludeDefaultCss = false,
                MaxOutputCharacters = expected.Length
            });

            Assert.True(bounded.Succeeded);
            Assert.Equal(expected, bounded.RequireValue());
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void Test_WordToHtml_OutputBudgetReservesCollapsedCodeSeparatorsBeforeMaterialization(bool insideTable) {
            using WordDocument document = WordDocument.Create();
            if (insideTable) {
                WordTableCell cell = document.AddTable(1, 1).Rows[0].Cells[0];
                for (int index = 0; index < 8192; index++) {
                    cell.AddParagraph().SetStyleId("HTMLPreformatted");
                }
            } else {
                for (int index = 0; index < 8192; index++) {
                    document.AddParagraph().SetStyleId("HTMLPreformatted");
                }
            }

            HtmlConversionLimitException exception = Assert.Throws<HtmlConversionLimitException>(() =>
                document.ToHtmlResult(new WordToHtmlOptions {
                    IncludeDefaultCss = false,
                    IncludeParagraphClasses = false,
                    MaxOutputCharacters = 4096
                }));

            Assert.Equal("WordHtmlOutputLimitExceeded", exception.Code);
            Assert.Equal("CodeBlockSeparators", exception.LimitSource);
            Assert.True(exception.Actual > exception.Limit);
        }

        [Fact]
        public void Test_WordToHtml_OutputBudgetReservesEncodedDocumentMetadataBeforeDomConstruction() {
            using WordDocument document = WordDocument.Create();
            document.BuiltinDocumentProperties.Title = new string('&', 1024);
            document.AddParagraph("Visible content");

            HtmlConversionLimitException exception = Assert.Throws<HtmlConversionLimitException>(() =>
                document.ToHtmlResult(new WordToHtmlOptions {
                    IncludeDefaultCss = false,
                    MaxOutputCharacters = 4500
                }));

            Assert.Equal("WordHtmlOutputLimitExceeded", exception.Code);
            Assert.Equal("DocumentMetadata:title", exception.LimitSource);
            Assert.True(exception.Actual > exception.Limit);
        }

        [Fact]
        public void Test_WordToHtml_OutputBudgetReservesCompleteAttributeSyntaxBeforeDomAssignment() {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph("Visible content");
            var options = new WordToHtmlOptions {
                IncludeDefaultCss = false,
                MaxOutputCharacters = 2500
            };
            options.AdditionalMetaTags.Add((new string('x', 4096), string.Empty));

            HtmlConversionLimitException exception = Assert.Throws<HtmlConversionLimitException>(() =>
                document.ToHtmlResult(options));

            Assert.Equal("WordHtmlOutputLimitExceeded", exception.Code);
            Assert.Equal("AdditionalMeta:name", exception.LimitSource);
            Assert.True(exception.Actual > exception.Limit);
        }

        [Fact]
        public void Test_WordToHtml_OutputBudgetReservesSectionMetadataBeforeDomAssignment() {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph("First section");
            for (int index = 1; index < 32; index++) {
                document.AddSection(SectionMarkValues.NextPage).AddParagraph("Section " + index);
            }
            foreach (WordSection section in document.Sections) {
                section.PageOrientation = PageOrientationValues.Landscape;
                section.PageSettings.PageSize = WordPageSize.Letter;
                section.Margins.Top = 1440;
                section.Margins.Right = 1200;
                section.Margins.Bottom = 720;
                section.Margins.Left = 1080;
            }

            string withoutMetadata = document.ToHtmlResult(new WordToHtmlOptions {
                IncludeDefaultCss = false
            }).RequireValue();
            string withMetadata = document.ToHtmlResult(new WordToHtmlOptions {
                IncludeDefaultCss = false,
                IncludeSectionMetadata = true
            }).RequireValue();
            long limit = withoutMetadata.Length + ((withMetadata.Length - withoutMetadata.Length) / 2L);

            HtmlConversionLimitException exception = Assert.Throws<HtmlConversionLimitException>(() =>
                document.ToHtmlResult(new WordToHtmlOptions {
                    IncludeDefaultCss = false,
                    IncludeSectionMetadata = true,
                    MaxOutputCharacters = limit
                }));

            Assert.Equal("WordHtmlOutputLimitExceeded", exception.Code);
            Assert.StartsWith("SectionMetadata:", exception.LimitSource, StringComparison.Ordinal);
            Assert.True(exception.Actual > exception.Limit);
        }

        [Fact]
        public void Test_WordToHtml_OutputBudgetReservesGeneratedStyleCssBeforeDomAssignment() {
            using WordDocument document = WordDocument.Create();
            const string styleId = "BudgetedStyle";
            document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(
                new Style(
                    new StyleRunProperties(
                        new RunFonts { Ascii = new string('A', 8192) })) {
                    Type = StyleValues.Paragraph,
                    StyleId = styleId
                });
            document.AddParagraph("Styled content").SetStyleId(styleId);

            HtmlConversionLimitException exception = Assert.Throws<HtmlConversionLimitException>(() =>
                document.ToHtmlResult(new WordToHtmlOptions {
                    IncludeDefaultCss = false,
                    IncludeParagraphClasses = true,
                    MaxOutputCharacters = 4096
                }));

            Assert.Equal("WordHtmlOutputLimitExceeded", exception.Code);
            Assert.Equal("GeneratedStyleCss", exception.LimitSource);
            Assert.True(exception.Actual > exception.Limit);
        }

        [Fact]
        public void Test_WordToHtml_OutputBudgetReservesRelationshipBackedHyperlinkBeforeDomAssignment() {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph().AddHyperLink(
                "bounded link",
                new Uri("https://example.test/" + new string('a', 8192), UriKind.Absolute));

            HtmlConversionLimitException exception = Assert.Throws<HtmlConversionLimitException>(() =>
                document.ToHtmlResult(new WordToHtmlOptions {
                    IncludeDefaultCss = false,
                    MaxOutputCharacters = 4096
                }));

            Assert.Equal("WordHtmlOutputLimitExceeded", exception.Code);
            Assert.Equal("Hyperlink:href", exception.LimitSource);
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
            Assert.StartsWith("DropDownOption:", exception.LimitSource, StringComparison.Ordinal);
            Assert.True(exception.Actual > exception.Limit);
        }

        [Fact]
        public void Test_WordToHtml_OutputBudgetReservesDistinctDropDownValueAndDisplayText() {
            using WordDocument document = WordDocument.Create();
            WordDropDownList dropDown = document.AddParagraph().AddDropDownList(new[] { "placeholder" });
            ListItem item = dropDown._sdtRun.SdtProperties!
                .GetFirstChild<SdtContentDropDownList>()!
                .Elements<ListItem>()
                .Single();
            item.Value = "short";
            item.DisplayText = new string('&', 1024);

            HtmlConversionLimitException exception = Assert.Throws<HtmlConversionLimitException>(() =>
                document.ToHtmlResult(new WordToHtmlOptions {
                    IncludeDefaultCss = false,
                    MaxOutputCharacters = 4096
                }));

            Assert.Equal("WordHtmlOutputLimitExceeded", exception.Code);
            Assert.Equal("DropDownOption:display-text", exception.LimitSource);
            Assert.True(exception.Actual > exception.Limit);
        }

        [Fact]
        public void Test_WordToHtml_ComboBoxPreservesDistinctValueAndDisplayText() {
            using WordDocument document = WordDocument.Create();
            WordComboBox comboBox = document.AddParagraph().AddComboBox(new[] { "placeholder" });
            ListItem item = comboBox._sdtRun.SdtProperties!
                .GetFirstChild<SdtContentComboBox>()!
                .Elements<ListItem>()
                .Single();
            item.Value = "internal-id";
            item.DisplayText = "Visible label";

            string html = document.ToHtmlResult(new WordToHtmlOptions {
                IncludeDefaultCss = false
            }).RequireValue();

            Assert.Contains("value=\"internal-id\"", html, StringComparison.Ordinal);
            Assert.Contains("label=\"Visible label\"", html, StringComparison.Ordinal);
        }

        [Fact]
        public void Test_WordToHtml_OutputBudgetReservesComboBoxDisplayText() {
            using WordDocument document = WordDocument.Create();
            WordComboBox comboBox = document.AddParagraph().AddComboBox(new[] { "placeholder" });
            ListItem item = comboBox._sdtRun.SdtProperties!
                .GetFirstChild<SdtContentComboBox>()!
                .Elements<ListItem>()
                .Single();
            item.Value = "short";
            item.DisplayText = new string('&', 1024);

            HtmlConversionLimitException exception = Assert.Throws<HtmlConversionLimitException>(() =>
                document.ToHtmlResult(new WordToHtmlOptions {
                    IncludeDefaultCss = false,
                    MaxOutputCharacters = 4096
                }));

            Assert.Equal("WordHtmlOutputLimitExceeded", exception.Code);
            Assert.Equal("ComboBoxOption:label", exception.LimitSource);
            Assert.True(exception.Actual > exception.Limit);
        }

        [Theory]
        [InlineData("png")]
        [InlineData("svg")]
        public void Test_WordToHtml_OutputBudgetReservesExternalImageSource(string extension) {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph().AddImage(
                new Uri("https://example.test/" + new string('a', 8192) + "/image." + extension),
                10,
                10);

            HtmlConversionLimitException exception = Assert.Throws<HtmlConversionLimitException>(() =>
                document.ToHtmlResult(new WordToHtmlOptions {
                    EmbedImagesAsBase64 = false,
                    IncludeDefaultCss = false,
                    MaxOutputCharacters = 6_000
                }));

            Assert.Equal("WordHtmlOutputLimitExceeded", exception.Code);
            Assert.Equal("Image:src", exception.LimitSource);
            Assert.True(exception.Actual > exception.Limit);
        }

        [Fact]
        public void Test_WordToHtml_OutputBudgetReservesRepeatedCommentReferenceMetadata() {
            using WordDocument document = WordDocument.Create();
            WordParagraph paragraph = document.AddParagraph("Reviewed text");
            paragraph.AddComment(new string('&', 256), "AB", "Review note");
            WordComment comment = Assert.Single(document.Comments);
            for (int index = 0; index < 8; index++) {
                paragraph._paragraph.Append(new Run(new CommentReference { Id = comment.Id }));
            }

            HtmlConversionLimitException exception = Assert.Throws<HtmlConversionLimitException>(() =>
                document.ToHtmlResult(new WordToHtmlOptions {
                    ExportComments = true,
                    IncludeDefaultCss = false,
                    MaxOutputCharacters = 4096
                }));

            Assert.Equal("WordHtmlOutputLimitExceeded", exception.Code);
            Assert.Equal("CommentReference:title", exception.LimitSource);
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
                MaxOutputCharacters = expected.Length + 16
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
        public void Test_WordToHtml_OutputBudgetCombinesRepeatedCallerFontStyles() {
            using WordDocument document = WordDocument.Create();
            WordParagraph paragraph = document.AddParagraph();
            for (int index = 0; index < 12; index++) {
                paragraph._paragraph.Append(new Run(
                    new RunProperties(new Bold()),
                    new Text("x")));
            }
            var options = new WordToHtmlOptions {
                FontFamily = new string('F', 512),
                IncludeDefaultCss = false,
                IncludeFontStyles = true,
                MaxOutputCharacters = 4096
            };

            HtmlConversionLimitException exception = Assert.Throws<HtmlConversionLimitException>(
                () => document.ToHtmlResult(options));

            Assert.Equal("WordHtmlOutputLimitExceeded", exception.Code);
            Assert.Equal("RunFontStyle", exception.LimitSource);
            Assert.True(exception.Actual > exception.Limit);
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

        [Fact]
        public void Test_WordToHtml_OutputBudgetCountsSharedHeaderSourceAttributesPerSection() {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph("Body");
            WordSection firstSection = document.Sections[0];
            WordParagraph headerParagraph = firstSection
                .GetOrCreateHeader(HeaderFooterValues.Default)
                .AddParagraph();
            headerParagraph._paragraph.Append(new Run(
                new RunProperties(new RunFonts { Ascii = new string('F', 2048) }),
                new Text("Header")));
            string relationshipId = firstSection._sectionProperties
                .GetFirstChild<HeaderReference>()!.Id!;
            for (int index = 0; index < 2; index++) {
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
                    IncludeFontStyles = true,
                    MaxOutputCharacters = 5_000
                }));

            Assert.Equal("WordHtmlOutputLimitExceeded", exception.Code);
            Assert.StartsWith("HeaderFooter:header:default:section-", exception.LimitSource, StringComparison.Ordinal);
            Assert.True(exception.Actual > exception.Limit);
        }

        [Fact]
        public void Test_WordToHtml_OutputBudgetBoundsMathMlBeforeFragmentParsing() {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph()._paragraph.Append(
                new M.OfficeMath(new M.Run(new M.Text(new string('&', 4096)))));

            HtmlConversionLimitException exception = Assert.Throws<HtmlConversionLimitException>(() =>
                document.ToHtmlResult(new WordToHtmlOptions {
                    IncludeDefaultCss = false,
                    MaxOutputCharacters = 4096
                }));

            Assert.Equal("WordHtmlOutputLimitExceeded", exception.Code);
            Assert.Equal("EquationMathMl", exception.LimitSource);
            Assert.True(exception.Actual > exception.Limit);
        }

        [Fact]
        public void Test_WordToHtml_OutputBudgetDoesNotRetainReplacedEquationFieldText() {
            using WordDocument document = WordDocument.Create();
            string cachedResult = new string('x', 4096);
            document.AddParagraph()._paragraph.Append(
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode(" EQ \\f(a,b) ")),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new Run(new Text(cachedResult)),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
            string expected = document.ToHtmlResult(new WordToHtmlOptions {
                IncludeDefaultCss = false
            }).RequireValue();

            HtmlTextConversionResult bounded = document.ToHtmlResult(new WordToHtmlOptions {
                IncludeDefaultCss = false,
                MaxOutputCharacters = expected.Length + 512
            });

            Assert.True(bounded.Succeeded);
            Assert.Equal(expected, bounded.RequireValue());
            Assert.Equal(2, bounded.RequireValue().Split(
                new[] { cachedResult },
                StringSplitOptions.None).Length - 1);
        }

        [Fact]
        public void Test_WordToHtml_OutputBudgetReservesRepeatedFootnoteReferenceMetadata() {
            using WordDocument document = WordDocument.Create();
            WordParagraph paragraph = document.AddParagraph("Notes");
            WordParagraph reference = paragraph.AddFootNote("shared note");
            for (int index = 0; index < 256; index++) {
                paragraph._paragraph.Append(reference._run!.CloneNode(true));
            }

            HtmlConversionLimitException exception = Assert.Throws<HtmlConversionLimitException>(() =>
                document.ToHtmlResult(new WordToHtmlOptions {
                    ExportFootnotes = true,
                    IncludeDefaultCss = false,
                    MaxOutputCharacters = 10_000
                }));

            Assert.Equal("WordHtmlOutputLimitExceeded", exception.Code);
            Assert.StartsWith("FootnoteReference:", exception.LimitSource, StringComparison.Ordinal);
            Assert.True(exception.Actual > exception.Limit);
        }
    }
}
