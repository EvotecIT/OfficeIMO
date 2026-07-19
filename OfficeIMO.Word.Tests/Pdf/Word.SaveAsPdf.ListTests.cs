using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Pdf;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Renders_Simple_Lists_With_Native_List_Blocks() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeLists.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeLists.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                WordList bulletList = document.AddList(WordListStyle.Bulleted);
                bulletList.AddItem("Native bullet one");
                bulletList.AddItem("Native bullet two");

                WordList numberedList = document.AddCustomList();
                numberedList.Numbering.AddLevel(new WordListLevel(WordListLevelKind.DecimalDot).SetStartNumberingValue(3));
                numberedList.AddItem("Native step three");
                numberedList.AddItem("Native step four");

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            Assert.True(File.Exists(pdfPath));
            byte[] bytes = File.ReadAllBytes(pdfPath);
            var listItems = PdfTextExtractor.ExtractListItemsByPage(bytes)
                .SelectMany(page => page.ListItems)
                .ToList();

            Assert.Contains(listItems, item => item.Text == "Native bullet one");
            Assert.Contains(listItems, item => item.Text == "Native bullet two");
            Assert.Contains(listItems, item => item.Marker == "3" && item.Text == "Native step three");
            Assert.Contains(listItems, item => item.Marker == "4" && item.Text == "Native step four");
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Uses_Word_List_Hanging_Indent_In_Native_List_Blocks() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeListHangingIndent.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeListHangingIndent.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                WordList bulletList = document.AddList(WordListStyle.Bulleted);
                WordListLevel level = bulletList.Numbering.Levels[0];
                level.IndentationLeft = 720;
                level.IndentationHanging = 360;
                bulletList.AddItem("Wrapped native bullet item with enough body text to flow onto a second line so the generated PDF can prove the continuation aligns with the Word text position.");

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
            var lineGroups = pdf.GetPage(1).Letters
                .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
                .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
                .OrderByDescending(group => group.Key)
                .Select(group => group.OrderBy(letter => letter.StartBaseLine.X).ToList())
                .ToList();

            int bulletLineIndex = lineGroups.FindIndex(line => line.Any(letter => letter.Value == "•"));
            Assert.True(bulletLineIndex >= 0, "Expected a native bullet marker in the generated PDF.");

            var bulletLine = lineGroups[bulletLineIndex];
            double bulletX = bulletLine.First(letter => letter.Value == "•").StartBaseLine.X;
            double textX = bulletLine.First(letter => letter.Value == "W").StartBaseLine.X;
            Assert.InRange(textX - bulletX, 14D, 24D);

            var continuationLine = lineGroups
                .Skip(bulletLineIndex + 1)
                .First(line => !line.Any(letter => letter.Value == "•"));
            double continuationX = continuationLine[0].StartBaseLine.X;
            Assert.InRange(continuationX, textX - 1D, textX + 1D);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Honors_Direct_List_Paragraph_Indentation() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeListDirectIndent.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeListDirectIndent.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                WordList bulletList = document.AddList(WordListStyle.Bulleted);
                WordListLevel level = bulletList.Numbering.Levels[0];
                level.IndentationLeft = 720;
                level.IndentationHanging = 360;

                bulletList.AddItem("RegularIndentMarker");
                WordParagraph wideIndent = bulletList.AddItem("WideIndentMarker");
                wideIndent.IndentationBeforePoints = 72;
                wideIndent.IndentationHangingPoints = 36;

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
            var words = pdf.GetPage(1).GetWords().ToList();
            double regularX = Assert.Single(words, word => word.Text == "RegularIndentMarker").BoundingBox.Left;
            double wideX = Assert.Single(words, word => word.Text == "WideIndentMarker").BoundingBox.Left;

            Assert.True(wideX > regularX + 30D, $"Expected direct Word list paragraph indentation to move the list text right. Regular x: {regularX:0.##}; wide x: {wideX:0.##}.");
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Honors_Paragraph_Style_List_Indentation() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeListStyleIndent.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeListStyleIndent.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                const string styleId = "NativeListStyleIndent";
                Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
                styles.Append(new Style(
                    new StyleName { Val = "Native List Style Indent" },
                    new BasedOn { Val = "Normal" },
                    new StyleParagraphProperties(new Indentation {
                        Left = "1440",
                        Hanging = "360"
                    }))
                {
                    Type = StyleValues.Paragraph,
                    StyleId = styleId,
                    CustomStyle = true
                });

                WordList bulletList = document.AddList(WordListStyle.Bulleted);
                WordListLevel level = bulletList.Numbering.Levels[0];
                level.IndentationLeft = 720;
                level.IndentationHanging = 360;

                bulletList.AddItem("RegularStyleFallbackMarker");
                WordParagraph styledIndent = bulletList.AddItem("StyledIndentMarker");
                styledIndent.SetStyleId(styleId);

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
            var words = pdf.GetPage(1).GetWords().ToList();
            double regularX = Assert.Single(words, word => word.Text == "RegularStyleFallbackMarker").BoundingBox.Left;
            double styledX = Assert.Single(words, word => word.Text == "StyledIndentMarker").BoundingBox.Left;

            Assert.True(styledX > regularX + 30D, $"Expected paragraph style indentation to move native Word list text right. Regular x: {regularX:0.##}; styled x: {styledX:0.##}.");
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Maps_List_Paragraph_Spacing_After_To_Item_Gaps() {
            double compactGap = RenderNativeListStyleSpacingGap("PdfNativeListCompactSpacing", spacingAfterTwips: "0", contextualSpacing: false);
            double spacedGap = RenderNativeListStyleSpacingGap("PdfNativeListParagraphSpacingAfter", spacingAfterTwips: "480", contextualSpacing: false);

            Assert.True(spacedGap > compactGap + 16D, $"Expected Word list paragraph spacing after to increase the PDF gap between list items. Compact gap: {compactGap:0.##}; spaced gap: {spacedGap:0.##}.");
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Honors_List_Style_Contextual_Spacing() {
            double spacedGap = RenderNativeListStyleSpacingGap("PdfNativeListSpacingNoContext", spacingAfterTwips: "480", contextualSpacing: false);
            double contextualGap = RenderNativeListStyleSpacingGap("PdfNativeListContextualSpacing", spacingAfterTwips: "480", contextualSpacing: true);

            Assert.True(spacedGap > contextualGap + 16D, $"Expected Word contextual spacing to suppress list item spacing between same-style paragraphs. Spaced gap: {spacedGap:0.##}; contextual gap: {contextualGap:0.##}.");
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Uses_Character_Style_Font_Size_For_List_Exact_Line_Spacing() {
            double gap = RenderNativeListCharacterStyleExactLineSpacingGap();

            Assert.InRange(gap, 16D, 22D);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Keeps_List_KeepWithNext_Chains_Together() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeListKeepWithNextChain.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeListKeepWithNextChain.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                const string styleId = "ListChainKeepWithNext";
                Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
                styles.Append(new Style(
                    new StyleName { Val = "List Chain Keep With Next" },
                    new BasedOn { Val = "Normal" },
                    new StyleParagraphProperties(new KeepNext()))
                {
                    Type = StyleValues.Paragraph,
                    StyleId = styleId,
                    CustomStyle = true
                });

                WordParagraph intro = document.AddParagraph("ListChainIntro");
                intro.LineSpacingAfterPoints = 120;

                WordList bulletList = document.AddList(WordListStyle.Bulleted);
                bulletList.AddItem("ListChainFirst").SetStyleId(styleId);
                bulletList.AddItem("ListChainSecond").SetStyleId(styleId);
                document.AddParagraph("ListChainBridge").SetStyleId(styleId);
                document.AddParagraph("ListChainTarget");

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false,
                    PageSize = new OfficeIMO.Pdf.PageSize(260, 260),
                    Margins = OfficeIMO.Pdf.PageMargins.Uniform(30),
                    FontFamily = "Helvetica"
                });
            }

            using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);

            Assert.Equal(2, pdf.NumberOfPages);
            Assert.Contains("ListChainIntro", pdf.GetPage(1).Text);
            Assert.DoesNotContain("ListChainFirst", pdf.GetPage(1).Text);
            Assert.DoesNotContain("ListChainSecond", pdf.GetPage(1).Text);
            Assert.DoesNotContain("ListChainBridge", pdf.GetPage(1).Text);
            Assert.DoesNotContain("ListChainTarget", pdf.GetPage(1).Text);
            Assert.Contains("ListChainFirst", pdf.GetPage(2).Text);
            Assert.Contains("ListChainSecond", pdf.GetPage(2).Text);
            Assert.Contains("ListChainBridge", pdf.GetPage(2).Text);
            Assert.Contains("ListChainTarget", pdf.GetPage(2).Text);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Measures_List_Blocks_Inside_KeepWithNext_Chains() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeParagraphListKeepWithNextChain.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeParagraphListKeepWithNextChain.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                const string styleId = "ParagraphListChainKeepWithNext";
                Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
                styles.Append(new Style(
                    new StyleName { Val = "Paragraph List Chain Keep With Next" },
                    new BasedOn { Val = "Normal" },
                    new StyleParagraphProperties(new KeepNext()))
                {
                    Type = StyleValues.Paragraph,
                    StyleId = styleId,
                    CustomStyle = true
                });

                WordParagraph intro = document.AddParagraph("ParagraphListChainIntro");
                intro.LineSpacingAfterPoints = 120;
                document.AddParagraph("ParagraphListChainLead").SetStyleId(styleId);

                WordList bulletList = document.AddList(WordListStyle.Bulleted);
                bulletList.AddItem("ParagraphListChainFirst").SetStyleId(styleId);
                bulletList.AddItem("ParagraphListChainSecond").SetStyleId(styleId);
                document.AddParagraph("ParagraphListChainTarget");

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false,
                    PageSize = new OfficeIMO.Pdf.PageSize(260, 260),
                    Margins = OfficeIMO.Pdf.PageMargins.Uniform(30),
                    FontFamily = "Helvetica"
                });
            }

            using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);

            Assert.Equal(2, pdf.NumberOfPages);
            Assert.Contains("ParagraphListChainIntro", pdf.GetPage(1).Text);
            Assert.DoesNotContain("ParagraphListChainLead", pdf.GetPage(1).Text);
            Assert.DoesNotContain("ParagraphListChainFirst", pdf.GetPage(1).Text);
            Assert.DoesNotContain("ParagraphListChainSecond", pdf.GetPage(1).Text);
            Assert.DoesNotContain("ParagraphListChainTarget", pdf.GetPage(1).Text);
            Assert.Contains("ParagraphListChainLead", pdf.GetPage(2).Text);
            Assert.Contains("ParagraphListChainFirst", pdf.GetPage(2).Text);
            Assert.Contains("ParagraphListChainSecond", pdf.GetPage(2).Text);
            Assert.Contains("ParagraphListChainTarget", pdf.GetPage(2).Text);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Renders_Custom_And_Nested_Word_List_Markers() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeCustomNestedListMarkers.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeCustomNestedListMarkers.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                WordList list = document.AddCustomList();
                list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.LowerLetterDot));
                list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.LowerRomanDot));
                list.Numbering.Levels[0].IndentationLeft = 720;
                list.Numbering.Levels[0].IndentationHanging = 360;
                list.Numbering.Levels[1].IndentationLeft = 1440;
                list.Numbering.Levels[1].IndentationHanging = 360;

                list.AddItem("Lower alpha item");
                list.AddItem("Nested roman item", 1);
                list.AddItem("Second alpha item");

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
            var page = pdf.GetPage(1);
            Assert.Contains("a.Lower alpha item", page.Text, StringComparison.Ordinal);
            Assert.Contains("i.Nested roman item", page.Text, StringComparison.Ordinal);
            Assert.Contains("b.Second alpha item", page.Text, StringComparison.Ordinal);

            var lineGroups = page.Letters
                .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
                .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
                .OrderByDescending(group => group.Key)
                .Select(group => group.OrderBy(letter => letter.StartBaseLine.X).ToList())
                .ToList();

            var alphaLine = lineGroups.First(line => string.Concat(line.Select(letter => letter.Value)).IndexOf("Loweralphaitem", StringComparison.Ordinal) >= 0);
            var nestedLine = lineGroups.First(line => string.Concat(line.Select(letter => letter.Value)).IndexOf("Nestedromanitem", StringComparison.Ordinal) >= 0);

            Assert.StartsWith("a.", string.Concat(alphaLine.Select(letter => letter.Value)), StringComparison.Ordinal);
            Assert.StartsWith("i.", string.Concat(nestedLine.Select(letter => letter.Value)), StringComparison.Ordinal);
            Assert.True(nestedLine[0].StartBaseLine.X > alphaLine[0].StartBaseLine.X + 30D, "Expected nested Word list marker to render with deeper indentation.");
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Renders_Word_Alphabetic_List_Markers_After_Z() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeAlphabeticListMarkersAfterZ.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeAlphabeticListMarkersAfterZ.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                WordList lowerList = document.AddCustomList();
                lowerList.Numbering.AddLevel(new WordListLevel(WordListLevelKind.LowerLetterDot).SetStartNumberingValue(27));
                lowerList.AddItem("Lower alphabetic wrap item");

                WordList upperList = document.AddCustomList();
                upperList.Numbering.AddLevel(new WordListLevel(WordListLevelKind.UpperLetterDot).SetStartNumberingValue(28));
                upperList.AddItem("Upper alphabetic wrap item");

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
            string text = pdf.GetPage(1).Text;
            Assert.Contains("aa.Lower alphabetic wrap item", text, StringComparison.Ordinal);
            Assert.Contains("AB.Upper alphabetic wrap item", text, StringComparison.Ordinal);
            Assert.DoesNotContain("{.Lower alphabetic wrap item", text, StringComparison.Ordinal);
            Assert.DoesNotContain("\\.Upper alphabetic wrap item", text, StringComparison.Ordinal);
        }

        private double RenderNativeListStyleSpacingGap(string fileNamePrefix, string spacingAfterTwips, bool contextualSpacing) {
            const string firstMarker = "FirstListGapMarker";
            const string secondMarker = "SecondListGapMarker";
            string docPath = Path.Combine(_directoryWithFiles, fileNamePrefix + ".docx");
            string pdfPath = Path.Combine(_directoryWithFiles, fileNamePrefix + ".pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                const string styleId = "NativeListSpacingStyle";
                var styleParagraphProperties = new StyleParagraphProperties(
                    new SpacingBetweenLines { After = spacingAfterTwips });
                if (contextualSpacing) {
                    styleParagraphProperties.Append(new ContextualSpacing());
                }

                Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
                styles.Append(new Style(
                    new StyleName { Val = "Native List Spacing Style" },
                    styleParagraphProperties)
                {
                    Type = StyleValues.Paragraph,
                    StyleId = styleId,
                    CustomStyle = true
                });

                WordList bulletList = document.AddList(WordListStyle.Bulleted);
                bulletList.AddItem(firstMarker).SetStyleId(styleId);
                bulletList.AddItem(secondMarker).SetStyleId(styleId);

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false,
                    PageSize = new OfficeIMO.Pdf.PageSize(320, 240),
                    Margins = PageMargins.Uniform(36),
                    FontFamily = "Helvetica"
                });
            }

            using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
            var words = pdf.GetPage(1).GetWords().ToList();
            double firstY = Assert.Single(words, word => word.Text == firstMarker).BoundingBox.Bottom;
            double secondY = Assert.Single(words, word => word.Text == secondMarker).BoundingBox.Bottom;
            return firstY - secondY;
        }

        private double RenderNativeListCharacterStyleExactLineSpacingGap() {
            const string firstMarker = "Alpha";
            const string secondMarker = "Beta";
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeListCharacterStyleExactLineSpacing.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeListCharacterStyleExactLineSpacing.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                const string styleId = "NativeListExactLineCharacterStyle";
                Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
                styles.Append(new Style(
                    new StyleName { Val = "Native List Exact Line Character Style" },
                    new StyleRunProperties(new FontSize { Val = "64" }))
                {
                    Type = StyleValues.Character,
                    StyleId = styleId,
                    CustomStyle = true
                });

                WordList bulletList = document.AddList(WordListStyle.Bulleted);
                WordParagraph item = bulletList.AddItem(string.Empty);
                item.AddText(firstMarker).SetCharacterStyleId(styleId);
                item.AddBreak();
                item.AddText(secondMarker).SetCharacterStyleId(styleId);
                item.LineSpacingAfterPoints = 0D;
                item.LineSpacingPoints = 18D;
                item.LineSpacingRule = LineSpacingRuleValues.Exact;

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false,
                    PageSize = new OfficeIMO.Pdf.PageSize(360, 260),
                    Margins = PageMargins.Uniform(36),
                    FontFamily = "Helvetica"
                });
            }

            using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
            var words = pdf.GetPage(1).GetWords().ToList();
            double firstY = Assert.Single(words, word => word.Text == firstMarker).BoundingBox.Bottom;
            double secondY = Assert.Single(words, word => word.Text == secondMarker).BoundingBox.Bottom;
            return firstY - secondY;
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Maps_List_Item_Bookmarks_Through_Native_List_Blocks() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeListBookmarks.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeListBookmarks.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                WordList bulletList = document.AddList(WordListStyle.Bulleted);
                bulletList.AddItem("Bookmarked native bullet").AddBookmark("NativeListBookmark");
                bulletList.AddItem("Following native bullet");

                WordParagraph linkParagraph = document.AddParagraph();
                linkParagraph.AddHyperLink("Jump to native list bookmark", "NativeListBookmark", addStyle: true, tooltip: "Native list bookmark metadata");

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            byte[] bytes = File.ReadAllBytes(pdfPath);
            PdfLogicalDocument logical = PdfLogicalDocument.Load(bytes, new PdfTextLayoutOptions {
                ForceSingleColumn = true
            });
            var listItems = PdfTextExtractor.ExtractListItemsByPage(bytes)
                .SelectMany(page => page.ListItems)
                .ToList();

            Assert.Contains(listItems, item => item.Text == "Bookmarked native bullet");
            Assert.Contains(listItems, item => item.Text == "Following native bullet");
            Assert.Contains(logical.NamedDestinations, destination => destination.Name == "NativeListBookmark");
            var bookmarkLinks = logical.GetLinksByDestinationName("NativeListBookmark").ToList();
            Assert.NotEmpty(bookmarkLinks);
            Assert.All(bookmarkLinks, link => Assert.Equal("Native list bookmark metadata", link.Contents));
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Maps_Rich_List_Item_Runs() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeRichListRuns.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeRichListRuns.pdf");
            const string linkUri = "https://evotec.xyz/native-list";

            using (WordDocument document = WordDocument.Create(docPath)) {
                WordList bulletList = document.AddList(WordListStyle.Bulleted);
                WordParagraph item = bulletList.AddItem(string.Empty);
                item.AddText("ListPlain ");
                WordParagraph red = item.AddText("ListRed");
                red.ColorHex = "FF0000";
                item.AddText(" ");
                item.AddText("ListBold").SetBold();
                item.AddText(" ");
                item.AddText("ListMarked").SetHighlight(HighlightColorValues.Yellow);
                item.AddText(" ");
                item.AddHyperLink("ListLink", new System.Uri(linkUri), addStyle: true, tooltip: "Native list link metadata");

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            byte[] bytes = File.ReadAllBytes(pdfPath);
            string content = ReadPdfPageContent(bytes);

            using (PdfPigDocument pdf = PdfPigDocument.Open(bytes)) {
                string pageText = string.Concat(pdf.GetPages().Select(page => page.Text));

                Assert.Equal(1, CountOccurrences(pageText, "ListPlain"));
                Assert.Equal(1, CountOccurrences(pageText, "ListRed"));
                Assert.Equal(1, CountOccurrences(pageText, "ListBold"));
                Assert.Equal(1, CountOccurrences(pageText, "ListMarked"));
                Assert.Equal(1, CountOccurrences(pageText, "ListLink"));
            }

            var listItems = PdfTextExtractor.ExtractListItemsByPage(bytes)
                .SelectMany(page => page.ListItems)
                .ToList();
            PdfLogicalDocument logical = PdfLogicalDocument.Load(bytes, new PdfTextLayoutOptions {
                ForceSingleColumn = true
            });
            PdfLogicalLinkAnnotation link = Assert.Single(logical.GetLinksByUri(linkUri));

            Assert.Contains(listItems, item => item.Text == "ListPlain ListRed ListBold ListMarked ListLink");
            Assert.Contains("1 0 0 rg", content, StringComparison.Ordinal);
            Assert.Matches(@"/F\d+\s+11\s+Tf", content);
            Assert.Contains("1 1 0 rg", content, StringComparison.Ordinal);
            Assert.Equal("Native list link metadata", link.Contents);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Maps_Paragraph_Style_Color_To_List_Marker() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeListStyleMarkerColor.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeListStyleMarkerColor.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                const string styleId = "NativeListMarkerColor";
                Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
                styles.Append(new Style(
                    new StyleName { Val = "Native List Marker Color" },
                    new BasedOn { Val = "Normal" },
                    new StyleRunProperties(new Color { Val = "C00000" }))
                {
                    Type = StyleValues.Paragraph,
                    StyleId = styleId,
                    CustomStyle = true
                });

                WordList numberedList = document.AddCustomList();
                numberedList.Numbering.AddLevel(new WordListLevel(WordListLevelKind.DecimalDot));
                WordParagraph styled = numberedList.AddItem("StyledListMarkerColor");
                styled.SetStyleId(styleId);

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false,
                    FontFamily = "Helvetica"
                });
            }

            byte[] bytes = File.ReadAllBytes(pdfPath);
            string content = ReadPdfPageContent(bytes);
            var listItems = PdfTextExtractor.ExtractListItemsByPage(bytes)
                .SelectMany(page => page.ListItems)
                .ToList();

            Assert.Contains(listItems, item => item.Marker == "1" && item.Text == "StyledListMarkerColor");
            Assert.True(
                CountOccurrences(content, "0.753 0 0 rg") >= 2,
                "Expected paragraph style color to be emitted for both the list marker and the list item text.");
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Maps_Paragraph_Style_BoldItalic_To_List_Marker() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeListStyleMarkerBoldItalic.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeListStyleMarkerBoldItalic.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                const string styleId = "NativeListMarkerBoldItalic";
                Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
                styles.Append(new Style(
                    new StyleName { Val = "Native List Marker Bold Italic" },
                    new BasedOn { Val = "Normal" },
                    new StyleRunProperties(new Bold(), new Italic()))
                {
                    Type = StyleValues.Paragraph,
                    StyleId = styleId,
                    CustomStyle = true
                });

                WordList numberedList = document.AddCustomList();
                numberedList.Numbering.AddLevel(new WordListLevel(WordListLevelKind.DecimalDot));
                WordParagraph styled = numberedList.AddItem("StyledListMarkerBoldItalic");
                styled.SetStyleId(styleId);

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false,
                    FontFamily = "Helvetica"
                });
            }

            byte[] bytes = File.ReadAllBytes(pdfPath);
            string content = ReadPdfPageContent(bytes);
            var listItems = PdfTextExtractor.ExtractListItemsByPage(bytes)
                .SelectMany(page => page.ListItems)
                .ToList();

            Assert.Contains(listItems, item => item.Marker == "1" && item.Text == "StyledListMarkerBoldItalic");
            Assert.True(
                Regex.Matches(content, @"/F4\s+11\s+Tf").Count >= 2,
                "Expected paragraph style bold italic typography to be emitted for both the list marker and the list item text.");
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Maps_Paragraph_Style_Font_To_List_Marker() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeListStyleMarkerFont.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeListStyleMarkerFont.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                const string styleId = "NativeListMarkerFont";
                Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
                styles.Append(new Style(
                    new StyleName { Val = "Native List Marker Font" },
                    new BasedOn { Val = "Normal" },
                    new StyleRunProperties(new RunFonts { Ascii = "Courier New", HighAnsi = "Courier New" }))
                {
                    Type = StyleValues.Paragraph,
                    StyleId = styleId,
                    CustomStyle = true
                });

                WordList numberedList = document.AddCustomList();
                numberedList.Numbering.AddLevel(new WordListLevel(WordListLevelKind.DecimalDot));
                WordParagraph styled = numberedList.AddItem("StyledListMarkerFont");
                styled.SetStyleId(styleId);

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false,
                    FontFamily = "Helvetica"
                });
            }

            byte[] bytes = File.ReadAllBytes(pdfPath);
            string content = ReadPdfPageContent(bytes);
            var listItems = PdfTextExtractor.ExtractListItemsByPage(bytes)
                .SelectMany(page => page.ListItems)
                .ToList();

            Assert.Contains(listItems, item => item.Marker == "1" && item.Text == "StyledListMarkerFont");
            Assert.True(
                Regex.Matches(content, @"/F19\s+11\s+Tf").Count >= 2,
                "Expected paragraph style font family to be emitted for both the list marker and the list item text.");
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Maps_Numbering_Level_Run_Properties_To_List_Marker() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeListLevelMarkerRunProperties.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeListLevelMarkerRunProperties.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                WordList numberedList = document.AddCustomList();
                numberedList.Numbering.AddLevel(new WordListLevel(WordListLevelKind.DecimalDot));
                numberedList.Numbering.Levels[0].OpenXmlElement.NumberingSymbolRunProperties = new NumberingSymbolRunProperties(
                    new RunFonts { Ascii = "Courier New", HighAnsi = "Courier New" },
                    new Color { Val = "C00000" });
                numberedList.AddItem("LevelMarkerRunPropertiesBody");

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false,
                    FontFamily = "Helvetica"
                });
            }

            byte[] bytes = File.ReadAllBytes(pdfPath);
            string content = ReadPdfPageContent(bytes);
            var listItems = PdfTextExtractor.ExtractListItemsByPage(bytes)
                .SelectMany(page => page.ListItems)
                .ToList();

            Assert.Contains(listItems, item => item.Marker == "1" && item.Text == "LevelMarkerRunPropertiesBody");
            Assert.Equal(1, CountOccurrences(content, "0.753 0 0 rg"));
            Assert.Equal(1, Regex.Matches(content, @"/F19\s+11\s+Tf").Count);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Maps_Numbering_Level_Font_Size_To_List_Marker() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeListLevelMarkerFontSize.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeListLevelMarkerFontSize.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                WordList numberedList = document.AddCustomList();
                numberedList.Numbering.AddLevel(new WordListLevel(WordListLevelKind.DecimalDot));
                numberedList.Numbering.Levels[0].OpenXmlElement.NumberingSymbolRunProperties = new NumberingSymbolRunProperties(
                    new FontSize { Val = "40" });
                numberedList.AddItem("LevelMarkerFontSizeBody");

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false,
                    FontFamily = "Helvetica"
                });
            }

            byte[] bytes = File.ReadAllBytes(pdfPath);
            string content = ReadPdfPageContent(bytes);
            var listItems = PdfTextExtractor.ExtractListItemsByPage(bytes)
                .SelectMany(page => page.ListItems)
                .ToList();

            Assert.Contains(listItems, item => item.Marker == "1" && item.Text == "LevelMarkerFontSizeBody");
            Assert.Equal(1, Regex.Matches(content, @"/F\d+\s+20\s+Tf").Count);
            Assert.True(
                Regex.Matches(content, @"/F\d+\s+11\s+Tf").Count >= 1,
                "Expected the list body text to keep the normal paragraph font size while only the marker uses the numbering level font size.");
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Honors_Numbering_Level_Marker_Justification() {
            (double MarkerX, double TextX) left = RenderNativeNumberedListMarkerJustification(
                "PdfNativeListMarkerJustificationLeft",
                LevelJustificationValues.Left,
                "LeftJustifiedMarkerBody");
            (double MarkerX, double TextX) right = RenderNativeNumberedListMarkerJustification(
                "PdfNativeListMarkerJustificationRight",
                LevelJustificationValues.Right,
                "RightJustifiedMarkerBody");

            Assert.True(
                right.MarkerX > left.MarkerX + 20D,
                $"Expected right-justified marker X ({right.MarkerX}) to move inside the marker column compared to left-justified marker X ({left.MarkerX}).");
            Assert.InRange(Math.Abs(right.TextX - left.TextX), 0D, 1.5D);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Honors_Numbering_Level_Marker_Suffix() {
            (double MarkerX, double TextX) nothing = RenderNativeNumberedListMarkerSuffix(
                "PdfNativeListMarkerSuffixNothing",
                LevelSuffixValues.Nothing,
                "NothingSuffixMarkerBody");
            (double MarkerX, double TextX) space = RenderNativeNumberedListMarkerSuffix(
                "PdfNativeListMarkerSuffixSpace",
                LevelSuffixValues.Space,
                "SpaceSuffixMarkerBody");
            (double MarkerX, double TextX) tab = RenderNativeNumberedListMarkerSuffix(
                "PdfNativeListMarkerSuffixTab",
                LevelSuffixValues.Tab,
                "TabSuffixMarkerBody");

            Assert.InRange(Math.Abs(space.MarkerX - nothing.MarkerX), 0D, 1.5D);
            Assert.InRange(Math.Abs(tab.MarkerX - nothing.MarkerX), 0D, 1.5D);
            Assert.True(
                space.TextX > nothing.TextX + 2D,
                $"Expected Word marker suffix 'space' to move list text after 'nothing'. Nothing x: {nothing.TextX:0.##}; space x: {space.TextX:0.##}.");
            Assert.True(
                tab.TextX > space.TextX + 20D,
                $"Expected Word marker suffix 'tab' to preserve the hanging-indent text position beyond 'space'. Space x: {space.TextX:0.##}; tab x: {tab.TextX:0.##}.");
        }

        private (double MarkerX, double TextX) RenderNativeNumberedListMarkerJustification(string fileNamePrefix, LevelJustificationValues justification, string bodyText) {
            string docPath = Path.Combine(_directoryWithFiles, fileNamePrefix + ".docx");
            string pdfPath = Path.Combine(_directoryWithFiles, fileNamePrefix + ".pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                WordList numberedList = document.AddCustomList();
                numberedList.Numbering.AddLevel(new WordListLevel(WordListLevelKind.DecimalDot));
                WordListLevel level = numberedList.Numbering.Levels[0];
                level.IndentationLeft = 1440;
                level.IndentationHanging = 720;
                level.OpenXmlElement.LevelJustification = new LevelJustification { Val = justification };
                numberedList.AddItem(bodyText);

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false,
                    FontFamily = "Helvetica"
                });
            }

            using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
            string firstBodyLetter = bodyText[0].ToString();
            var line = pdf.GetPage(1).Letters
                .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
                .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
                .OrderByDescending(group => group.Key)
                .Select(group => group.OrderBy(letter => letter.StartBaseLine.X).ToList())
                .First(group => string.Concat(group.Select(letter => letter.Value)).Contains(bodyText));

            double markerX = line.First(letter => letter.Value == "1").StartBaseLine.X;
            double textX = line.First(letter => letter.Value == firstBodyLetter).StartBaseLine.X;
            return (markerX, textX);
        }

        private (double MarkerX, double TextX) RenderNativeNumberedListMarkerSuffix(string fileNamePrefix, LevelSuffixValues suffix, string bodyText) {
            string docPath = Path.Combine(_directoryWithFiles, fileNamePrefix + ".docx");
            string pdfPath = Path.Combine(_directoryWithFiles, fileNamePrefix + ".pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                WordList numberedList = document.AddCustomList();
                numberedList.Numbering.AddLevel(new WordListLevel(WordListLevelKind.DecimalDot));
                WordListLevel level = numberedList.Numbering.Levels[0];
                level.IndentationLeft = 1440;
                level.IndentationHanging = 720;
                level.LevelSuffix = suffix;
                numberedList.AddItem(bodyText);

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false,
                    FontFamily = "Helvetica"
                });
            }

            using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
            string firstBodyLetter = bodyText[0].ToString();
            var line = pdf.GetPage(1).Letters
                .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
                .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
                .OrderByDescending(group => group.Key)
                .Select(group => group.OrderBy(letter => letter.StartBaseLine.X).ToList())
                .First(group => string.Concat(group.Select(letter => letter.Value)).Contains(bodyText));

            double markerX = line.First(letter => letter.Value == "1").StartBaseLine.X;
            double textX = line.First(letter => letter.Value == firstBodyLetter).StartBaseLine.X;
            return (markerX, textX);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Maps_List_Item_Footnotes_Through_Native_List_Blocks() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeListFootnotes.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeListFootnotes.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                WordList bulletList = document.AddList(WordListStyle.Bulleted);
                bulletList.AddItem("Footnoted native list item").AddFootNote("Native list footnote text");
                bulletList.AddItem("Following native list item");

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            Assert.True(File.Exists(pdfPath));
            byte[] bytes = File.ReadAllBytes(pdfPath);
            var listItems = PdfTextExtractor.ExtractListItemsByPage(bytes)
                .SelectMany(page => page.ListItems)
                .ToList();

            using (PdfPigDocument pdf = PdfPigDocument.Open(bytes)) {
                string allText = string.Concat(pdf.GetPages().Select(page => page.Text));

                Assert.Contains("Footnoted native list item1", allText);
                Assert.Contains("Following native list item", allText);
                Assert.Equal(1, CountOccurrences(allText, "Native list footnote text"));
            }

            Assert.Contains(listItems, item => item.Text == "Footnoted native list item");
            Assert.DoesNotContain(listItems, item => item.Text == "Footnoted native list item1");
            Assert.Contains(listItems, item => item.Text == "Following native list item");
        }
    }
}
