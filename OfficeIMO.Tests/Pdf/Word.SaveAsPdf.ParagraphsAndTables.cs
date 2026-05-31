using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Pdf;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using UglyToad.PdfPig;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void SaveAsPdf_Renders_Paragraphs() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfParagraphs.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfParagraphs.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("First paragraph");
                document.AddParagraph("Second paragraph");
                document.Save();
                document.SaveAsPdf(pdfPath);
            }

            Assert.True(File.Exists(pdfPath));
            using (PdfDocument pdf = PdfDocument.Open(pdfPath)) {
                string allText = string.Concat(pdf.GetPages().Select(p => p.Text));
                int first = allText.IndexOf("First paragraph", StringComparison.Ordinal);
                int second = allText.IndexOf("Second paragraph", StringComparison.Ordinal);
                Assert.True(first >= 0 && second > first);
            }
        }

        [Fact]
        public void SaveAsPdf_Renders_Tables() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfTableContent.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfTableContent.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                WordTable table = document.AddTable(2, 2);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "A1";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "B1";
                table.Rows[1].Cells[0].Paragraphs[0].Text = "A2";
                table.Rows[1].Cells[1].Paragraphs[0].Text = "B2";
                document.Save();
                document.SaveAsPdf(pdfPath);
            }

            Assert.True(File.Exists(pdfPath));
            using (PdfDocument pdf = PdfDocument.Open(pdfPath)) {
                string allText = string.Concat(pdf.GetPages().Select(p => p.Text));
                int a1 = allText.IndexOf("A1", StringComparison.Ordinal);
                int b1 = allText.IndexOf("B1", StringComparison.Ordinal);
                int a2 = allText.IndexOf("A2", StringComparison.Ordinal);
                int b2 = allText.IndexOf("B2", StringComparison.Ordinal);
                Assert.True(a1 >= 0 && b1 > a1 && a2 > b1 && b2 > a2);
            }
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Renders_Paragraphs_Headings_And_Tables() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeContent.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeContent.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("Native heading").SetStyle(WordParagraphStyles.Heading1);
                document.AddParagraph("First native paragraph");
                document.AddParagraph("Second native paragraph").SetBold().SetItalic();
                WordTable table = document.AddTable(2, 2);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "N-A1";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "N-B1";
                table.Rows[1].Cells[0].Paragraphs[0].Text = "N-A2";
                table.Rows[1].Cells[1].Paragraphs[0].Text = "N-B2";
                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            Assert.True(File.Exists(pdfPath));
            using (PdfDocument pdf = PdfDocument.Open(pdfPath)) {
                string allText = string.Concat(pdf.GetPages().Select(p => p.Text));
                Assert.Contains("Native heading", allText);
                int first = allText.IndexOf("First native paragraph", StringComparison.Ordinal);
                int second = allText.IndexOf("Second native paragraph", StringComparison.Ordinal);
                int a1 = allText.IndexOf("N-A1", StringComparison.Ordinal);
                int b1 = allText.IndexOf("N-B1", StringComparison.Ordinal);
                int a2 = allText.IndexOf("N-A2", StringComparison.Ordinal);
                int b2 = allText.IndexOf("N-B2", StringComparison.Ordinal);
                Assert.True(first >= 0 && second > first);
                Assert.True(a1 >= 0 && b1 > a1 && a2 > b1 && b2 > a2);
            }
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Renders_Paragraph_Baseline_Formatting() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeParagraphBaseline.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeParagraphBaseline.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("Before native baseline formatting");
                WordParagraph superscript = document.AddParagraph("Native paragraph superscript");
                superscript.FontSize = 20;
                superscript.SetSuperScript();
                WordParagraph subscript = document.AddParagraph("Native paragraph subscript");
                subscript.FontSize = 20;
                subscript.SetSubScript();
                WordParagraph mixed = document.AddParagraph();
                mixed.AddText("Native mixed baseline ");
                WordParagraph runSuperscript = mixed.AddText("run superscript");
                runSuperscript.FontSize = 20;
                runSuperscript.SetSuperScript();
                mixed.AddText(" ");
                WordParagraph runSubscript = mixed.AddText("run subscript");
                runSubscript.FontSize = 20;
                runSubscript.SetSubScript();
                document.AddParagraph("After native baseline formatting");
                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            Assert.True(File.Exists(pdfPath));
            using (PdfDocument pdf = PdfDocument.Open(pdfPath)) {
                string allText = string.Concat(pdf.GetPages().Select(p => p.Text));
                Assert.Contains("Before native baseline formatting", allText);
                Assert.Contains("Native paragraph superscript", allText);
                Assert.Contains("Native paragraph subscript", allText);
                Assert.Contains("Native mixed baseline", allText);
                Assert.Contains("run superscript", allText);
                Assert.Contains("run subscript", allText);
                Assert.Contains("After native baseline formatting", allText);
            }

            string content = Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath));
            Assert.Contains("7 Ts", content);
            Assert.Contains("-3.6 Ts", content);
            Assert.Contains("0 Ts", content);
            Assert.Contains("/F1 13 Tf", content);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Renders_Word_Text_Wrapping_Breaks() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTextWrappingBreaks.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTextWrappingBreaks.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("Before native text wrapping breaks");
                WordParagraph paragraph = document.AddParagraph("NativeSoftFirst");
                paragraph.AddBreak();
                paragraph.AddText("NativeSoftSecond");
                WordTable table = document.AddTable(1, 1);
                WordParagraph cellParagraph = table.Rows[0].Cells[0].Paragraphs[0];
                cellParagraph.Text = string.Empty;
                cellParagraph.AddText("CellSoftFirst");
                cellParagraph.AddBreak();
                cellParagraph.AddText("CellSoftSecond");
                document.AddParagraph("After native text wrapping breaks");
                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            Assert.True(File.Exists(pdfPath));
            using (PdfDocument pdf = PdfDocument.Open(pdfPath)) {
                string allText = string.Concat(pdf.GetPages().Select(p => p.Text));
                Assert.Contains("NativeSoftFirst", allText);
                Assert.Contains("NativeSoftSecond", allText);
                Assert.Contains("CellSoftFirst", allText);
                Assert.Contains("CellSoftSecond", allText);

                var words = pdf.GetPage(1).GetWords().ToList();
                double paragraphFirstY = Assert.Single(words, word => word.Text == "NativeSoftFirst").BoundingBox.Bottom;
                double paragraphSecondY = Assert.Single(words, word => word.Text == "NativeSoftSecond").BoundingBox.Bottom;
                double cellFirstY = Assert.Single(words, word => word.Text == "CellSoftFirst").BoundingBox.Bottom;
                double cellSecondY = Assert.Single(words, word => word.Text == "CellSoftSecond").BoundingBox.Bottom;

                Assert.True(paragraphFirstY > paragraphSecondY + 8D, $"Expected Word paragraph soft break to move following text to the next line. First y: {paragraphFirstY:0.##}, second y: {paragraphSecondY:0.##}.");
                Assert.True(cellFirstY > cellSecondY + 8D, $"Expected Word table cell soft break to move following text to the next line. First y: {cellFirstY:0.##}, second y: {cellSecondY:0.##}.");
                Assert.InRange(paragraphFirstY - paragraphSecondY, 10.5D, 14.5D);
            }
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Maps_Justified_Paragraphs() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeJustifiedParagraph.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeJustifiedParagraph.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                WordParagraph paragraph = document.AddParagraph("Native justified paragraph alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu nu xi omicron pi rho sigma tau wraps across multiple visual lines.");
                paragraph.ParagraphAlignment = JustificationValues.Both;
                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false,
                    OfficeIMOPageSize = new OfficeIMO.Pdf.PageSize(240, 360),
                    OfficeIMOMargins = OfficeIMO.Pdf.PageMargins.Uniform(24)
                });
            }

            Assert.True(File.Exists(pdfPath));
            using (PdfDocument pdf = PdfDocument.Open(pdfPath)) {
                string allText = string.Concat(pdf.GetPages().Select(page => page.Text));
                Assert.Contains("Native justified paragraph", allText);
            }

            string content = Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath));
            Assert.Matches(new Regex(@"(?:0\.[1-9]\d*|[1-9]\d*(?:\.\d+)?) Tw"), content);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Renders_Linked_Heading_As_Heading_Link() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeLinkedHeading.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeLinkedHeading.pdf");
            const string linkUri = "https://evotec.xyz/native-heading";

            using (WordDocument document = WordDocument.Create(docPath)) {
                WordParagraph heading = document.AddParagraph();
                heading.SetStyle(WordParagraphStyles.Heading1);
                heading.AddHyperLink("Native linked heading", new System.Uri(linkUri), addStyle: true, tooltip: "Native heading metadata");
                document.AddParagraph("Native body after linked heading");

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            byte[] bytes = File.ReadAllBytes(pdfPath);
            PdfLogicalDocument logical = PdfLogicalDocument.Load(bytes, new PdfTextLayoutOptions {
                ForceSingleColumn = true
            });

            Assert.Contains(logical.Headings, heading => heading.Text == "Native linked heading");
            PdfLogicalLinkAnnotation link = Assert.Single(logical.GetLinksByUri(linkUri));
            Assert.Equal("Native heading metadata", link.Contents);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Renders_Bookmark_Linked_Heading_As_Internal_Heading_Link() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeBookmarkLinkedHeading.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeBookmarkLinkedHeading.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("Native heading bookmark target").AddBookmark("NativeHeadingTarget");
                WordParagraph heading = document.AddParagraph();
                heading.SetStyle(WordParagraphStyles.Heading1);
                heading.AddHyperLink("Native bookmark linked heading", "NativeHeadingTarget", addStyle: true, tooltip: "Native bookmark heading metadata");
                document.AddParagraph("Native body after bookmark linked heading");

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            byte[] bytes = File.ReadAllBytes(pdfPath);
            PdfLogicalDocument logical = PdfLogicalDocument.Load(bytes, new PdfTextLayoutOptions {
                ForceSingleColumn = true
            });

            Assert.Contains(logical.Headings, heading => heading.Text == "Native bookmark linked heading");
            Assert.Contains(logical.NamedDestinations, destination => destination.Name == "NativeHeadingTarget");
            PdfLogicalLinkAnnotation link = Assert.Single(logical.GetLinksByDestinationName("NativeHeadingTarget"));
            Assert.True(link.IsNamedDestinationLink);
            Assert.Equal("Native bookmark heading metadata", link.Contents);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Preserves_Paragraph_Link_Metadata() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeParagraphLinkMetadata.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeParagraphLinkMetadata.pdf");
            const string linkUri = "https://evotec.xyz/native-paragraph-link";

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("Native paragraph bookmark target").AddBookmark("NativeParagraphTarget");
                WordParagraph external = document.AddParagraph();
                external.AddHyperLink("Native paragraph external link", new System.Uri(linkUri), addStyle: true, tooltip: "Native paragraph external metadata");
                WordParagraph internalLink = document.AddParagraph();
                internalLink.AddHyperLink("Native paragraph bookmark link", "NativeParagraphTarget", addStyle: true, tooltip: "Native paragraph bookmark metadata");

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            byte[] bytes = File.ReadAllBytes(pdfPath);
            PdfLogicalDocument logical = PdfLogicalDocument.Load(bytes, new PdfTextLayoutOptions {
                ForceSingleColumn = true
            });

            var externalLinks = logical.GetLinksByUri(linkUri).ToList();
            Assert.NotEmpty(externalLinks);
            Assert.All(externalLinks, link => Assert.Equal("Native paragraph external metadata", link.Contents));
            Assert.Contains(logical.NamedDestinations, destination => destination.Name == "NativeParagraphTarget");
            var bookmarkLinks = logical.GetLinksByDestinationName("NativeParagraphTarget").ToList();
            Assert.NotEmpty(bookmarkLinks);
            Assert.All(bookmarkLinks, link => {
                Assert.True(link.IsNamedDestinationLink);
                Assert.Equal("Native paragraph bookmark metadata", link.Contents);
            });
        }

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

            using PdfDocument pdf = PdfDocument.Open(pdfPath);
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

            using PdfDocument pdf = PdfDocument.Open(pdfPath);
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
        public void SaveAsPdf_OfficeIMOEngine_Renders_TableOfContents_With_Heading_Entries() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTableOfContents.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTableOfContents.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddTableOfContent();
                document.AddParagraph("Native TOC first heading").SetStyle(WordParagraphStyles.Heading1);
                document.AddParagraph("Native TOC first body");
                document.AddPageBreak();
                document.AddParagraph("Native TOC second heading").SetStyle(WordParagraphStyles.Heading2);
                document.AddParagraph("Native TOC second body");

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            Assert.True(File.Exists(pdfPath));
            byte[] bytes = File.ReadAllBytes(pdfPath);
            using (PdfDocument pdf = PdfDocument.Open(bytes)) {
                string allText = string.Concat(pdf.GetPages().Select(page => page.Text));

                Assert.Contains("Table of Contents", allText);
                Assert.True(CountOccurrences(allText, "Native TOC first heading") >= 2, "Expected the first heading in the TOC and again in body content.");
                Assert.True(CountOccurrences(allText, "Native TOC second heading") >= 2, "Expected the second heading in the TOC and again in body content.");
                Assert.True(allText.IndexOf("Native TOC first heading", StringComparison.Ordinal) < allText.LastIndexOf("Native TOC first heading", StringComparison.Ordinal));
                Assert.True(allText.IndexOf("Native TOC second heading", StringComparison.Ordinal) < allText.LastIndexOf("Native TOC second heading", StringComparison.Ordinal));
            }

            PdfLogicalDocument logical = PdfLogicalDocument.Load(bytes, new PdfTextLayoutOptions {
                ForceSingleColumn = true
            });
            const string firstDestination = "officeimo-heading-native-toc-first-heading";
            const string secondDestination = "officeimo-heading-native-toc-second-heading";

            Assert.Contains(logical.NamedDestinations, destination => destination.Name == firstDestination);
            Assert.Contains(logical.NamedDestinations, destination => destination.Name == secondDestination);
            var firstTocLinks = logical.GetLinksByDestinationName(firstDestination).ToList();
            var secondTocLinks = logical.GetLinksByDestinationName(secondDestination).ToList();
            Assert.NotEmpty(firstTocLinks);
            Assert.NotEmpty(secondTocLinks);
            Assert.All(firstTocLinks, link => Assert.Equal("Table of contents: Native TOC first heading", link.Contents));
            Assert.All(secondTocLinks, link => Assert.Equal("Table of contents: Native TOC second heading", link.Contents));
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Creates_Pdf_Outlines_From_Word_Headings() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeHeadingOutlines.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeHeadingOutlines.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("Native outline root").SetStyle(WordParagraphStyles.Heading1);
                document.AddParagraph("Native outline body");
                document.AddParagraph("Native outline child").SetStyle(WordParagraphStyles.Heading2);
                document.AddParagraph("Native outline child body");
                document.AddPageBreak();
                document.AddParagraph("Native outline appendix").SetStyle(WordParagraphStyles.Heading1);
                document.AddParagraph("Native appendix body");

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            byte[] bytes = File.ReadAllBytes(pdfPath);
            PdfDocumentInfo info = PdfInspector.Inspect(bytes);

            Assert.Equal(2, info.Outlines.Count);
            Assert.Equal("Native outline root", info.Outlines[0].Title);
            Assert.Equal(1, info.Outlines[0].Level);
            Assert.Equal(1, info.Outlines[0].PageNumber);

            PdfOutlineItem child = Assert.Single(info.Outlines[0].Children);
            Assert.Equal("Native outline child", child.Title);
            Assert.Equal(2, child.Level);
            Assert.Equal(1, child.PageNumber);

            Assert.Equal("Native outline appendix", info.Outlines[1].Title);
            Assert.Equal(1, info.Outlines[1].Level);
            Assert.Equal(2, info.Outlines[1].PageNumber);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Renders_Normal_Word_Headings_As_Logical_Headings() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeNormalHeading.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeNormalHeading.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("Native normal heading").SetStyle(WordParagraphStyles.Heading1);
                document.AddParagraph("Native body after normal heading");

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            byte[] bytes = File.ReadAllBytes(pdfPath);
            PdfLogicalDocument logical = PdfLogicalDocument.Load(bytes, new PdfTextLayoutOptions {
                ForceSingleColumn = true
            });

            Assert.Contains(logical.Headings, heading => heading.Text == "Native normal heading");
            string rawPdf = Encoding.ASCII.GetString(bytes);
            Assert.DoesNotContain("/Helvetica-Bold", rawPdf, StringComparison.Ordinal);

            MethodInfo method = typeof(WordPdfConverterExtensions).GetMethod("CreateNativeWordHeadingStyle", BindingFlags.NonPublic | BindingFlags.Static)!;
            PdfHeadingStyle headingStyle = Assert.IsType<PdfHeadingStyle>(method.Invoke(null, new object[] { 1 }));
            Assert.True(headingStyle.ApplySpacingBeforeAtTop);
            Assert.Equal(24D, headingStyle.SpacingBefore);
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
                red.ColorHex = "ff0000";
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
            string content = Encoding.ASCII.GetString(bytes);
            int redText = content.IndexOf("<4C697374526564>", StringComparison.Ordinal);
            int boldText = content.IndexOf("<4C697374426F6C64>", StringComparison.Ordinal);
            int markedText = content.IndexOf("<4C6973744D61726B6564>", StringComparison.Ordinal);

            using (PdfDocument pdf = PdfDocument.Open(bytes)) {
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
            Assert.True(redText >= 0, "Expected encoded 'ListRed' text in the native list PDF content stream.");
            Assert.True(boldText > redText, "Expected encoded 'ListBold' text after the colored list run.");
            Assert.True(markedText > boldText, "Expected encoded 'ListMarked' text after the bold list run.");
            Assert.True(content.LastIndexOf("1 0 0 rg", redText, StringComparison.Ordinal) >= 0, "Expected Word list run color to emit a red PDF fill color.");
            Assert.True(content.LastIndexOf("/F2 11 Tf", boldText, StringComparison.Ordinal) >= 0, "Expected Word list bold run to use the bold PDF font resource.");
            Assert.True(content.LastIndexOf("1 1 0 rg", markedText, StringComparison.Ordinal) >= 0, "Expected Word list run highlight to emit a yellow PDF fill color.");
            Assert.Equal("Native list link metadata", link.Contents);
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

            using (PdfDocument pdf = PdfDocument.Open(bytes)) {
                string allText = string.Concat(pdf.GetPages().Select(page => page.Text));

                Assert.Contains("Footnoted native list item1", allText);
                Assert.Contains("Following native list item", allText);
                Assert.Equal(1, CountOccurrences(allText, "Native list footnote text"));
            }

            Assert.Contains(listItems, item => item.Text == "Footnoted native list item");
            Assert.DoesNotContain(listItems, item => item.Text == "Footnoted native list item1");
            Assert.Contains(listItems, item => item.Text == "Following native list item");
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Renders_Paragraph_Shading_And_Uniform_Borders() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeParagraphPanel.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeParagraphPanel.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                WordParagraph paragraph = document.AddParagraph("Native shaded panel paragraph");
                paragraph.ShadingFillColorHex = "e6f2ff";
                paragraph.Borders.TopStyle = BorderValues.Single;
                paragraph.Borders.BottomStyle = BorderValues.Single;
                paragraph.Borders.LeftStyle = BorderValues.Single;
                paragraph.Borders.RightStyle = BorderValues.Single;
                paragraph.Borders.TopColorHex = "336699";
                paragraph.Borders.BottomColorHex = "336699";
                paragraph.Borders.LeftColorHex = "336699";
                paragraph.Borders.RightColorHex = "336699";
                paragraph.Borders.TopSize = 8;
                paragraph.Borders.BottomSize = 8;
                paragraph.Borders.LeftSize = 8;
                paragraph.Borders.RightSize = 8;

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            Assert.True(File.Exists(pdfPath));
            byte[] bytes = File.ReadAllBytes(pdfPath);
            using (PdfDocument pdf = PdfDocument.Open(bytes)) {
                string text = string.Concat(pdf.GetPages().Select(p => p.Text));
                Assert.Contains("Native shaded panel paragraph", text);
            }

            string raw = Encoding.ASCII.GetString(bytes);
            Assert.Contains("0.902 0.949 1 rg", raw);
            Assert.Contains("0.2 0.4 0.6 RG", raw);
            Assert.Contains("1 w", raw);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Renders_Word_Horizontal_Line() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeHorizontalLine.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeHorizontalLine.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("Before native horizontal line");
                document.AddHorizontalLine(BorderValues.Single, OfficeIMO.Drawing.OfficeColor.Red, size: 16, space: 4);
                document.AddParagraph("After native horizontal line");

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false,
                    OfficeIMOPageSize = new OfficeIMO.Pdf.PageSize(300, 180),
                    OfficeIMOMargins = OfficeIMO.Pdf.PageMargins.Uniform(30)
                });
            }

            byte[] bytes = File.ReadAllBytes(pdfPath);
            string raw = Encoding.ASCII.GetString(bytes);
            using (PdfDocument pdf = PdfDocument.Open(bytes)) {
                string text = string.Concat(pdf.GetPages().Select(page => page.Text));

                Assert.Contains("Before native horizontal", text);
                Assert.Contains("After native horizontal", text);
            }

            Assert.Contains("1 0 0 RG", raw);
            Assert.Contains("2 w", raw);
            Assert.DoesNotContain("Before native horizontal lineAfter native horizontal line", raw);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Renders_Paragraph_Bottom_Border_As_Rule() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeParagraphBottomBorder.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeParagraphBottomBorder.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                WordParagraph paragraph = document.AddParagraph("Native bordered paragraph heading");
                paragraph.Borders.BottomStyle = BorderValues.Single;
                paragraph.Borders.BottomColorHex = "336699";
                paragraph.Borders.BottomSize = 12;
                paragraph.Borders.BottomSpace = 3;
                paragraph.LineSpacingAfterPoints = 8;

                document.AddParagraph("After native bottom border");

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false,
                    OfficeIMOPageSize = new OfficeIMO.Pdf.PageSize(360, 200),
                    OfficeIMOMargins = OfficeIMO.Pdf.PageMargins.Uniform(30)
                });
            }

            byte[] bytes = File.ReadAllBytes(pdfPath);
            string raw = Encoding.ASCII.GetString(bytes);
            using (PdfDocument pdf = PdfDocument.Open(bytes)) {
                string text = string.Concat(pdf.GetPages().Select(page => page.Text));

                Assert.Contains("Native bordered paragraph", text);
                Assert.Contains("After native bottom border", text);
            }

            Assert.Contains("0.2 0.4 0.6 RG", raw);
            Assert.Contains("1.5 w", raw);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Renders_Paragraph_Top_Border_As_Rule() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeParagraphTopBorder.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeParagraphTopBorder.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("Before native top border");

                WordParagraph paragraph = document.AddParagraph("Native top bordered paragraph");
                paragraph.Borders.TopStyle = BorderValues.Single;
                paragraph.Borders.TopColorHex = "008000";
                paragraph.Borders.TopSize = 16;
                paragraph.Borders.TopSpace = 4;
                paragraph.LineSpacingBeforePoints = 8;

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false,
                    OfficeIMOPageSize = new OfficeIMO.Pdf.PageSize(360, 200),
                    OfficeIMOMargins = OfficeIMO.Pdf.PageMargins.Uniform(30)
                });
            }

            byte[] bytes = File.ReadAllBytes(pdfPath);
            string raw = Encoding.ASCII.GetString(bytes);
            using (PdfDocument pdf = PdfDocument.Open(bytes)) {
                string text = string.Concat(pdf.GetPages().Select(page => page.Text));

                Assert.Contains("Before native top border", text);
                Assert.Contains("Native top bordered", text);
                Assert.Contains("paragraph", text);
            }

            Assert.Contains("0 0.502 0 RG", raw);
            Assert.Contains("2 w", raw);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Renders_NonUniform_Paragraph_Borders_As_Panel_Sides() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeParagraphSideBorders.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeParagraphSideBorders.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                WordParagraph paragraph = document.AddParagraph("Native side bordered paragraph");
                paragraph.Borders.LeftStyle = BorderValues.Single;
                paragraph.Borders.LeftColorHex = "ff0000";
                paragraph.Borders.LeftSize = 12;
                paragraph.Borders.RightStyle = BorderValues.Single;
                paragraph.Borders.RightColorHex = "0000ff";
                paragraph.Borders.RightSize = 20;

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false,
                    OfficeIMOPageSize = new OfficeIMO.Pdf.PageSize(360, 200),
                    OfficeIMOMargins = OfficeIMO.Pdf.PageMargins.Uniform(30)
                });
            }

            byte[] bytes = File.ReadAllBytes(pdfPath);
            string raw = Encoding.ASCII.GetString(bytes);
            using (PdfDocument pdf = PdfDocument.Open(bytes)) {
                string text = string.Concat(pdf.GetPages().Select(page => page.Text));

                Assert.Contains("Native side bordered", text);
                Assert.Contains("paragraph", text);
            }

            Assert.Contains("1 0 0 RG", raw);
            Assert.Contains("1.5 w", raw);
            Assert.Contains("0 0 1 RG", raw);
            Assert.Contains("2.5 w", raw);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Renders_Paragraph_Tab_Leaders() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeParagraphTabs.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeParagraphTabs.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                WordParagraph paragraph = document.AddParagraph("Revenue\t12");
                paragraph.AddTabStop(4320, TabStopValues.Right, TabStopLeaderCharValues.Dot);

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false,
                    OfficeIMOPageSize = new OfficeIMO.Pdf.PageSize(360, 180),
                    OfficeIMOMargins = PageMargins.Uniform(36)
                });
            }

            byte[] bytes = File.ReadAllBytes(pdfPath);
            using (PdfDocument pdf = PdfDocument.Open(bytes)) {
                var page = pdf.GetPage(1);
                string text = page.Text;
                Assert.Contains("Revenue", text);
                Assert.Contains("12", text);

                int dotCount = page.Letters.Count(letter => letter.Value == ".");
                Assert.True(dotCount >= 15, $"Expected Word tab stop leaders to render across the native paragraph tab gap. Dot count: {dotCount}.");
            }
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Does_Not_Leak_Run_Color_To_Following_Runs() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeRunColorReset.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeRunColorReset.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                WordParagraph paragraph = document.AddParagraph();
                paragraph.AddText("Before ");
                WordParagraph redRun = paragraph.AddText("Red");
                redRun.ColorHex = "ff0000";
                paragraph.AddText("After");

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            byte[] bytes = File.ReadAllBytes(pdfPath);
            string content = Encoding.ASCII.GetString(bytes);
            int redText = content.IndexOf("<526564>", StringComparison.Ordinal);
            int afterText = content.IndexOf("<4166746572>", StringComparison.Ordinal);

            Assert.True(redText >= 0, "Expected encoded 'Red' text in the generated PDF content stream.");
            Assert.True(afterText > redText, "Expected encoded 'After' text after the red Word run.");

            int redColorBeforeRed = content.LastIndexOf("1 0 0 rg", redText, StringComparison.Ordinal);
            int blackColorBeforeAfter = content.LastIndexOf("0 0 0 rg", afterText, StringComparison.Ordinal);
            int redColorBeforeAfter = content.LastIndexOf("1 0 0 rg", afterText, StringComparison.Ordinal);

            Assert.True(redColorBeforeRed >= 0, "Expected the Word run color to emit a red PDF fill color.");
            Assert.True(blackColorBeforeAfter > redText, "Expected the following uncolored Word run to reset to black/default PDF fill color.");
            Assert.True(redColorBeforeAfter < blackColorBeforeAfter, "Expected the following Word run not to inherit the previous red fill color.");
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Maps_Run_And_Paragraph_Font_Sizes() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeRunFontSizes.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeRunFontSizes.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("ParagraphSized").SetFontSize(16);

                WordParagraph paragraph = document.AddParagraph();
                paragraph.AddText("Small");
                paragraph.AddText("Large").SetFontSize(18);
                paragraph.AddText("Normal");

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            byte[] bytes = File.ReadAllBytes(pdfPath);
            string content = Encoding.ASCII.GetString(bytes);
            int paragraphText = content.IndexOf("<50617261677261706853697A6564>", StringComparison.Ordinal);
            int smallText = content.IndexOf("<536D616C6C>", StringComparison.Ordinal);
            int largeText = content.IndexOf("<4C61726765>", StringComparison.Ordinal);
            int normalText = content.IndexOf("<4E6F726D616C>", StringComparison.Ordinal);

            Assert.True(paragraphText >= 0, "Expected encoded 'ParagraphSized' text in the native PDF content stream.");
            Assert.True(smallText > paragraphText, "Expected encoded 'Small' text after the paragraph-sized paragraph.");
            Assert.True(largeText > smallText, "Expected encoded 'Large' text after the default-sized run.");
            Assert.True(normalText > largeText, "Expected encoded 'Normal' text after the large run.");

            int paragraphSizeBeforeParagraph = content.LastIndexOf("/F1 16 Tf", paragraphText, StringComparison.Ordinal);
            int defaultSizeBeforeSmall = content.LastIndexOf("/F1 11 Tf", smallText, StringComparison.Ordinal);
            int largeSizeBeforeLarge = content.LastIndexOf("/F1 18 Tf", largeText, StringComparison.Ordinal);
            int defaultSizeBeforeNormal = content.LastIndexOf("/F1 11 Tf", normalText, StringComparison.Ordinal);

            Assert.True(paragraphSizeBeforeParagraph >= 0, "Expected paragraph FontSize to emit a 16-point native PDF run.");
            Assert.True(defaultSizeBeforeSmall > paragraphText, "Expected the next paragraph to return to the default font size.");
            Assert.True(largeSizeBeforeLarge > smallText, "Expected the Word run font size to emit an 18-point native PDF run.");
            Assert.True(defaultSizeBeforeNormal > largeText, "Expected following Word runs not to inherit the previous run font size.");
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Maps_Run_Highlight_To_Background() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeRunHighlight.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeRunHighlight.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                WordParagraph paragraph = document.AddParagraph();
                paragraph.AddText("Before ");
                paragraph.AddText("Marked").SetHighlight(HighlightColorValues.Yellow);
                paragraph.AddText("After");

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            byte[] bytes = File.ReadAllBytes(pdfPath);
            string content = Encoding.ASCII.GetString(bytes);
            int markedText = content.IndexOf("<4D61726B6564>", StringComparison.Ordinal);
            int afterText = content.IndexOf("<4166746572>", StringComparison.Ordinal);
            using (PdfDocument pdf = PdfDocument.Open(bytes)) {
                string pageText = string.Concat(pdf.GetPages().Select(page => page.Text));

                Assert.Equal(1, CountOccurrences(pageText, "Before"));
                Assert.Equal(1, CountOccurrences(pageText, "Marked"));
                Assert.Equal(1, CountOccurrences(pageText, "After"));
            }

            Assert.True(markedText >= 0, "Expected encoded 'Marked' text in the native PDF content stream.");
            Assert.True(afterText > markedText, "Expected encoded 'After' text after the highlighted Word run.");

            int highlightFill = content.LastIndexOf("1 1 0 rg", markedText, StringComparison.Ordinal);
            int highlightRect = content.LastIndexOf(" re f", markedText, StringComparison.Ordinal);

            Assert.True(highlightFill >= 0, "Expected Word run highlight to emit a yellow PDF fill color.");
            Assert.True(highlightRect > highlightFill, "Expected Word run highlight to emit a filled rectangle behind the text.");
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Maps_Document_Background_Color() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeDocumentBackground.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeDocumentBackground.pdf");
            string marker = "WordBackgroundPdfMarker";

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.Background.SetColorHex("EAF4FF");
                document.AddParagraph(marker);

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            byte[] bytes = File.ReadAllBytes(pdfPath);
            string content = Encoding.ASCII.GetString(bytes);
            string markerHex = BitConverter.ToString(Encoding.ASCII.GetBytes(marker)).Replace("-", string.Empty);
            int backgroundFill = content.IndexOf("0.918 0.957 1 rg", StringComparison.Ordinal);
            int backgroundRect = backgroundFill < 0 ? -1 : content.IndexOf(" re f", backgroundFill, StringComparison.Ordinal);
            int markerText = content.IndexOf("<" + markerHex + ">", StringComparison.Ordinal);

            using (PdfDocument pdf = PdfDocument.Open(bytes)) {
                string pageText = string.Concat(pdf.GetPages().Select(page => page.Text));

                Assert.Contains(marker, pageText, StringComparison.Ordinal);
            }

            Assert.True(backgroundFill >= 0, "Expected Word document background color to emit a PDF page fill color.");
            Assert.True(backgroundRect > backgroundFill, "Expected Word document background color to emit a filled page rectangle.");
            Assert.True(markerText > backgroundRect, "Expected document background to render before Word paragraph text.");
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Maps_Table_Cell_Rich_Runs() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTableCellRichRuns.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTableCellRichRuns.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                WordTable table = document.AddTable(1, 1);
                WordParagraph paragraph = table.Rows[0].Cells[0].Paragraphs[0];
                paragraph.Text = string.Empty;
                paragraph.AddText("CellPlain ");
                WordParagraph red = paragraph.AddText("CellRed");
                red.ColorHex = "ff0000";
                paragraph.AddText(" ");
                paragraph.AddText("CellBold").SetBold();
                paragraph.AddText(" ");
                paragraph.AddText("CellMarked").SetHighlight(HighlightColorValues.Yellow);
                paragraph.AddText(" ");
                paragraph.AddText("CellLarge").SetFontSize(18);

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            byte[] bytes = File.ReadAllBytes(pdfPath);
            string content = Encoding.ASCII.GetString(bytes);
            int redText = content.IndexOf("<43656C6C526564>", StringComparison.Ordinal);
            int boldText = content.IndexOf("<43656C6C426F6C64>", StringComparison.Ordinal);
            int markedText = content.IndexOf("<43656C6C4D61726B6564>", StringComparison.Ordinal);
            int largeText = content.IndexOf("<43656C6C4C61726765>", StringComparison.Ordinal);

            using (PdfDocument pdf = PdfDocument.Open(bytes)) {
                string pageText = string.Concat(pdf.GetPages().Select(page => page.Text));

                Assert.Equal(1, CountOccurrences(pageText, "CellPlain"));
                Assert.Equal(1, CountOccurrences(pageText, "CellRed"));
                Assert.Equal(1, CountOccurrences(pageText, "CellBold"));
                Assert.Equal(1, CountOccurrences(pageText, "CellMarked"));
                Assert.Equal(1, CountOccurrences(pageText, "CellLarge"));
            }

            Assert.True(redText >= 0, "Expected encoded 'CellRed' text in the native table PDF content stream.");
            Assert.True(boldText > redText, "Expected encoded 'CellBold' text after the colored table cell run.");
            Assert.True(markedText > boldText, "Expected encoded 'CellMarked' text after the bold table cell run.");
            Assert.True(largeText > markedText, "Expected encoded 'CellLarge' text after the highlighted table cell run.");
            Assert.True(content.LastIndexOf("1 0 0 rg", redText, StringComparison.Ordinal) >= 0, "Expected Word table cell run color to emit a red PDF fill color.");
            Assert.True(content.LastIndexOf("/F2 11 Tf", boldText, StringComparison.Ordinal) >= 0, "Expected Word table cell bold run to use the bold PDF font resource.");
            Assert.True(content.LastIndexOf("1 1 0 rg", markedText, StringComparison.Ordinal) >= 0, "Expected Word table cell run highlight to emit a yellow PDF fill color.");
            Assert.True(content.LastIndexOf(" 18 Tf", largeText, StringComparison.Ordinal) >= 0, "Expected Word table cell run font size to emit an 18-point PDF run.");
        }

        private static int CountOccurrences(string value, string search) {
            int count = 0;
            int index = 0;
            while ((index = value.IndexOf(search, index, StringComparison.Ordinal)) >= 0) {
                count++;
                index += search.Length;
            }

            return count;
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Maps_Paragraph_Pagination_And_Tab_Style() {
            using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeParagraphStyle.docx"));
            WordParagraph paragraph = document.AddParagraph("Native style flags");
            paragraph.KeepLinesTogether = true;
            paragraph.KeepWithNext = true;
            paragraph.AvoidWidowAndOrphan = true;
            paragraph.AddTabStop(1440, TabStopValues.Right, TabStopLeaderCharValues.Dot);

            MethodInfo method = typeof(WordPdfConverterExtensions).GetMethod("CreateNativeParagraphStyle", BindingFlags.NonPublic | BindingFlags.Static)!;
            PdfParagraphStyle style = Assert.IsType<PdfParagraphStyle>(method.Invoke(null, new object[] { paragraph }));

            Assert.Equal(72D, style.DefaultTabStopWidth);
            Assert.Equal(1.15D, style.LineHeight);
            Assert.Equal(8D, style.SpacingAfter);
            Assert.True(style.KeepTogether);
            Assert.True(style.KeepWithNext);
            Assert.True(style.WidowControl);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_Maps_Explicit_Paragraph_Line_Spacing() {
            using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeExplicitParagraphStyle.docx"));
            WordParagraph exactParagraph = document.AddParagraph("Native exact line spacing");
            exactParagraph.FontSize = 12;
            exactParagraph.LineSpacingPoints = 24;

            WordParagraph autoParagraph = document.AddParagraph("Native auto line spacing");
            autoParagraph.LineSpacing = 276;
            autoParagraph.LineSpacingRule = LineSpacingRuleValues.Auto;

            MethodInfo method = typeof(WordPdfConverterExtensions).GetMethod("CreateNativeParagraphStyle", BindingFlags.NonPublic | BindingFlags.Static)!;
            PdfParagraphStyle exactStyle = Assert.IsType<PdfParagraphStyle>(method.Invoke(null, new object[] { exactParagraph }));
            PdfParagraphStyle autoStyle = Assert.IsType<PdfParagraphStyle>(method.Invoke(null, new object[] { autoParagraph }));

            Assert.Equal(2D, exactStyle.LineHeight);
            Assert.Equal(1.15D, autoStyle.LineHeight);
        }
    }
}
