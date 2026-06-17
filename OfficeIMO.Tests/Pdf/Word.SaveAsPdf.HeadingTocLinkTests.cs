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
            using (PdfPigDocument pdf = PdfPigDocument.Open(bytes)) {
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
        public void NativeHeadingDestinationNames_TrackSuffixes_ForRepeatedHeadings() {
            MethodInfo method = typeof(WordPdfConverterExtensions).GetMethod("CreateNativeHeadingDestinationName", BindingFlags.NonPublic | BindingFlags.Static)!;
            var used = new HashSet<string>(StringComparer.Ordinal);
            var nextSuffixByBaseName = new Dictionary<string, int>(StringComparer.Ordinal);
            const string headingText = "Repeated native heading";
            const string baseDestination = "officeimo-heading-repeated-native-heading";

            for (int index = 1; index <= 256; index++) {
                string destinationName = (string)method.Invoke(null, new object[] {
                    headingText,
                    index,
                    used,
                    nextSuffixByBaseName
                })!;

                string expectedName = index == 1
                    ? baseDestination
                    : baseDestination + "-" + index.ToString(System.Globalization.CultureInfo.InvariantCulture);
                Assert.Equal(expectedName, destinationName);
                used.Add(destinationName);
            }

            Assert.Equal(257, nextSuffixByBaseName[baseDestination]);

            used.Add(baseDestination + "-257");
            string skippedCollision = (string)method.Invoke(null, new object[] {
                headingText,
                257,
                used,
                nextSuffixByBaseName
            })!;

            Assert.Equal(baseDestination + "-258", skippedCollision);
            Assert.Equal(259, nextSuffixByBaseName[baseDestination]);
        }

        [Fact]
        public void SaveAsPdf_OfficeIMOEngine_TableOfContents_Accounts_For_Section_Page_Starts() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTableOfContentsSections.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTableOfContentsSections.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddTableOfContent();
                document.AddParagraph("Native TOC first section heading").SetStyle(WordParagraphStyles.Heading1);
                document.AddParagraph("Native TOC first section body");
                WordSection secondSection = document.AddSection();
                secondSection.AddParagraph("Native TOC second section heading").SetStyle(WordParagraphStyles.Heading1);
                secondSection.AddParagraph("Native TOC second section body");

                MethodInfo method = typeof(WordPdfConverterExtensions).GetMethod("BuildNativeTableOfContentsEntries", BindingFlags.NonPublic | BindingFlags.Static)!;
                object entries = method.Invoke(null, new object[] {
                    document,
                    new PdfSaveOptions { IncludePageNumbers = false },
                    new Dictionary<DocumentFormat.OpenXml.Wordprocessing.Paragraph, string>()
                })!;
                object secondEntry = ((System.Collections.IEnumerable)entries)
                    .Cast<object>()
                    .First(entry => string.Equals((string)entry.GetType().GetProperty("Text")!.GetValue(entry)!, "Native TOC second section heading", StringComparison.Ordinal));
                int secondEntryPage = (int)secondEntry.GetType().GetProperty("PageNumber")!.GetValue(secondEntry)!;
                Assert.Equal(2, secondEntryPage);

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false
                });
            }

            using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
            Assert.True(pdf.NumberOfPages >= 2, "Expected the second Word section to start on a new PDF page.");
            string secondPageText = pdf.GetPage(2).Text;
            Assert.Contains("Native TOC second section heading", secondPageText);
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
    }
}
