using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using W15 = DocumentFormat.OpenXml.Office2013.Word;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_MergeKeepDifferentStylesCopiesClashingHeadingUnderNewId() {
            var destinationPath = Path.Combine(_directoryWithFiles, "MergeKeepDifferentBase.docx");
            var sourcePath = Path.Combine(_directoryWithFiles, "MergeKeepDifferentSource.docx");
            CreateHeadingDocument(destinationPath, "Destination heading", "FF0000");
            CreateHeadingDocument(sourcePath, "Source heading", "0000FF");

            using (var destination = WordDocument.Load(destinationPath))
            using (var source = WordDocument.Load(sourcePath)) {
                destination.AppendDocument(source, new WordDocumentMergeOptions { StyleMode = WordMergeStyleMode.KeepDifferentStyles });
                destination.Save();
            }

            using (var merged = WordDocument.Load(destinationPath)) {
                var styles = merged._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Elements<Style>().ToList();
                Style original = styles.Single(_ => _.StyleId == "Heading1");
                Style copy = styles.Single(_ => _.StyleId == "Heading1_0");

                Assert.Equal("FF0000", original.StyleRunProperties!.Color!.Val!.Value);
                Assert.Equal("0000FF", copy.StyleRunProperties!.Color!.Val!.Value);
                Assert.Equal(original.StyleName!.Val!.Value + "_0", copy.StyleName!.Val!.Value);
                Assert.Contains(styles, _ => _.StyleId == copy.BasedOn?.Val?.Value);
                Assert.Single(styles, _ => _.StyleName?.Val?.Value == original.StyleName!.Val!.Value);

                Assert.Equal("Heading1", merged.Paragraphs.Single(_ => _.Text == "Destination heading").StyleId);
                Assert.Equal("Heading1_0", merged.Paragraphs.Single(_ => _.Text == "Source heading").StyleId);
                AssertMergedDocumentIsValid(merged);
            }
        }

        [Fact]
        public void Test_MergeKeepDifferentStylesSharesIdenticalStyles() {
            var destinationPath = Path.Combine(_directoryWithFiles, "MergeKeepDifferentSameBase.docx");
            var sourcePath = Path.Combine(_directoryWithFiles, "MergeKeepDifferentSameSource.docx");
            CreateHeadingDocument(destinationPath, "Destination heading", "FF0000");
            CreateHeadingDocument(sourcePath, "Source heading", "FF0000");

            using (var destination = WordDocument.Load(destinationPath))
            using (var source = WordDocument.Load(sourcePath)) {
                destination.AppendDocument(source, new WordDocumentMergeOptions { StyleMode = WordMergeStyleMode.KeepDifferentStyles });
                destination.Save();
            }

            using (var merged = WordDocument.Load(destinationPath)) {
                var styles = merged._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Elements<Style>().ToList();
                Assert.DoesNotContain(styles, _ => _.StyleId?.Value?.EndsWith("_0", StringComparison.Ordinal) == true);
                Assert.Equal("Heading1", merged.Paragraphs.Single(_ => _.Text == "Source heading").StyleId);
            }
        }

        [Fact]
        public void Test_MergeUseDestinationStylesKeepsDestinationDefinition() {
            var destinationPath = Path.Combine(_directoryWithFiles, "MergeUseDestinationBase.docx");
            var sourcePath = Path.Combine(_directoryWithFiles, "MergeUseDestinationSource.docx");
            CreateHeadingDocument(destinationPath, "Destination heading", "FF0000");
            CreateHeadingDocument(sourcePath, "Source heading", "0000FF");

            using (var destination = WordDocument.Load(destinationPath))
            using (var source = WordDocument.Load(sourcePath)) {
                destination.AppendDocument(source);
                destination.Save();
            }

            using (var merged = WordDocument.Load(destinationPath)) {
                var styles = merged._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Elements<Style>().ToList();
                Assert.Equal("FF0000", styles.Single(_ => _.StyleId == "Heading1").StyleRunProperties!.Color!.Val!.Value);
                Assert.DoesNotContain(styles, _ => _.StyleId == "Heading1_0");
                Assert.Equal("Heading1", merged.Paragraphs.Single(_ => _.Text == "Source heading").StyleId);
            }
        }

        [Fact]
        public void Test_MergeKeepsPageSetupPerSection() {
            var destinationPath = Path.Combine(_directoryWithFiles, "MergeSectionsBase.docx");
            var sourcePath = Path.Combine(_directoryWithFiles, "MergeSectionsSource.docx");

            using (var document = WordDocument.Create(destinationPath)) {
                document.AddParagraph("Portrait content");
                document.Save();
            }

            using (var document = WordDocument.Create(sourcePath)) {
                document.Sections[0].PageOrientation = OfficePageOrientation.Landscape;
                document.AddParagraph("Landscape content");
                document.Save();
            }

            using (var destination = WordDocument.Load(destinationPath))
            using (var source = WordDocument.Load(sourcePath)) {
                destination.AppendDocument(source);
                Assert.Equal(2, destination.Sections.Count);
                destination.Save();
            }

            using (var merged = WordDocument.Load(destinationPath)) {
                Body body = merged._document.Body!;
                Assert.IsType<SectionProperties>(body.LastChild);
                Assert.Single(body.Elements<SectionProperties>());
                Assert.Equal(2, merged.Sections.Count);
                Assert.Equal(OfficePageOrientation.Portrait, merged.Sections[0].PageSettings.Orientation);
                Assert.Equal(OfficePageOrientation.Landscape, merged.Sections[1].PageSettings.Orientation);
                Assert.Contains(merged.Sections[0].Paragraphs, _ => _.Text == "Portrait content");
                Assert.Contains(merged.Sections[1].Paragraphs, _ => _.Text == "Landscape content");
                AssertMergedDocumentIsValid(merged);
            }
        }

        [Fact]
        public void Test_MergeUseDestinationSectionsKeepsDestinationPageSetup() {
            var destinationPath = Path.Combine(_directoryWithFiles, "MergeDestinationSectionsBase.docx");
            var sourcePath = Path.Combine(_directoryWithFiles, "MergeDestinationSectionsSource.docx");

            using (var document = WordDocument.Create(destinationPath)) {
                document.AddParagraph("Portrait content");
                document.Save();
            }

            using (var document = WordDocument.Create(sourcePath)) {
                document.Sections[0].PageOrientation = OfficePageOrientation.Landscape;
                document.AddParagraph("First source section");
                document.AddSection();
                document.AddParagraph("Second source section");
                document.Save();
            }

            using (var destination = WordDocument.Load(destinationPath))
            using (var source = WordDocument.Load(sourcePath)) {
                destination.AppendDocument(source, new WordDocumentMergeOptions { SectionMode = WordMergeSectionMode.UseDestinationSections });
                destination.Save();
            }

            using (var merged = WordDocument.Load(destinationPath)) {
                Body body = merged._document.Body!;
                Assert.IsType<SectionProperties>(body.LastChild);
                Assert.Single(body.Descendants<SectionProperties>());
                Assert.Single(merged.Sections);
                Assert.Equal(OfficePageOrientation.Portrait, merged.Sections[0].PageSettings.Orientation);
                Assert.Equal(new[] { "Portrait content", "First source section", "Second source section" },
                    merged.Paragraphs.Where(_ => !string.IsNullOrEmpty(_.Text)).Select(_ => _.Text).ToArray());
            }
        }

        [Fact]
        public void Test_MergeCarriesCommentsFootnotesAndEndnotes() {
            var destinationPath = Path.Combine(_directoryWithFiles, "MergeStoriesBase.docx");
            var sourcePath = Path.Combine(_directoryWithFiles, "MergeStoriesSource.docx");

            using (var document = WordDocument.Create(destinationPath)) {
                document.AddParagraph("Destination text").AddComment("Dest", "D", "Destination comment");
                document.AddParagraph("Destination note").AddFootNote("Destination footnote");
                document.Save();
            }

            using (var document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("Source text").AddComment("Source", "S", "Source comment");
                document.Comments.Single(_ => _.Text == "Source comment").AddReply("Reviewer", "R", "Source reply");
                document.AddParagraph("Source note").AddFootNote("Source footnote");
                document.AddParagraph("Source end").AddEndNote("Source endnote");
                document.Save();
            }

            using (var destination = WordDocument.Load(destinationPath))
            using (var source = WordDocument.Load(sourcePath)) {
                destination.AppendDocument(source);
                destination.Save();
            }

            using (var merged = WordDocument.Load(destinationPath)) {
                MainDocumentPart main = merged._wordprocessingDocument.MainDocumentPart!;
                Body body = main.Document!.Body!;

                var comments = main.WordprocessingCommentsPart!.Comments!.Elements<Comment>().ToList();
                Assert.Equal(3, comments.Count);
                Assert.Equal(comments.Count, comments.Select(_ => _.Id!.Value).Distinct().Count());
                foreach (CommentReference reference in body.Descendants<CommentReference>()) {
                    Assert.Contains(comments, _ => _.Id!.Value == reference.Id!.Value);
                }

                Assert.Contains(merged.Comments, _ => _.Text == "Destination comment");
                WordComment sourceComment = merged.Comments.Single(_ => _.Text == "Source comment");
                WordComment reply = merged.Comments.Single(_ => _.Text == "Source reply");
                Assert.Equal(sourceComment.ParaId, reply.ParentComment?.ParaId);

                var footnotes = main.FootnotesPart!.Footnotes!.Elements<Footnote>().ToList();
                var footnoteReferences = body.Descendants<FootnoteReference>().ToList();
                Assert.Equal(2, footnoteReferences.Count);
                Assert.Equal(2, footnoteReferences.Select(_ => _.Id!.Value).Distinct().Count());
                Assert.Contains("Source footnote", footnotes.Single(_ => _.Id!.Value == footnoteReferences[1].Id!.Value).InnerText);

                EndnoteReference endnoteReference = Assert.Single(body.Descendants<EndnoteReference>());
                Assert.Contains("Source endnote", main.EndnotesPart!.Endnotes!.Elements<Endnote>().Single(_ => _.Id!.Value == endnoteReference.Id!.Value).InnerText);
                AssertMergedDocumentIsValid(merged);
            }
        }

        [Fact]
        public void Test_MergeRenumbersDrawingAndBookmarkIds() {
            var destinationPath = Path.Combine(_directoryWithFiles, "MergeIdsBase.docx");
            var sourcePath = Path.Combine(_directoryWithFiles, "MergeIdsSource.docx");
            string imagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");

            foreach (var (path, text) in new[] { (destinationPath, "Destination"), (sourcePath, "Source") }) {
                using var document = WordDocument.Create(path);
                document.AddParagraph(text + " intro").AddBookmark("Intro");
                document.AddParagraph(text + " image").AddImage(imagePath, 32, 32);
                document.AddParagraph("Jump").AddHyperLink(text + " link", "Intro");
                document.AddParagraph(text + " list").AddFootNote(text + " footnote");
                document.Save();
            }

            using (var destination = WordDocument.Load(destinationPath))
            using (var source = WordDocument.Load(sourcePath)) {
                destination.AppendDocument(source);
                destination.Save();
            }

            using (var merged = WordDocument.Load(destinationPath)) {
                Body body = merged._document.Body!;
                var drawingIds = body.Descendants<DW.DocProperties>().Select(_ => _.Id!.Value).ToList();
                Assert.Equal(2, drawingIds.Count);
                Assert.Equal(drawingIds.Count, drawingIds.Distinct().Count());

                var bookmarks = body.Descendants<BookmarkStart>().ToList();
                Assert.Equal(bookmarks.Count, bookmarks.Select(_ => _.Id!.Value).Distinct().Count());
                Assert.Equal(bookmarks.Count, bookmarks.Select(_ => _.Name!.Value).Distinct(StringComparer.OrdinalIgnoreCase).Count());
                foreach (BookmarkEnd end in body.Descendants<BookmarkEnd>()) {
                    Assert.Contains(bookmarks, _ => _.Id!.Value == end.Id!.Value);
                }

                BookmarkStart renamed = bookmarks.Single(_ => _.Name == "Intro_1");
                Assert.Contains("Source intro", ((Paragraph)renamed.Parent!).InnerText);
                Hyperlink sourceLink = body.Descendants<Hyperlink>().Single(_ => _.InnerText == "Source link");
                Assert.Equal("Intro_1", sourceLink.Anchor!.Value);
                Assert.Equal("Intro", body.Descendants<Hyperlink>().Single(_ => _.InnerText == "Destination link").Anchor!.Value);
                AssertMergedDocumentIsValid(merged);
            }
        }

        [Fact]
        public void Test_MergeKeepsNumberingSchemaOrder() {
            var destinationPath = Path.Combine(_directoryWithFiles, "MergeNumberingOrderBase.docx");
            var sourcePath = Path.Combine(_directoryWithFiles, "MergeNumberingOrderSource.docx");

            foreach (string path in new[] { destinationPath, sourcePath }) {
                using var document = WordDocument.Create(path);
                var list = document.AddList(WordListStyle.Numbered);
                list.AddItem("One");
                list.AddItem("Two");
                document.Save();
            }

            using (var destination = WordDocument.Load(destinationPath))
            using (var source = WordDocument.Load(sourcePath)) {
                destination.AppendDocument(source);
                destination.Save();
            }

            using (var merged = WordDocument.Load(destinationPath)) {
                Numbering numbering = merged._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!;
                var children = numbering.ChildElements.ToList();
                int lastAbstract = children.FindLastIndex(_ => _ is AbstractNum);
                int firstInstance = children.FindIndex(_ => _ is NumberingInstance);
                Assert.True(lastAbstract < firstInstance, "Every w:abstractNum must precede the first w:num.");
                var abstractIds = numbering.Elements<AbstractNum>().Select(_ => _.AbstractNumberId!.Value).ToList();
                Assert.Equal(abstractIds.Count, abstractIds.Distinct().Count());
                var nsids = numbering.Elements<AbstractNum>().Select(_ => _.Nsid?.Val?.Value).Where(_ => _ != null).ToList();
                Assert.Equal(nsids.Count, nsids.Distinct().Count());
                AssertMergedDocumentIsValid(merged);
            }
        }

        private static void CreateHeadingDocument(string path, string headingText, string color) {
            using (var document = WordDocument.Create(path)) {
                document.AddParagraph(headingText).Style = WordParagraphStyles.Heading1;
                document.AddParagraph("Body text under " + headingText);
                document.Save();
            }

            using (var document = WordDocument.Load(path)) {
                Style heading = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                    .Elements<Style>().Single(_ => _.StyleId == "Heading1");
                heading.StyleRunProperties ??= new StyleRunProperties();
                heading.StyleRunProperties.Color = new Color { Val = color };
                document.Save();
            }
        }

        private static void AssertMergedDocumentIsValid(WordDocument document) {
            var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(document._wordprocessingDocument).ToList();
            Assert.True(errors.Count == 0, string.Join(Environment.NewLine, errors.Select(_ => _.Id + ": " + _.Description + " @ " + _.Path?.XPath + " in " + _.Part?.Uri)));
        }
    }
}
