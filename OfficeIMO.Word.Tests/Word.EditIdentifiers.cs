using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2013.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void BookmarkIdsStayUniqueAcrossBodyAndHeaders() {
            using var stream = new MemoryStream();
            using (WordDocument document = WordDocument.Create(stream)) {
                document.Sections[0].PageSetup.Header.AddParagraph("Header").AddBookmark("HeaderMark");
                for (int i = 0; i < 5; i++) {
                    document.AddParagraph($"Body {i}").AddBookmark($"Body{i}");
                }

                document.Save();
            }

            stream.Position = 0;
            using WordprocessingDocument package = WordprocessingDocument.Open(stream, false);
            MainDocumentPart main = package.MainDocumentPart!;
            BookmarkStart[] starts = main.Document.Body!.Descendants<BookmarkStart>()
                .Concat(main.HeaderParts.SelectMany(part => part.Header.Descendants<BookmarkStart>()))
                .ToArray();
            Assert.Equal(6, starts.Length);
            Assert.Equal(starts.Length, starts.Select(start => start.Id!.Value).Distinct().Count());
            string[] endIds = main.Document.Body!.Descendants<BookmarkEnd>()
                .Concat(main.HeaderParts.SelectMany(part => part.Header.Descendants<BookmarkEnd>()))
                .Select(end => end.Id!.Value!).ToArray();
            Assert.Equal(starts.Select(start => start.Id!.Value).OrderBy(id => id), endIds.OrderBy(id => id));
        }

        [Fact]
        public void CommentsAddedWithoutIntermediateSavesArePersistedWithDistinctIdentities() {
            using var created = new MemoryStream();
            using var copy = new MemoryStream();
            using (WordDocument document = WordDocument.Create(created)) {
                WordComment? first = null;
                for (int i = 0; i < 4; i++) {
                    WordComment comment = document.AddParagraph($"Paragraph {i}").AddComment("Author", "AU", $"Comment {i}");
                    first ??= comment;
                }

                first!.AddReply("Reviewer", "RV", "Reply");
                document.Save();
                document.Save(copy);
            }

            foreach (MemoryStream stream in new[] { created, copy }) {
                stream.Position = 0;
                using (WordDocument loaded = WordDocument.Load(stream)) {
                    Assert.Equal(new[] { "Comment 0", "Comment 1", "Comment 2", "Comment 3", "Reply" },
                        loaded.Comments.Select(comment => comment.Text).ToArray());
                }

                stream.Position = 0;
                using WordprocessingDocument package = WordprocessingDocument.Open(stream, false);
                Comment[] comments = package.MainDocumentPart!.WordprocessingCommentsPart!.Comments.Elements<Comment>().ToArray();
                CommentEx[] extensions = package.MainDocumentPart.WordprocessingCommentsExPart!.CommentsEx.Elements<CommentEx>().ToArray();
                Assert.Equal(comments.Length, comments.Select(comment => comment.Id!.Value).Distinct().Count());
                Assert.Equal(comments.Length, extensions.Select(extension => extension.ParaId!.Value).Distinct().Count());
                Assert.Single(extensions, extension => extension.ParaIdParent != null);
            }
        }
        [Fact]
        public void IdentifiersStayUniqueWhenALoadedDocumentIsExtended() {
            using var stream = new MemoryStream();
            using (WordDocument document = WordDocument.Create(stream)) {
                AddIdentifiedItems(document, "First", 3);
                document.Sections[0].PageSetup.Header.AddParagraph("Header").AddBookmark("HeaderMark");
                document.Save();
            }

            stream.Position = 0;
            using (WordDocument loaded = WordDocument.Load(stream)) {
                AddIdentifiedItems(loaded, "Second", 3);
                loaded.Comments[0].AddReply("Reviewer", "RV", "Reply");
                loaded.Save();
            }

            AssertUniqueIdentifiers(stream, bookmarks: 7, drawings: 6, comments: 7);
        }

        [Fact]
        public void IdentifiersStayUniqueWhenItemsAreRemovedBetweenAdds() {
            using var stream = new MemoryStream();
            using (WordDocument document = WordDocument.Create(stream)) {
                AddIdentifiedItems(document, "First", 3);
                document.Bookmarks[2].Remove();
                document.Bookmarks[0].Remove();
                document.Comments[2].Remove();
                document.Comments[0].Remove();
                document.Images[2].Remove();
                AddIdentifiedItems(document, "Second", 3);
                document.Save();
            }

            AssertUniqueIdentifiers(stream, bookmarks: 4, drawings: 5, comments: 4);
        }

        [Fact]
        public void IdentifiersStayUniqueAcrossMergedDocumentsAndOtherAuthoringApis() {
            using var stream = new MemoryStream();
            using (WordDocument document = WordDocument.Create(stream)) {
                AddIdentifiedItems(document, "First", 2);
                using (WordDocument source = WordDocument.Create()) {
                    AddIdentifiedItems(source, "Source", 4);
                    document.AppendDocument(source);
                }

                WordParagraph paragraph = document.AddParagraph("Range text and more");
                paragraph.GetRange(0, 10).AddBookmark("RangeMark");
                paragraph.GetRange(11, 8).AddComment("Author", "AU", "Range comment");
                document.AddShapeDrawing(WordShapeType.Rectangle, OfficeLength.Points(20), OfficeLength.Points(20));
                document.Sections[0].PageSetup.Header.AddShapeDrawing(WordShapeType.Ellipse, OfficeLength.Points(20), OfficeLength.Points(20));
                document.AddParagraph().AddTextBox("Box", WordImageTextWrapping.Square);
                AddIdentifiedItems(document, "Second", 2);
                document.Save();
            }

            AssertUniqueIdentifiers(stream, bookmarks: 9, drawings: 11, comments: 9);
        }

        [Fact]
        public void IdentifiersWrittenDirectlyToThePackageAreNotReused() {
            using var stream = new MemoryStream();
            using (WordDocument document = WordDocument.Create(stream)) {
                // Two of each item take bookmark and comment ids 1-2 and drawing ids 1-4; the raw items below
                // use the next ids the document would otherwise hand out.
                AddIdentifiedItems(document, "First", 2);

                // Comments appended outside the API are seen on the next add.
                MainDocumentPart main = document._wordprocessingDocument.MainDocumentPart!;
                main.WordprocessingCommentsPart!.Comments.AppendChild(new Comment(new Paragraph(new Run(new Text("Raw")))) { Id = "3", Author = "Raw" });
                document.AddParagraph("Commented").AddComment("Author", "AU", "After raw");

                // Bookmarks and drawings written directly are seen once the document has been saved.
                Paragraph raw = main.Document.Body!.AppendChild(new Paragraph(new BookmarkStart { Id = "3", Name = "Raw" }, new BookmarkEnd { Id = "3" }));
                Paragraph imageParagraph = document.Paragraphs.First(p => p.IsImage)._paragraph;
                Paragraph rawImage = (Paragraph)imageParagraph.CloneNode(true);
                rawImage.Descendants<DW.DocProperties>().Single().Id = 5U;
                raw.InsertAfterSelf(rawImage);
                document.Save();

                document.AddParagraph("After").AddBookmark("AfterRaw");
                document.AddParagraph().AddImage(Path.Combine(_directoryWithImages, "EvotecLogo.png"));
                document.Save();
            }

            AssertUniqueIdentifiers(stream, bookmarks: 4, drawings: 4, comments: 4);
        }

        private void AddIdentifiedItems(WordDocument document, string prefix, int count) {
            for (int i = 0; i < count; i++) {
                document.AddParagraph($"{prefix} bookmark {i}").AddBookmark($"{prefix}Mark{i}");
                document.AddParagraph($"{prefix} comment {i}").AddComment("Author", "AU", $"{prefix} comment {i}");
                document.AddParagraph().AddImage(Path.Combine(_directoryWithImages, "EvotecLogo.png"));
            }
        }

        private static void AssertUniqueIdentifiers(MemoryStream stream, int bookmarks, int drawings, int comments) {
            stream.Position = 0;
            using WordprocessingDocument package = WordprocessingDocument.Open(stream, false);
            MainDocumentPart main = package.MainDocumentPart!;
            var roots = new List<OpenXmlElement> { main.Document.Body! };
            roots.AddRange(main.HeaderParts.Select(part => part.Header));
            roots.AddRange(main.FooterParts.Select(part => part.Footer));

            string[] bookmarkIds = roots.SelectMany(root => root.Descendants<BookmarkStart>())
                .Where(start => start.Name?.Value != "_GoBack").Select(start => start.Id!.Value!).ToArray();
            Assert.Equal(bookmarks, bookmarkIds.Length);
            Assert.Equal(bookmarkIds.Length, bookmarkIds.Distinct().Count());

            uint[] drawingIds = roots.SelectMany(root => root.Descendants<DW.DocProperties>()).Select(properties => properties.Id!.Value).ToArray();
            Assert.Equal(drawings, drawingIds.Length);
            Assert.Equal(drawingIds.Length, drawingIds.Distinct().Count());

            Comment[] commentElements = main.WordprocessingCommentsPart!.Comments.Elements<Comment>().ToArray();
            Assert.Equal(comments, commentElements.Length);
            Assert.Equal(comments, commentElements.Select(comment => comment.Id!.Value).Distinct().Count());
            string[] paragraphIds = commentElements.SelectMany(comment => comment.Elements<Paragraph>())
                .Select(paragraph => paragraph.ParagraphId?.Value).Where(id => id != null).Select(id => id!).ToArray();
            Assert.Equal(paragraphIds.Length, paragraphIds.Distinct().Count());
            string[] extensionIds = main.WordprocessingCommentsExPart!.CommentsEx.Elements<CommentEx>().Select(extension => extension.ParaId!.Value!).ToArray();
            Assert.Equal(extensionIds.Length, extensionIds.Distinct().Count());
        }
    }
}
