using DocumentFormat.OpenXml.Office2013.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

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
    }
}
