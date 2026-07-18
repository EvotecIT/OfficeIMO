using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointPresentation {
        private static void ProjectLegacyComments(PowerPointPresentation presentation,
            LegacyPptPresentation legacy) {
            LegacyPptComment[] comments = legacy.Slides.SelectMany(slide => slide.Comments).ToArray();
            if (comments.Length == 0) return;

            PresentationPart presentationPart = presentation._presentationPart;
            CommentAuthorsPart authorsPart = presentationPart.AddNewPart<CommentAuthorsPart>();
            var authors = new Dictionary<string, ProjectedLegacyCommentAuthor>(StringComparer.Ordinal);
            var authorOrder = new List<ProjectedLegacyCommentAuthor>();
            foreach (LegacyPptComment comment in comments) {
                string key = comment.Author + "\0" + comment.Initials;
                if (!authors.TryGetValue(key, out ProjectedLegacyCommentAuthor? author)) {
                    uint id = checked((uint)authorOrder.Count);
                    author = new ProjectedLegacyCommentAuthor(id, comment.Author, comment.Initials);
                    authors.Add(key, author);
                    authorOrder.Add(author);
                }
                author.LastIndex = Math.Max(author.LastIndex, checked((uint)comment.Index));
            }
            authorsPart.CommentAuthorList = new P.CommentAuthorList(authorOrder.Select(author =>
                new P.CommentAuthor {
                    Id = author.Id,
                    Name = author.Name,
                    Initials = author.Initials,
                    LastIndex = author.LastIndex,
                    ColorIndex = author.Id
                }));

            for (int slideIndex = 0; slideIndex < legacy.Slides.Count; slideIndex++) {
                LegacyPptSlide sourceSlide = legacy.Slides[slideIndex];
                if (sourceSlide.Comments.Count == 0) continue;
                PowerPointSlide targetSlide = presentation.Slides[slideIndex];
                SlideCommentsPart commentsPart = targetSlide.SlidePart.AddNewPart<SlideCommentsPart>();
                var commentList = new P.CommentList();
                foreach (LegacyPptComment comment in sourceSlide.Comments) {
                    ProjectedLegacyCommentAuthor author = authors[
                        comment.Author + "\0" + comment.Initials];
                    var projected = new P.Comment(
                        new P.Position { X = comment.X, Y = comment.Y },
                        new P.Text(comment.Text)) {
                        AuthorId = author.Id,
                        Index = checked((uint)comment.Index)
                    };
                    if (comment.CreatedAtUtc.HasValue) {
                        projected.DateTime = comment.CreatedAtUtc.Value;
                    }
                    commentList.Append(projected);
                }
                commentsPart.CommentList = commentList;
            }
        }

        private sealed class ProjectedLegacyCommentAuthor {
            internal ProjectedLegacyCommentAuthor(uint id, string name, string initials) {
                Id = id;
                Name = name;
                Initials = initials;
            }

            internal uint Id { get; }
            internal string Name { get; }
            internal string Initials { get; }
            internal uint LastIndex { get; set; }
        }
    }
}
