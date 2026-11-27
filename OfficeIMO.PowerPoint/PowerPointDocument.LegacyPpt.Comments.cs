using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        private static void ProjectLegacyComments(LegacyPptImportParts parts,
            LegacyPptDocument legacy,
            IReadOnlyList<PowerPointPartRef> slideParts) {
            LegacyPptComment[] comments = legacy.Slides.SelectMany(slide => slide.Comments).ToArray();
            if (comments.Length == 0) return;

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
            parts.AddXmlPart(parts.PresentationPart, LegacyCommentAuthorsRelationshipType, "/ppt/commentAuthors.xml",
                LegacyCommentAuthorsContentType, LegacyPptXml.P, "cmAuthorLst", list => {
                    foreach (ProjectedLegacyCommentAuthor author in authorOrder) {
                        list.AddP("cmAuthor")
                            .Set("id", (long)author.Id)
                            .Set("name", author.Name)
                            .Set("initials", author.Initials)
                            .Set("lastIdx", (long)author.LastIndex)
                            .Set("clrIdx", (long)author.Id);
                    }
                }, out _, numbered: false);

            for (int slideIndex = 0; slideIndex < legacy.Slides.Count; slideIndex++) {
                LegacyPptSlide sourceSlide = legacy.Slides[slideIndex];
                if (sourceSlide.Comments.Count == 0) continue;
                PowerPointPartRef targetSlide = slideParts[slideIndex];
                parts.AddXmlPart(targetSlide, LegacyCommentsRelationshipType, "/ppt/comments/comment.xml",
                    LegacyCommentsContentType, LegacyPptXml.P, "cmLst", list => {
                        foreach (LegacyPptComment comment in sourceSlide.Comments) {
                            ProjectedLegacyCommentAuthor author = authors[
                                comment.Author + "\0" + comment.Initials];
                            OfficeXmlNode projected = list.AddP("cm")
                                .Set("authorId", (long)author.Id);
                            if (comment.CreatedAtUtc.HasValue) {
                                projected.Set("dt", FormatLegacyCommentDate(comment.CreatedAtUtc.Value));
                            }
                            projected.Set("idx", (long)checked((uint)comment.Index));
                            projected.AddP("pos").Set("x", (long)comment.X).Set("y", (long)comment.Y);
                            projected.AddP("text").SetInnerText(comment.Text);
                        }
                    }, out _);
            }
        }

        // xsd:dateTime as the SDK writes a DateTimeValue.
        private static string FormatLegacyCommentDate(DateTime value) =>
            System.Xml.XmlConvert.ToString(value, System.Xml.XmlDateTimeSerializationMode.RoundtripKind);

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
