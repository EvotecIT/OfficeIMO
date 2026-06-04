using AngleSharp.Dom;
using System.Globalization;
using System.Threading;

namespace OfficeIMO.Word.Html {
    internal partial class HtmlToWordConverter {
        private sealed class HtmlCommentInfo {
            public string Text { get; init; } = string.Empty;
            public string Author { get; init; } = string.Empty;
            public string Initials { get; init; } = string.Empty;
            public DateTime? Date { get; init; }
            public List<HtmlCommentInfo> Replies { get; } = new();
        }

        private void CaptureCommentSections(IDocument document, CancellationToken cancellationToken = default) {
            var commentSection = document.QuerySelector("section.comments");
            if (commentSection == null) {
                return;
            }

            var rootList = commentSection.Children.FirstOrDefault(element => string.Equals(element.TagName, "OL", StringComparison.OrdinalIgnoreCase));
            if (rootList != null) {
                foreach (var item in rootList.Children.Where(element => string.Equals(element.TagName, "LI", StringComparison.OrdinalIgnoreCase))) {
                    cancellationToken.ThrowIfCancellationRequested();
                    var id = item.GetAttribute("id");
                    if (!string.IsNullOrEmpty(id)) {
                        _commentMap[id!] = CaptureCommentListItem(item, cancellationToken);
                    }
                }
            }

            commentSection.Remove();
        }

        private static HtmlCommentInfo CaptureCommentListItem(IElement item, CancellationToken cancellationToken) {
            var textParagraph = item.Children.FirstOrDefault(element => string.Equals(element.TagName, "P", StringComparison.OrdinalIgnoreCase));
            var comment = new HtmlCommentInfo {
                Text = textParagraph?.TextContent?.Trim() ?? item.TextContent?.Trim() ?? string.Empty,
                Author = item.GetAttribute("data-author") ?? string.Empty,
                Initials = item.GetAttribute("data-initials") ?? string.Empty,
                Date = TryParseCommentDate(item.GetAttribute("data-date"))
            };

            var repliesList = item.Children.FirstOrDefault(element =>
                string.Equals(element.TagName, "OL", StringComparison.OrdinalIgnoreCase) &&
                string.Equals(element.GetAttribute("class"), "comment-replies", StringComparison.OrdinalIgnoreCase));
            if (repliesList != null) {
                foreach (var reply in repliesList.Children.Where(element => string.Equals(element.TagName, "LI", StringComparison.OrdinalIgnoreCase))) {
                    cancellationToken.ThrowIfCancellationRequested();
                    comment.Replies.Add(CaptureCommentListItem(reply, cancellationToken));
                }
            }

            return comment;
        }

        private static DateTime? TryParseCommentDate(string? value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return null;
            }

            if (DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var date)) {
                return date;
            }

            return null;
        }

        private bool TryProcessCommentAnchor(
            string anchor,
            WordSection section,
            ref WordParagraph? currentParagraph,
            WordTableCell? cell,
            WordHeaderFooter? headerFooter) {
            if (!_commentMap.TryGetValue(anchor, out var comment)) {
                return false;
            }

            currentParagraph ??= cell != null ? cell.AddParagraph("", true) : headerFooter != null ? headerFooter.AddParagraph("") : section.AddParagraph("");
            if (!currentParagraph.GetRuns().Any()) {
                currentParagraph.AddText(string.Empty);
            }

            currentParagraph.AddComment(comment.Author, comment.Initials, comment.Text);
            var created = currentParagraph._document.Comments.LastOrDefault();
            if (created != null) {
                if (comment.Date.HasValue) {
                    created.DateTime = comment.Date.Value;
                }

                foreach (var reply in comment.Replies) {
                    var createdReply = created.AddReply(reply.Author, reply.Initials, reply.Text);
                    if (reply.Date.HasValue) {
                        createdReply.DateTime = reply.Date.Value;
                    }
                }
            }

            return true;
        }
    }
}
