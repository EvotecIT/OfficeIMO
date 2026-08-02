using AngleSharp.Dom;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;
using System.Threading;

namespace OfficeIMO.Word.Html {
    internal partial class WordToHtmlConverter {
        private static bool TryAppendCommentReference(
            IDocument htmlDoc,
            WordParagraph run,
            WordToHtmlOptions options,
            Dictionary<string, WordComment> commentsById,
            List<(int Number, WordComment Comment)> comments,
            Dictionary<string, int> commentMap,
            List<INode> nodes) {
            if (!options.ExportComments || run._run == null) {
                return false;
            }

            var commentReference = run._run.Elements<CommentReference>().FirstOrDefault();
            var commentId = commentReference?.Id?.Value;
            if (string.IsNullOrEmpty(commentId)) {
                return false;
            }

            if (!commentsById.TryGetValue(commentId!, out var comment)) {
                return true;
            }

            if (!commentMap.TryGetValue(commentId!, out var number)) {
                number = comments.Count + 1;
                commentMap[commentId!] = number;
                comments.Add((number, comment));
            }

            var sup = CreateOutputElement(htmlDoc, "sup");
            var anchor = CreateOutputElement(htmlDoc, "a");
            string numberText = number.ToString(CultureInfo.InvariantCulture);
            SetOutputAttribute(htmlDoc, anchor, "href", $"#comment{numberText}", "CommentReference:href");
            SetOutputAttribute(htmlDoc, anchor, "id", $"commentref{numberText}", "CommentReference:id");
            SetOutputAttribute(htmlDoc, anchor, "data-word-comment-id", commentId!, "CommentReference:comment-id");
            if (!string.IsNullOrEmpty(comment.Author)) {
                SetOutputAttribute(htmlDoc, anchor, "title", $"Comment by {comment.Author}", "CommentReference:title");
            }
            SetOutputText(htmlDoc, anchor, $"[{numberText}]", "CommentReference:text");
            sup.AppendChild(anchor);
            nodes.Add(sup);
            return true;
        }

        private static void AppendComments(
            IDocument htmlDoc,
            IElement body,
            List<(int Number, WordComment Comment)> comments,
            WordToHtmlOptions options,
            CancellationToken cancellationToken) {
            if (!options.ExportComments || comments.Count == 0) {
                return;
            }

            var commentSection = CreateOutputElement(htmlDoc, "section");
            commentSection.SetAttribute("class", "comments");
            var hr = CreateOutputElement(htmlDoc, "hr");
            commentSection.AppendChild(hr);
            var ol = CreateOutputElement(htmlDoc, "ol");
            foreach (var (number, comment) in comments) {
                cancellationToken.ThrowIfCancellationRequested();
                AppendCommentListItem(htmlDoc, ol, $"comment{number.ToString(CultureInfo.InvariantCulture)}", comment, cancellationToken);
            }
            commentSection.AppendChild(ol);
            body.AppendChild(commentSection);
        }

        private static void AppendCommentListItem(IDocument htmlDoc, IElement parent, string id, WordComment comment, CancellationToken cancellationToken) {
            var li = CreateOutputElement(htmlDoc, "li");
            li.SetAttribute("id", id);
            SetIfNotEmpty(li, "data-word-comment-id", comment.Id);
            SetIfNotEmpty(li, "data-author", comment.Author);
            SetIfNotEmpty(li, "data-initials", comment.Initials);
            if (comment.DateTime.HasValue) {
                li.SetAttribute("data-date", comment.DateTime.Value.ToString("o", CultureInfo.InvariantCulture));
            }

            var paragraph = CreateOutputElement(htmlDoc, "p");
            paragraph.TextContent = comment.Text ?? string.Empty;
            li.AppendChild(paragraph);

            var replies = comment.Replies;
            if (replies.Count > 0) {
                var replyList = CreateOutputElement(htmlDoc, "ol");
                replyList.SetAttribute("class", "comment-replies");
                for (int i = 0; i < replies.Count; i++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    AppendCommentListItem(htmlDoc, replyList, $"{id}-reply{(i + 1).ToString(CultureInfo.InvariantCulture)}", replies[i], cancellationToken);
                }
                li.AppendChild(replyList);
            }

            parent.AppendChild(li);
        }

        private static void SetIfNotEmpty(IElement element, string name, string? value) {
            if (!string.IsNullOrEmpty(value)) {
                element.SetAttribute(name, value!);
            }
        }
    }
}
