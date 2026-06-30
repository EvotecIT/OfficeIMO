namespace OfficeIMO.Word {
    /// <summary>
    /// Summarizes comments, tracked revisions, unsupported review metadata, and optional accept/reject actions.
    /// </summary>
    public sealed class WordReviewReport {
        private WordReviewReport(WordReviewInfo review, IReadOnlyList<WordRevisionOperationReport> actions) {
            Comments = review.Comments.ToArray();
            CommentThreads = BuildCommentThreads(Comments);
            Revisions = review.Revisions.ToArray();
            UnsupportedMetadata = review.UnsupportedMetadata.ToArray();
            Actions = actions.ToArray();
        }

        /// <summary>Gets all parsed comments in deterministic document order.</summary>
        public IReadOnlyList<WordCommentInfo> Comments { get; }

        /// <summary>Gets top-level comments grouped with their replies.</summary>
        public IReadOnlyList<WordCommentThreadInfo> CommentThreads { get; }

        /// <summary>Gets parsed tracked revisions in deterministic document order.</summary>
        public IReadOnlyList<WordRevisionInfo> Revisions { get; }

        /// <summary>Gets detected modern review metadata that OfficeIMO preserves but does not fully parse yet.</summary>
        public IReadOnlyList<string> UnsupportedMetadata { get; }

        /// <summary>Gets accept/reject operation reports supplied by the caller.</summary>
        public IReadOnlyList<WordRevisionOperationReport> Actions { get; }

        /// <summary>Gets comments that are known unresolved or do not expose resolved metadata.</summary>
        public IReadOnlyList<WordCommentInfo> UnresolvedThreads =>
            Comments.Where(comment => !comment.IsReply && comment.IsResolved != true).ToArray();

        /// <summary>Gets grouped comment threads that are known unresolved or do not expose resolved metadata.</summary>
        public IReadOnlyList<WordCommentThreadInfo> UnresolvedCommentThreads =>
            CommentThreads.Where(thread => thread.IsResolved != true).ToArray();

        /// <summary>Gets the number of parsed comments.</summary>
        public int CommentCount => Comments.Count;

        /// <summary>Gets the number of top-level comment threads.</summary>
        public int CommentThreadCount => CommentThreads.Count;

        /// <summary>Gets the number of parsed tracked revisions.</summary>
        public int RevisionCount => Revisions.Count;

        /// <summary>Gets the number of unresolved top-level comment threads.</summary>
        public int UnresolvedThreadCount => UnresolvedThreads.Count;

        /// <summary>Gets the number of supplied accept/reject action reports.</summary>
        public int ActionCount => Actions.Count;

        /// <summary>Gets the number of detected unsupported review metadata entries.</summary>
        public int UnsupportedMetadataCount => UnsupportedMetadata.Count;

        /// <summary>
        /// Creates a review report from an existing review read model.
        /// </summary>
        /// <param name="review">Review metadata returned by <see cref="WordDocument.InspectReview"/>.</param>
        /// <param name="actions">Optional accept/reject operation reports to include.</param>
        public static WordReviewReport From(WordReviewInfo review, params WordRevisionOperationReport[] actions) {
            if (review == null) {
                throw new ArgumentNullException(nameof(review));
            }

            return new WordReviewReport(review, actions ?? Array.Empty<WordRevisionOperationReport>());
        }

        /// <summary>
        /// Serializes this report to deterministic JSON.
        /// </summary>
        public string ToJson() {
            var builder = new StringBuilder();
            builder.AppendLine("{");
            AppendJsonProperty(builder, 1, "commentCount", CommentCount, comma: true);
            AppendJsonProperty(builder, 1, "commentThreadCount", CommentThreadCount, comma: true);
            AppendJsonProperty(builder, 1, "revisionCount", RevisionCount, comma: true);
            AppendJsonProperty(builder, 1, "unresolvedThreadCount", UnresolvedThreadCount, comma: true);
            AppendJsonProperty(builder, 1, "actionCount", ActionCount, comma: true);
            AppendJsonProperty(builder, 1, "unsupportedMetadataCount", UnsupportedMetadataCount, comma: true);

            AppendJsonComments(builder, 1, "comments", Comments, comma: true);
            AppendJsonComments(builder, 1, "unresolvedThreads", UnresolvedThreads, comma: true);
            AppendJsonCommentThreads(builder, 1, "commentThreads", CommentThreads, comma: true);
            AppendJsonRevisions(builder, 1, "revisions", Revisions, comma: true);
            AppendJsonActions(builder, 1, "actions", Actions, comma: true);
            AppendJsonStringArray(builder, 1, "unsupportedMetadata", UnsupportedMetadata, comma: false);
            builder.Append('}');
            return builder.ToString();
        }

        /// <summary>
        /// Renders this report as Markdown suitable for review notes and automation logs.
        /// </summary>
        public string ToMarkdown() {
            var builder = new StringBuilder();
            builder.AppendLine("# Word Review Report");
            builder.AppendLine();
            builder.AppendLine("## Summary");
            builder.AppendLine();
            builder.AppendLine("| Metric | Count |");
            builder.AppendLine("| --- | ---: |");
            builder.AppendLine($"| Comments | {CommentCount} |");
            builder.AppendLine($"| Comment threads | {CommentThreadCount} |");
            builder.AppendLine($"| Revisions | {RevisionCount} |");
            builder.AppendLine($"| Unresolved threads | {UnresolvedThreadCount} |");
            builder.AppendLine($"| Actions | {ActionCount} |");
            builder.AppendLine($"| Unsupported metadata | {UnsupportedMetadataCount} |");
            builder.AppendLine();

            AppendCommentMarkdown(builder, "Comments", Comments);
            AppendCommentMarkdown(builder, "Unresolved Threads", UnresolvedThreads);
            AppendCommentThreadMarkdown(builder);
            AppendRevisionMarkdown(builder, "Revisions", Revisions);
            AppendActionMarkdown(builder);
            AppendUnsupportedMetadataMarkdown(builder);

            return builder.ToString().TrimEnd();
        }

        private static IReadOnlyList<WordCommentThreadInfo> BuildCommentThreads(IReadOnlyList<WordCommentInfo> comments) {
            var repliesByParentParaId = comments
                .Where(comment => comment.IsReply && !string.IsNullOrWhiteSpace(comment.ParentParaId))
                .GroupBy(comment => comment.ParentParaId!, StringComparer.Ordinal)
                .ToDictionary(group => group.Key, group => (IReadOnlyList<WordCommentInfo>)group.ToArray(), StringComparer.Ordinal);

            var threads = new List<WordCommentThreadInfo>();
            var groupedReplies = new HashSet<WordCommentInfo>();
            foreach (WordCommentInfo parent in comments.Where(comment => !comment.IsReply)) {
                IReadOnlyList<WordCommentInfo> replies = Array.Empty<WordCommentInfo>();
                if (!string.IsNullOrWhiteSpace(parent.ParaId) && repliesByParentParaId.TryGetValue(parent.ParaId!, out IReadOnlyList<WordCommentInfo>? foundReplies)) {
                    replies = foundReplies;
                    foreach (WordCommentInfo reply in foundReplies) {
                        groupedReplies.Add(reply);
                    }
                }

                threads.Add(new WordCommentThreadInfo(parent, replies));
            }

            foreach (WordCommentInfo reply in comments.Where(comment => comment.IsReply && !groupedReplies.Contains(comment))) {
                threads.Add(new WordCommentThreadInfo(reply, Array.Empty<WordCommentInfo>()));
            }

            return threads.ToArray();
        }

        private static void AppendCommentMarkdown(StringBuilder builder, string title, IReadOnlyList<WordCommentInfo> comments) {
            builder.AppendLine($"## {title}");
            builder.AppendLine();
            if (comments.Count == 0) {
                builder.AppendLine("_None._");
                builder.AppendLine();
                return;
            }

            builder.AppendLine("| # | Id | Author | Resolved | Reply | Text | Target | Location |");
            builder.AppendLine("| ---: | --- | --- | --- | --- | --- | --- | --- |");
            foreach (WordCommentInfo comment in comments) {
                builder.Append("| ");
                builder.Append(comment.Index);
                builder.Append(" | ");
                builder.Append(EscapeMarkdownCell(comment.Id));
                builder.Append(" | ");
                builder.Append(EscapeMarkdownCell(comment.Author));
                builder.Append(" | ");
                builder.Append(EscapeMarkdownCell(FormatNullableBool(comment.IsResolved)));
                builder.Append(" | ");
                builder.Append(comment.IsReply ? "yes" : "no");
                builder.Append(" | ");
                builder.Append(EscapeMarkdownCell(comment.Text));
                builder.Append(" | ");
                builder.Append(EscapeMarkdownCell(comment.TargetText));
                builder.Append(" | ");
                builder.Append(EscapeMarkdownCell(FormatCommentLocation(comment)));
                builder.AppendLine(" |");
            }

            builder.AppendLine();
        }

        private static void AppendRevisionMarkdown(StringBuilder builder, string title, IReadOnlyList<WordRevisionInfo> revisions) {
            builder.AppendLine($"## {title}");
            builder.AppendLine();
            if (revisions.Count == 0) {
                builder.AppendLine("_None._");
                builder.AppendLine();
                return;
            }

            builder.AppendLine("| # | Id | Type | Author | Date | Text | Location text | Location |");
            builder.AppendLine("| ---: | --- | --- | --- | --- | --- | --- | --- |");
            foreach (WordRevisionInfo revision in revisions) {
                AppendRevisionMarkdownRow(builder, revision);
            }

            builder.AppendLine();
        }

        private void AppendCommentThreadMarkdown(StringBuilder builder) {
            builder.AppendLine("## Comment Threads");
            builder.AppendLine();
            if (CommentThreads.Count == 0) {
                builder.AppendLine("_None._");
                builder.AppendLine();
                return;
            }

            builder.AppendLine("| Thread | Parent | Replies | Resolved | Target | Location |");
            builder.AppendLine("| ---: | --- | ---: | --- | --- | --- |");
            foreach (WordCommentThreadInfo thread in CommentThreads) {
                builder.Append("| ");
                builder.Append(thread.Parent.Index);
                builder.Append(" | ");
                builder.Append(EscapeMarkdownCell(thread.Parent.Text));
                builder.Append(" | ");
                builder.Append(thread.ReplyCount);
                builder.Append(" | ");
                builder.Append(EscapeMarkdownCell(FormatNullableBool(thread.IsResolved)));
                builder.Append(" | ");
                builder.Append(EscapeMarkdownCell(thread.TargetText));
                builder.Append(" | ");
                builder.Append(EscapeMarkdownCell(FormatCommentLocation(thread.Parent)));
                builder.AppendLine(" |");
            }

            builder.AppendLine();
        }

        private void AppendActionMarkdown(StringBuilder builder) {
            builder.AppendLine("## Actions");
            builder.AppendLine();
            if (Actions.Count == 0) {
                builder.AppendLine("_None._");
                builder.AppendLine();
                return;
            }

            builder.AppendLine("| Action | Matched | Revision | Type | Author | Text | Location text | Location |");
            builder.AppendLine("| --- | ---: | ---: | --- | --- | --- | --- | --- |");
            foreach (WordRevisionOperationReport action in Actions) {
                if (action.MatchedRevisions.Count == 0) {
                    builder.Append("| ");
                    builder.Append(action.Operation);
                    builder.Append(" | 0 |  |  |  |  |  |  |");
                    builder.AppendLine();
                    continue;
                }

                foreach (WordRevisionInfo revision in action.MatchedRevisions) {
                    builder.Append("| ");
                    builder.Append(action.Operation);
                    builder.Append(" | ");
                    builder.Append(action.MatchedCount);
                    builder.Append(" | ");
                    builder.Append(revision.Index);
                    builder.Append(" | ");
                    builder.Append(revision.RevisionType);
                    builder.Append(" | ");
                    builder.Append(EscapeMarkdownCell(revision.Author));
                    builder.Append(" | ");
                    builder.Append(EscapeMarkdownCell(revision.AffectedText));
                    builder.Append(" | ");
                    builder.Append(EscapeMarkdownCell(revision.LocationText));
                    builder.Append(" | ");
                    builder.Append(EscapeMarkdownCell(FormatRevisionLocation(revision)));
                    builder.AppendLine(" |");
                }
            }

            builder.AppendLine();
        }

        private void AppendUnsupportedMetadataMarkdown(StringBuilder builder) {
            builder.AppendLine("## Unsupported Review Metadata");
            builder.AppendLine();
            if (UnsupportedMetadata.Count == 0) {
                builder.AppendLine("_None._");
                return;
            }

            foreach (string detail in UnsupportedMetadata) {
                builder.Append("- ");
                builder.AppendLine(EscapeMarkdownText(detail));
            }
        }

        private static void AppendRevisionMarkdownRow(StringBuilder builder, WordRevisionInfo revision) {
            builder.Append("| ");
            builder.Append(revision.Index);
            builder.Append(" | ");
            builder.Append(EscapeMarkdownCell(revision.Id));
            builder.Append(" | ");
            builder.Append(revision.RevisionType);
            builder.Append(" | ");
            builder.Append(EscapeMarkdownCell(revision.Author));
            builder.Append(" | ");
            builder.Append(EscapeMarkdownCell(FormatDate(revision.DateTime)));
            builder.Append(" | ");
            builder.Append(EscapeMarkdownCell(revision.AffectedText));
            builder.Append(" | ");
            builder.Append(EscapeMarkdownCell(revision.LocationText));
            builder.Append(" | ");
            builder.Append(EscapeMarkdownCell(FormatRevisionLocation(revision)));
            builder.AppendLine(" |");
        }

        private static void AppendJsonComments(StringBuilder builder, int depth, string name, IReadOnlyList<WordCommentInfo> comments, bool comma) {
            AppendIndent(builder, depth);
            AppendJsonString(builder, name);
            builder.AppendLine(": [");
            for (int i = 0; i < comments.Count; i++) {
                WordCommentInfo comment = comments[i];
                AppendIndent(builder, depth + 1);
                builder.AppendLine("{");
                AppendJsonProperty(builder, depth + 2, "index", comment.Index, comma: true);
                AppendJsonProperty(builder, depth + 2, "id", comment.Id, comma: true);
                AppendJsonProperty(builder, depth + 2, "author", comment.Author, comma: true);
                AppendJsonProperty(builder, depth + 2, "initials", comment.Initials, comma: true);
                AppendJsonProperty(builder, depth + 2, "dateTime", comment.DateTime, comma: true);
                AppendJsonProperty(builder, depth + 2, "text", comment.Text, comma: true);
                AppendJsonProperty(builder, depth + 2, "paraId", comment.ParaId, comma: true);
                AppendJsonProperty(builder, depth + 2, "parentParaId", comment.ParentParaId, comma: true);
                AppendJsonProperty(builder, depth + 2, "isReply", comment.IsReply, comma: true);
                AppendJsonProperty(builder, depth + 2, "isResolved", comment.IsResolved, comma: true);
                AppendJsonProperty(builder, depth + 2, "targetText", comment.TargetText, comma: true);
                AppendJsonProperty(builder, depth + 2, "targetLocationKind", comment.TargetLocationKind?.ToString(), comma: true);
                AppendJsonProperty(builder, depth + 2, "targetPartUri", comment.TargetPartUri, comma: true);
                AppendJsonProperty(builder, depth + 2, "isInTable", comment.IsInTable, comma: true);
                AppendJsonProperty(builder, depth + 2, "isInContentControl", comment.IsInContentControl, comma: true);
                AppendJsonProperty(builder, depth + 2, "isInTextBox", comment.IsInTextBox, comma: false);
                AppendIndent(builder, depth + 1);
                builder.Append('}');
                builder.AppendLine(i == comments.Count - 1 ? string.Empty : ",");
            }

            AppendIndent(builder, depth);
            builder.Append(']');
            builder.AppendLine(comma ? "," : string.Empty);
        }

        private static void AppendJsonCommentThreads(StringBuilder builder, int depth, string name, IReadOnlyList<WordCommentThreadInfo> threads, bool comma) {
            AppendIndent(builder, depth);
            AppendJsonString(builder, name);
            builder.AppendLine(": [");
            for (int i = 0; i < threads.Count; i++) {
                WordCommentThreadInfo thread = threads[i];
                AppendIndent(builder, depth + 1);
                builder.AppendLine("{");
                AppendJsonProperty(builder, depth + 2, "replyCount", thread.ReplyCount, comma: true);
                AppendJsonProperty(builder, depth + 2, "isResolved", thread.IsResolved, comma: true);
                AppendJsonProperty(builder, depth + 2, "targetText", thread.TargetText, comma: true);
                AppendIndent(builder, depth + 2);
                AppendJsonString(builder, "parent");
                builder.AppendLine(": {");
                AppendJsonCommentProperties(builder, depth + 3, thread.Parent);
                AppendIndent(builder, depth + 2);
                builder.AppendLine("},");
                AppendJsonComments(builder, depth + 2, "replies", thread.Replies, comma: false);
                AppendIndent(builder, depth + 1);
                builder.Append('}');
                builder.AppendLine(i == threads.Count - 1 ? string.Empty : ",");
            }

            AppendIndent(builder, depth);
            builder.Append(']');
            builder.AppendLine(comma ? "," : string.Empty);
        }

        private static void AppendJsonCommentProperties(StringBuilder builder, int depth, WordCommentInfo comment) {
            AppendJsonProperty(builder, depth, "index", comment.Index, comma: true);
            AppendJsonProperty(builder, depth, "id", comment.Id, comma: true);
            AppendJsonProperty(builder, depth, "author", comment.Author, comma: true);
            AppendJsonProperty(builder, depth, "initials", comment.Initials, comma: true);
            AppendJsonProperty(builder, depth, "dateTime", comment.DateTime, comma: true);
            AppendJsonProperty(builder, depth, "text", comment.Text, comma: true);
            AppendJsonProperty(builder, depth, "paraId", comment.ParaId, comma: true);
            AppendJsonProperty(builder, depth, "parentParaId", comment.ParentParaId, comma: true);
            AppendJsonProperty(builder, depth, "isReply", comment.IsReply, comma: true);
            AppendJsonProperty(builder, depth, "isResolved", comment.IsResolved, comma: true);
            AppendJsonProperty(builder, depth, "targetText", comment.TargetText, comma: true);
            AppendJsonProperty(builder, depth, "targetLocationKind", comment.TargetLocationKind?.ToString(), comma: true);
            AppendJsonProperty(builder, depth, "targetPartUri", comment.TargetPartUri, comma: true);
            AppendJsonProperty(builder, depth, "isInTable", comment.IsInTable, comma: true);
            AppendJsonProperty(builder, depth, "isInContentControl", comment.IsInContentControl, comma: true);
            AppendJsonProperty(builder, depth, "isInTextBox", comment.IsInTextBox, comma: false);
        }

        private static void AppendJsonRevisions(StringBuilder builder, int depth, string name, IReadOnlyList<WordRevisionInfo> revisions, bool comma) {
            AppendIndent(builder, depth);
            AppendJsonString(builder, name);
            builder.AppendLine(": [");
            for (int i = 0; i < revisions.Count; i++) {
                AppendJsonRevision(builder, depth + 1, revisions[i]);
                builder.AppendLine(i == revisions.Count - 1 ? string.Empty : ",");
            }

            AppendIndent(builder, depth);
            builder.Append(']');
            builder.AppendLine(comma ? "," : string.Empty);
        }

        private static void AppendJsonActions(StringBuilder builder, int depth, string name, IReadOnlyList<WordRevisionOperationReport> actions, bool comma) {
            AppendIndent(builder, depth);
            AppendJsonString(builder, name);
            builder.AppendLine(": [");
            for (int i = 0; i < actions.Count; i++) {
                WordRevisionOperationReport action = actions[i];
                AppendIndent(builder, depth + 1);
                builder.AppendLine("{");
                AppendJsonProperty(builder, depth + 2, "operation", action.Operation.ToString(), comma: true);
                AppendJsonProperty(builder, depth + 2, "matchedCount", action.MatchedCount, comma: true);
                AppendJsonRevisions(builder, depth + 2, "matchedRevisions", action.MatchedRevisions, comma: false);
                AppendIndent(builder, depth + 1);
                builder.Append('}');
                builder.AppendLine(i == actions.Count - 1 ? string.Empty : ",");
            }

            AppendIndent(builder, depth);
            builder.Append(']');
            builder.AppendLine(comma ? "," : string.Empty);
        }

        private static void AppendJsonRevision(StringBuilder builder, int depth, WordRevisionInfo revision) {
            AppendIndent(builder, depth);
            builder.AppendLine("{");
            AppendJsonProperty(builder, depth + 1, "index", revision.Index, comma: true);
            AppendJsonProperty(builder, depth + 1, "revisionType", revision.RevisionType.ToString(), comma: true);
            AppendJsonProperty(builder, depth + 1, "elementName", revision.ElementName, comma: true);
            AppendJsonProperty(builder, depth + 1, "id", revision.Id, comma: true);
            AppendJsonProperty(builder, depth + 1, "author", revision.Author, comma: true);
            AppendJsonProperty(builder, depth + 1, "dateTime", revision.DateTime, comma: true);
            AppendJsonProperty(builder, depth + 1, "affectedText", revision.AffectedText, comma: true);
            AppendJsonProperty(builder, depth + 1, "locationText", revision.LocationText, comma: true);
            AppendJsonProperty(builder, depth + 1, "locationKind", revision.LocationKind.ToString(), comma: true);
            AppendJsonProperty(builder, depth + 1, "partUri", revision.PartUri, comma: true);
            AppendJsonProperty(builder, depth + 1, "isInTable", revision.IsInTable, comma: true);
            AppendJsonProperty(builder, depth + 1, "isInContentControl", revision.IsInContentControl, comma: true);
            AppendJsonProperty(builder, depth + 1, "isInTextBox", revision.IsInTextBox, comma: false);
            AppendIndent(builder, depth);
            builder.Append('}');
        }

        private static void AppendJsonStringArray(StringBuilder builder, int depth, string name, IReadOnlyList<string> values, bool comma) {
            AppendIndent(builder, depth);
            AppendJsonString(builder, name);
            builder.AppendLine(": [");
            for (int i = 0; i < values.Count; i++) {
                AppendIndent(builder, depth + 1);
                AppendJsonString(builder, values[i]);
                builder.AppendLine(i == values.Count - 1 ? string.Empty : ",");
            }

            AppendIndent(builder, depth);
            builder.Append(']');
            builder.AppendLine(comma ? "," : string.Empty);
        }

        private static void AppendJsonProperty(StringBuilder builder, int depth, string name, string? value, bool comma) {
            AppendIndent(builder, depth);
            AppendJsonString(builder, name);
            builder.Append(": ");
            if (value == null) {
                builder.Append("null");
            } else {
                AppendJsonString(builder, value);
            }

            builder.AppendLine(comma ? "," : string.Empty);
        }

        private static void AppendJsonProperty(StringBuilder builder, int depth, string name, int value, bool comma) {
            AppendIndent(builder, depth);
            AppendJsonString(builder, name);
            builder.Append(": ");
            builder.Append(value.ToString(System.Globalization.CultureInfo.InvariantCulture));
            builder.AppendLine(comma ? "," : string.Empty);
        }

        private static void AppendJsonProperty(StringBuilder builder, int depth, string name, bool value, bool comma) {
            AppendIndent(builder, depth);
            AppendJsonString(builder, name);
            builder.Append(": ");
            builder.Append(value ? "true" : "false");
            builder.AppendLine(comma ? "," : string.Empty);
        }

        private static void AppendJsonProperty(StringBuilder builder, int depth, string name, bool? value, bool comma) {
            AppendIndent(builder, depth);
            AppendJsonString(builder, name);
            builder.Append(": ");
            builder.Append(value.HasValue ? (value.Value ? "true" : "false") : "null");
            builder.AppendLine(comma ? "," : string.Empty);
        }

        private static void AppendJsonProperty(StringBuilder builder, int depth, string name, DateTime? value, bool comma) {
            AppendIndent(builder, depth);
            AppendJsonString(builder, name);
            builder.Append(": ");
            if (value.HasValue) {
                AppendJsonString(builder, FormatDate(value));
            } else {
                builder.Append("null");
            }

            builder.AppendLine(comma ? "," : string.Empty);
        }

        private static void AppendJsonString(StringBuilder builder, string value) {
            builder.Append('"');
            foreach (char ch in value) {
                switch (ch) {
                    case '"':
                        builder.Append("\\\"");
                        break;
                    case '\\':
                        builder.Append("\\\\");
                        break;
                    case '\b':
                        builder.Append("\\b");
                        break;
                    case '\f':
                        builder.Append("\\f");
                        break;
                    case '\n':
                        builder.Append("\\n");
                        break;
                    case '\r':
                        builder.Append("\\r");
                        break;
                    case '\t':
                        builder.Append("\\t");
                        break;
                    default:
                        if (char.IsControl(ch)) {
                            builder.Append("\\u");
                            builder.Append(((int)ch).ToString("x4", System.Globalization.CultureInfo.InvariantCulture));
                        } else {
                            builder.Append(ch);
                        }

                        break;
                }
            }

            builder.Append('"');
        }

        private static void AppendIndent(StringBuilder builder, int depth) {
            builder.Append(' ', depth * 2);
        }

        private static string FormatDate(DateTime? dateTime) =>
            dateTime?.ToUniversalTime().ToString("O", System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty;

        private static string FormatNullableBool(bool? value) =>
            value.HasValue ? (value.Value ? "yes" : "no") : "unknown";

        private static string FormatCommentLocation(WordCommentInfo comment) {
            string kind = comment.TargetLocationKind?.ToString() ?? "Unknown";
            return FormatLocation(kind, comment.TargetPartUri, comment.IsInTable, comment.IsInContentControl, comment.IsInTextBox);
        }

        private static string FormatRevisionLocation(WordRevisionInfo revision) =>
            FormatLocation(revision.LocationKind.ToString(), revision.PartUri, revision.IsInTable, revision.IsInContentControl, revision.IsInTextBox);

        private static string FormatLocation(string kind, string? partUri, bool isInTable, bool isInContentControl, bool isInTextBox) {
            var details = new List<string> { kind };
            if (!string.IsNullOrWhiteSpace(partUri)) {
                details.Add(partUri!);
            }

            if (isInTable) {
                details.Add("table");
            }

            if (isInContentControl) {
                details.Add("content-control");
            }

            if (isInTextBox) {
                details.Add("text-box");
            }

            return string.Join(" / ", details);
        }

        private static string EscapeMarkdownCell(string? value) =>
            EscapeMarkdownText(value ?? string.Empty)
                .Replace("|", "\\|");

        private static string EscapeMarkdownText(string value) =>
            value.Replace("\r", " ").Replace("\n", " ");
    }

    /// <summary>
    /// Groups a top-level Word comment with its parsed replies for review-report consumers.
    /// </summary>
    public sealed class WordCommentThreadInfo {
        internal WordCommentThreadInfo(WordCommentInfo parent, IReadOnlyList<WordCommentInfo> replies) {
            Parent = parent ?? throw new ArgumentNullException(nameof(parent));
            Replies = replies.ToArray();
        }

        /// <summary>Gets the top-level comment for this thread.</summary>
        public WordCommentInfo Parent { get; }

        /// <summary>Gets replies associated with the parent comment.</summary>
        public IReadOnlyList<WordCommentInfo> Replies { get; }

        /// <summary>Gets the number of replies in this thread.</summary>
        public int ReplyCount => Replies.Count;

        /// <summary>Gets whether the parent comment is marked resolved, or null when the document does not expose that metadata.</summary>
        public bool? IsResolved => Parent.IsResolved;

        /// <summary>Gets the text targeted by the parent comment.</summary>
        public string TargetText => Parent.TargetText;
    }
}
