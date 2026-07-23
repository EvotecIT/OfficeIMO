namespace OfficeIMO.Word {
    public static partial class WordDocumentComparer {
        private static void AnalyzeComments(WordDocument source, WordDocument target, WordComparisonResult result, WordComparisonOptions options) {
            AddFeatureRangeFindings(
                GetCommentSnapshots(source, options),
                GetCommentSnapshots(target, options),
                WordComparisonScope.Comment,
                "comment",
                "Comment",
                result);
        }

        private static void AnalyzeRevisions(WordDocument source, WordDocument target, WordComparisonResult result, WordComparisonOptions options) {
            AddFeatureRangeFindings(
                GetRevisionSnapshots(source, options),
                GetRevisionSnapshots(target, options),
                WordComparisonScope.Revision,
                "revision",
                "Revision",
                result);
        }

        private static List<CommentSnapshot> GetCommentSnapshots(WordDocument document, WordComparisonOptions options) {
            return document.InspectReview().Comments
                .Where(comment => options.CompareCommentReplies || !comment.IsReply)
                .Select(comment => new CommentSnapshot(
                    comment.Index,
                    GetCommentMatchKey(comment, options),
                    GetCommentSignature(comment, options),
                    GetCommentDisplayText(comment),
                    GetCommentDetailedLocation(comment),
                    comment.DocumentOrder))
                .ToList();
        }

        private static string GetCommentMatchKey(WordCommentInfo comment, WordComparisonOptions options) {
            return string.Join(
                "|",
                options.CompareCommentTargets ? (comment.TargetLocationKind?.ToString() ?? string.Empty) : string.Empty,
                options.CompareCommentTargets ? (comment.TargetPartUri ?? string.Empty) : string.Empty,
                comment.IsReply ? "reply" : "comment",
                options.CompareCommentAuthors ? NormalizeComparisonText(comment.Author ?? string.Empty, options) : string.Empty,
                options.CompareCommentAuthors ? NormalizeComparisonText(comment.Initials ?? string.Empty, options) : string.Empty,
                options.CompareCommentText ? NormalizeComparisonText(comment.Text, options) : string.Empty,
                options.CompareCommentTargets ? NormalizeComparisonText(comment.TargetText, options) : string.Empty,
                options.CompareCommentResolvedState ? FormatReviewBoolean(comment.IsResolved) : string.Empty);
        }

        private static string GetCommentSignature(WordCommentInfo comment, WordComparisonOptions options) {
            return string.Join(
                "|",
                options.CompareGeneratedIds ? NormalizeComparisonText(comment.Id ?? string.Empty, options) : string.Empty,
                options.CompareCommentAuthors ? NormalizeComparisonText(comment.Author ?? string.Empty, options) : string.Empty,
                options.CompareCommentAuthors ? NormalizeComparisonText(comment.Initials ?? string.Empty, options) : string.Empty,
                options.CompareVolatileMetadata ? FormatReviewDate(comment.DateTime) : string.Empty,
                options.CompareCommentText ? NormalizeComparisonText(comment.Text, options) : string.Empty,
                options.CompareGeneratedIds ? NormalizeComparisonText(comment.ParaId ?? string.Empty, options) : string.Empty,
                options.CompareGeneratedIds ? NormalizeComparisonText(comment.ParentParaId ?? string.Empty, options) : string.Empty,
                options.CompareCommentResolvedState ? FormatReviewBoolean(comment.IsResolved) : string.Empty,
                options.CompareCommentTargets ? NormalizeComparisonText(comment.TargetText, options) : string.Empty,
                options.CompareCommentTargets ? (comment.TargetLocationKind?.ToString() ?? string.Empty) : string.Empty,
                options.CompareCommentTargets ? (comment.TargetPartUri ?? string.Empty) : string.Empty,
                options.CompareCommentTargets && comment.IsInTable ? "table" : string.Empty,
                options.CompareCommentTargets && comment.IsInContentControl ? "content-control" : string.Empty,
                options.CompareCommentTargets && comment.IsInTextBox ? "text-box" : string.Empty);
        }

        private static string GetCommentDisplayText(WordCommentInfo comment) {
            return string.Join(
                "; ",
                (comment.IsReply ? "reply" : "comment") + " id=" + (comment.Id ?? string.Empty),
                "author=" + (comment.Author ?? string.Empty),
                "initials=" + (comment.Initials ?? string.Empty),
                "date=" + FormatReviewDate(comment.DateTime),
                "resolved=" + FormatReviewBoolean(comment.IsResolved),
                "target=" + comment.TargetText,
                "text=" + comment.Text);
        }

        private static string GetCommentDetailedLocation(WordCommentInfo comment) {
            return JoinFeatureLocation(
                comment.TargetLocationKind?.ToString() ?? string.Empty,
                comment.TargetPartUri ?? string.Empty,
                FeatureLocation("comment", comment.Index),
                comment.IsReply ? "reply" : "comment",
                comment.IsInTable ? "table" : string.Empty,
                comment.IsInContentControl ? "content-control" : string.Empty,
                comment.IsInTextBox ? "text-box" : string.Empty);
        }

        private static List<RevisionSnapshot> GetRevisionSnapshots(WordDocument document, WordComparisonOptions options) {
            return document.InspectReview().Revisions
                .Select(revision => new RevisionSnapshot(
                    revision.Index,
                    GetRevisionMatchKey(revision, options),
                    GetRevisionSignature(revision, options),
                    GetRevisionDisplayText(revision),
                    GetRevisionDetailedLocation(revision),
                    revision.DocumentOrder))
                .ToList();
        }

        private static string GetRevisionMatchKey(WordRevisionInfo revision, WordComparisonOptions options) {
            return string.Join(
                "|",
                revision.RevisionType.ToString(),
                revision.ElementName,
                options.CompareRevisionText ? NormalizeComparisonText(revision.AffectedText, options) : string.Empty,
                options.CompareRevisionLocations ? revision.LocationKind.ToString() : string.Empty,
                options.CompareRevisionLocations ? revision.PartUri : string.Empty,
                options.CompareRevisionLocations ? NormalizeComparisonText(revision.LocationText, options) : string.Empty,
                options.CompareRevisionLocations && revision.IsInTable ? "table" : string.Empty,
                options.CompareRevisionLocations && revision.IsInContentControl ? "content-control" : string.Empty,
                options.CompareRevisionLocations && revision.IsInTextBox ? "text-box" : string.Empty);
        }

        private static string GetRevisionSignature(WordRevisionInfo revision, WordComparisonOptions options) {
            return string.Join(
                "|",
                revision.RevisionType.ToString(),
                revision.ElementName,
                options.CompareGeneratedIds ? NormalizeComparisonText(revision.Id ?? string.Empty, options) : string.Empty,
                options.CompareRevisionAuthors ? NormalizeComparisonText(revision.Author ?? string.Empty, options) : string.Empty,
                options.CompareVolatileMetadata ? FormatReviewDate(revision.DateTime) : string.Empty,
                options.CompareRevisionText ? NormalizeComparisonText(revision.AffectedText, options) : string.Empty,
                options.CompareRevisionText ? NormalizeComparisonText(revision.LocationText, options) : string.Empty,
                options.CompareRevisionLocations ? revision.LocationKind.ToString() : string.Empty,
                options.CompareRevisionLocations ? revision.PartUri : string.Empty,
                options.CompareRevisionLocations && revision.IsInTable ? "table" : string.Empty,
                options.CompareRevisionLocations && revision.IsInContentControl ? "content-control" : string.Empty,
                options.CompareRevisionLocations && revision.IsInTextBox ? "text-box" : string.Empty);
        }

        private static string GetRevisionDisplayText(WordRevisionInfo revision) {
            return string.Join(
                "; ",
                revision.RevisionType + " element=" + revision.ElementName,
                "id=" + (revision.Id ?? string.Empty),
                "author=" + (revision.Author ?? string.Empty),
                "date=" + FormatReviewDate(revision.DateTime),
                "text=" + revision.AffectedText,
                "location=" + revision.LocationText);
        }

        private static string GetRevisionDetailedLocation(WordRevisionInfo revision) {
            return JoinFeatureLocation(
                revision.LocationKind.ToString(),
                revision.PartUri,
                FeatureLocation("revision", revision.Index),
                revision.RevisionType.ToString(),
                revision.ElementName,
                revision.IsInTable ? "table" : string.Empty,
                revision.IsInContentControl ? "content-control" : string.Empty,
                revision.IsInTextBox ? "text-box" : string.Empty);
        }

        private static string FormatReviewDate(DateTime? value) {
            return value?.ToString("O", System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty;
        }

        private static string FormatReviewBoolean(bool? value) {
            return value.HasValue ? (value.Value ? "true" : "false") : "unknown";
        }

        private sealed class CommentSnapshot : IFeatureSnapshot {
            internal CommentSnapshot(int index, string matchKey, string signature, string displayText, string detailedLocation, int documentOrder) {
                Index = index;
                MatchKey = matchKey;
                Signature = signature;
                DisplayText = displayText;
                DetailedLocation = detailedLocation;
                DocumentOrder = documentOrder;
            }

            public int Index { get; }

            public string MatchKey { get; }

            public string Signature { get; }

            public string DisplayText { get; }

            public string DetailedLocation { get; }

            public int DocumentOrder { get; }

            ulong IComparisonFingerprint.ComparisonFingerprint => GetOrdinalTextFingerprint(MatchKey);
        }

        private sealed class RevisionSnapshot : IFeatureSnapshot {
            internal RevisionSnapshot(int index, string matchKey, string signature, string displayText, string detailedLocation, int documentOrder) {
                Index = index;
                MatchKey = matchKey;
                Signature = signature;
                DisplayText = displayText;
                DetailedLocation = detailedLocation;
                DocumentOrder = documentOrder;
            }

            public int Index { get; }

            public string MatchKey { get; }

            public string Signature { get; }

            public string DisplayText { get; }

            public string DetailedLocation { get; }

            public int DocumentOrder { get; }

            ulong IComparisonFingerprint.ComparisonFingerprint => GetOrdinalTextFingerprint(MatchKey);
        }
    }
}
