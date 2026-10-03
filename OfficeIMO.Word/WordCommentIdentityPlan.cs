using System.Globalization;
using System.Threading;
using DocumentFormat.OpenXml.Office2013.Word;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

/// <summary>Prepares stable identities for legacy roots before appending modern comment metadata.</summary>
internal sealed class WordCommentIdentityPlan {
    private readonly List<(Paragraph Paragraph, string Identifier, CommentEx Extension, bool Append)> _updates = new();
    internal uint LastParagraphIdentifier { get; private set; }
    internal bool HasChanges => _updates.Count > 0;

    internal static WordCommentIdentityPlan Prepare(Comments? comments, CommentsEx? commentsEx,
        CancellationToken cancellationToken = default) {
        var plan = new WordCommentIdentityPlan();
        Comment[] roots = comments?.Elements<Comment>().ToArray() ?? Array.Empty<Comment>();
        CommentEx[] extensions = commentsEx?.Elements<CommentEx>().ToArray() ?? Array.Empty<CommentEx>();
        var index = WordComment.IndexCommentExByParagraphId(extensions);
        var owners = new Dictionary<CommentEx, Comment>();
        var assignedIdentifiers = new HashSet<string>(StringComparer.Ordinal);
        foreach (Comment root in roots) {
            cancellationToken.ThrowIfCancellationRequested();
            foreach (Paragraph paragraph in root.Elements<Paragraph>()) plan.Observe(paragraph.ParagraphId?.Value);
            string? identifier = WordComment.GetCommentParagraphId(root);
            if (identifier != null && !assignedIdentifiers.Add(identifier))
                throw new InvalidOperationException("Existing comment paragraph identities are ambiguous.");
            if (identifier != null && index.TryGetValue(identifier, out CommentEx? extension)) {
                if (owners.ContainsKey(extension))
                    throw new InvalidOperationException("Existing comment identities are ambiguous.");
                owners.Add(extension, root);
            }
        }
        foreach (CommentEx extension in extensions) {
            cancellationToken.ThrowIfCancellationRequested();
            plan.Observe(extension.ParaId?.Value);
        }
        for (int position = 0; position < roots.Length; position++) {
            cancellationToken.ThrowIfCancellationRequested();
            Comment root = roots[position];
            if (WordComment.GetCommentParagraphId(root) != null) continue;
            Paragraph paragraph = root.Elements<Paragraph>().FirstOrDefault()
                ?? throw new InvalidOperationException("An existing comment must contain a paragraph.");
            CommentEx? extension = WordComment.FindCommentExForComment(root, extensions, index, position);
            if (extension != null) {
                if (owners.ContainsKey(extension))
                    throw new InvalidOperationException("Existing legacy comment metadata has an ambiguous owner.");
                owners.Add(extension, root);
            }
            string? identifier = extension?.ParaId?.Value;
            if (string.IsNullOrWhiteSpace(identifier)) {
                if (plan.LastParagraphIdentifier == uint.MaxValue)
                    throw new InvalidOperationException("The comment paragraph identity range is exhausted.");
                identifier = (++plan.LastParagraphIdentifier).ToString("X8", CultureInfo.InvariantCulture);
            }
            if (!assignedIdentifiers.Add(identifier!))
                throw new InvalidOperationException("Existing legacy comment identities are ambiguous.");
            plan._updates.Add((paragraph, identifier!, extension ?? new CommentEx(), extension == null));
        }
        return plan;
    }

    internal void Apply(CommentsEx commentsEx) {
        foreach (var update in _updates) {
            update.Paragraph.ParagraphId = update.Identifier;
            update.Extension.ParaId = update.Identifier;
            if (update.Append) commentsEx.AppendChild(update.Extension);
        }
    }

    private void Observe(string? identifier) {
        if (uint.TryParse(identifier, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out uint value))
            LastParagraphIdentifier = Math.Max(LastParagraphIdentifier, value);
    }
}
