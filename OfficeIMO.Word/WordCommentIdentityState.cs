using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2013.Word;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

/// <summary>
/// Next comment and comment paragraph identifiers for one document. It is built by one pass over the comments
/// and stays valid while both comment roots are the same elements and still end with the last comment added
/// through it; any other append, removal of the last comment or replaced part makes the next comment rescan.
/// </summary>
internal sealed class WordCommentIdentityState {
    private readonly Comments _comments;
    private readonly CommentsEx _commentsEx;
    private OpenXmlElement? _lastComment;
    private OpenXmlElement? _lastCommentEx;
    private int _lastId;
    private uint _lastParagraphId;

    private WordCommentIdentityState(Comments comments, CommentsEx commentsEx) {
        _comments = comments;
        _commentsEx = commentsEx;
    }

    /// <summary>Returns the current state, normalizing legacy comment identities and rescanning when it is stale.</summary>
    internal static WordCommentIdentityState Get(WordDocument document, Comments comments, CommentsEx commentsEx) {
        WordCommentIdentityState? state = document.CommentIdentities;
        if (state != null && !document._openXmlDocumentAccessed && state.IsCurrent(comments, commentsEx)) return state;

        WordCommentIdentityPlan plan = WordCommentIdentityPlan.Prepare(comments, commentsEx);
        plan.Apply(commentsEx);
        state = new WordCommentIdentityState(comments, commentsEx) { _lastParagraphId = plan.LastParagraphIdentifier };
        foreach (Comment comment in comments.Elements<Comment>()) {
            if (int.TryParse(comment.Id?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int id) && id > state._lastId) {
                state._lastId = id;
            }
        }
        state.Record();
        document.CommentIdentities = state;
        return state;
    }

    internal string NextId() {
        if (_lastId == int.MaxValue) throw new InvalidOperationException("The comment identifier range is exhausted.");
        return (++_lastId).ToString(CultureInfo.InvariantCulture);
    }

    internal string NextParagraphId() {
        if (_lastParagraphId == uint.MaxValue) throw new InvalidOperationException("The comment paragraph identity range is exhausted.");
        return (++_lastParagraphId).ToString("X8", CultureInfo.InvariantCulture);
    }

    /// <summary>Takes in a comment just appended through this state, including paragraph ids it carried in.</summary>
    internal void Added(Comment comment) {
        foreach (Paragraph paragraph in comment.Elements<Paragraph>()) {
            if (uint.TryParse(paragraph.ParagraphId?.Value, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out uint value) &&
                value > _lastParagraphId) {
                _lastParagraphId = value;
            }
        }
        Record();
    }

    private bool IsCurrent(Comments comments, CommentsEx commentsEx) =>
        ReferenceEquals(comments, _comments) && ReferenceEquals(commentsEx, _commentsEx) &&
        ReferenceEquals(comments.LastChild, _lastComment) && ReferenceEquals(commentsEx.LastChild, _lastCommentEx);

    private void Record() {
        _lastComment = _comments.LastChild;
        _lastCommentEx = _commentsEx.LastChild;
    }
}
