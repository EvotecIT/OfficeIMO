using DocumentFormat.OpenXml.Office2013.Word;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

public partial class WordComment {
    internal static WordComment CreateDetached(WordDocument document, WordCellComment definition,
        string identifier, string paragraphIdentifier) {
        var paragraph = new Paragraph(new Run(new Text(string.Empty))) { ParagraphId = paragraphIdentifier };
        var comment = new Comment(paragraph) {
            Id = identifier, Author = definition.Author, Initials = definition.Initials
        };
        if (definition.DateTime.HasValue) comment.Date = definition.DateTime.Value;
        var result = new WordComment(document, comment, new CommentEx { ParaId = paragraphIdentifier });
        result.Text = definition.Text;
        return result;
    }

    internal void AppendTo(Comments comments, CommentsEx commentsEx) {
        comments.AppendChild(_comment);
        commentsEx.AppendChild(_commentEx!);
    }
}
