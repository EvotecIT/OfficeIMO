using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2013.Word;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Rtf;

public static partial class WordRtfConverterExtensions {
    private static void AppendWordComments(WordParagraph wordParagraph, RtfParagraph paragraph, RtfDocument rtfDocument, Dictionary<string, int> revisionAuthorIndexes) {
        List<string> commentIds = wordParagraph._paragraph
            .Descendants<CommentRangeStart>()
            .Select(start => start.Id?.Value)
            .Concat(wordParagraph._paragraph.Descendants<CommentReference>().Select(reference => reference.Id?.Value))
            .Where(id => !string.IsNullOrWhiteSpace(id))
            .Distinct(StringComparer.Ordinal)
            .Cast<string>()
            .ToList();

        if (commentIds.Count == 0) {
            return;
        }

        Dictionary<string, WordComment> commentsById = wordParagraph._document.Comments
            .Where(comment => !string.IsNullOrWhiteSpace(comment.Id))
            .GroupBy(comment => comment.Id!, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);

        foreach (string commentId in commentIds) {
            if (!commentsById.TryGetValue(commentId, out WordComment? comment)) {
                continue;
            }

            AttachAnnotation(paragraph, CreateAnnotation(comment, rtfDocument, revisionAuthorIndexes));
        }
    }

    private static RtfNote CreateAnnotation(WordComment comment, RtfDocument rtfDocument, Dictionary<string, int> revisionAuthorIndexes) {
        var note = new RtfNote(RtfNoteKind.Annotation) {
            Id = comment.Id,
            Author = comment.Author,
            Created = comment.DateTime
        };

        foreach (WordParagraph wordParagraph in comment.Paragraphs.GroupBy(paragraph => paragraph._paragraph).Select(group => group.First())) {
            RtfParagraph paragraph = note.AddParagraph();
            CopyParagraphFormatting(wordParagraph, paragraph, rtfDocument);
            AppendFormattedRuns(wordParagraph, paragraph, rtfDocument, revisionAuthorIndexes);
        }

        if (note.Paragraphs.Count == 0) {
            note.AddParagraph(comment.Text ?? string.Empty);
        }

        return note;
    }

    private static void AppendAnnotationComment(WordParagraph wordRun, RtfNote note, RtfDocument? rtfDocument) {
        Comments comments = WordComment.GetCommentsPart(wordRun._document);
        CommentsEx commentsEx = WordComment.GetCommentsExPart(wordRun._document);
        string commentId = WordComment.GetNewId(wordRun._document, comments);

        var comment = new Comment {
            Id = commentId,
            Author = GetAnnotationAuthor(note),
            Initials = GetAnnotationInitials(note),
            Date = note.Created ?? DateTime.Now
        };

        if (note.Paragraphs.Count == 0) {
            comment.AppendChild(new Paragraph(new Run(new Text(string.Empty))));
        } else {
            foreach (RtfParagraph paragraph in note.Paragraphs) {
                WordParagraph wordParagraph = CreateDetachedWordParagraph(wordRun._document, paragraph, rtfDocument);
                comment.AppendChild((Paragraph)wordParagraph._paragraph.CloneNode(true));
            }
        }

        comments.AppendChild(comment);
        comments.Save();

        commentsEx.AppendChild(new CommentEx { ParaId = WordComment.GetNewParaId(commentsEx) });
        commentsEx.Save();

        AttachCommentReference(wordRun, commentId);
    }

    private static void AttachCommentReference(WordParagraph wordRun, string commentId) {
        Run? firstRun = wordRun._paragraph.GetFirstChild<Run>();
        if (firstRun == null) {
            firstRun = new Run();
            wordRun._paragraph.Append(firstRun);
        }

        wordRun._paragraph.InsertBefore(new CommentRangeStart { Id = commentId }, firstRun);

        Run lastRun = wordRun._paragraph.Elements<Run>().Last();
        OpenXmlElement commentEnd = wordRun._paragraph.InsertAfter(new CommentRangeEnd { Id = commentId }, lastRun);
        wordRun._paragraph.InsertAfter(new Run(new CommentReference { Id = commentId }), commentEnd);
    }

    private static void AttachAnnotation(RtfParagraph paragraph, RtfNote annotation) {
        RtfRun? run = paragraph.Runs.LastOrDefault();
        if (run == null || run.Note != null) {
            run = paragraph.AddText(string.Empty);
        }

        run.Note = annotation;
    }

    private static string GetAnnotationAuthor(RtfNote note) =>
        string.IsNullOrWhiteSpace(note.Author) ? "OfficeIMO" : note.Author!;

    private static string GetAnnotationInitials(RtfNote note) {
        string author = GetAnnotationAuthor(note);
        string initials = new string(author
            .Split(new[] { ' ', '\t', '.', ',', ';', '-' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(part => part[0])
            .Take(3)
            .ToArray());
        return string.IsNullOrWhiteSpace(initials) ? "OI" : initials;
    }
}
