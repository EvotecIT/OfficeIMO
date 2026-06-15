using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Rtf;

public static partial class WordRtfConverterExtensions {
    private static void AppendWordComments(WordParagraph wordParagraph, RtfParagraph paragraph) {
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

            AttachAnnotation(paragraph, CreateAnnotation(comment));
        }
    }

    private static RtfNote CreateAnnotation(WordComment comment) {
        var note = new RtfNote(RtfNoteKind.Annotation) {
            Id = comment.Id,
            Author = comment.Author,
            Created = comment.DateTime
        };
        note.AddParagraph(comment.Text ?? string.Empty);
        return note;
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
