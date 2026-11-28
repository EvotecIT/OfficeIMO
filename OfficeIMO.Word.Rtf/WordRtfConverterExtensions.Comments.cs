namespace OfficeIMO.Word.Rtf;

public static partial class WordRtfConverterExtensions {
    private static void AppendWordComments(WordParagraph wordParagraph, RtfParagraph paragraph, RtfDocument rtfDocument, Dictionary<string, int> revisionAuthorIndexes) {
        HashSet<string> commentIds = CollectReferencedCommentIds(new[] { wordParagraph.Node() });

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

    private static HashSet<string> CollectReferencedCommentIds(IEnumerable<OfficeXmlNode> roots) {
        var commentIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (OfficeXmlNode root in roots) {
            foreach (string? id in root.ModelDescendants("commentRangeStart").Select(start => start.WAttr("id"))
                         .Concat(root.ModelDescendants("commentReference").Select(reference => reference.WAttr("id")))) {
                if (!string.IsNullOrWhiteSpace(id)) commentIds.Add(id!);
            }
        }
        return commentIds;
    }

    private static RtfNote CreateAnnotation(WordComment comment, RtfDocument rtfDocument, Dictionary<string, int> revisionAuthorIndexes) {
        var note = new RtfNote(RtfNoteKind.Annotation) {
            Id = comment.Id,
            Author = comment.Author,
            Created = comment.DateTime
        };

        foreach (WordParagraph wordParagraph in comment.Paragraphs.GroupBy(paragraph => paragraph.Node()).Select(group => group.First())) {
            RtfParagraph paragraph = note.AddParagraph();
            CopyParagraphFormatting(wordParagraph, paragraph, rtfDocument);
            AppendFormattedRuns(wordParagraph, paragraph, rtfDocument, revisionAuthorIndexes);
        }

        if (note.Paragraphs.Count == 0) {
            note.AddParagraph(comment.Text ?? string.Empty);
        }

        return note;
    }

    private static void AppendAnnotationComment(WordParagraph wordRun, RtfNote note, RtfDocument? rtfDocument, RtfWordNumberingPlan numbering) {
        var paragraphs = new List<OfficeXmlNode>(note.Paragraphs.Count);
        foreach (RtfParagraph paragraph in note.Paragraphs) {
            paragraphs.Add(CreateDetachedWordParagraph(wordRun._document, paragraph, rtfDocument, numbering).Node());
        }

        WordComment comment = WordConverterNodes.CreateComment(wordRun._document, GetAnnotationAuthor(note), GetAnnotationInitials(note), paragraphs);
        comment.DateTime = note.Created ?? DateTime.Now;
        AttachCommentReference(wordRun, comment.Id!);
    }

    private static void AttachCommentReference(WordParagraph wordRun, string commentId) {
        OfficeXmlNode paragraph = wordRun.Node();
        OfficeXmlNode firstRun = paragraph.ModelChild("r");
        if (firstRun.IsNone) {
            firstRun = paragraph.Append(paragraph.CreateElement(WordXml.W, "r"));
        }

        OfficeXmlNode rangeStart = paragraph.CreateElement(WordXml.W, "commentRangeStart");
        rangeStart.SetAttribute(WordXml.W, "id", commentId);
        paragraph.InsertBefore(rangeStart, firstRun);

        OfficeXmlNode lastRun = paragraph.ModelChildren("r").Last();
        OfficeXmlNode rangeEnd = paragraph.CreateElement(WordXml.W, "commentRangeEnd");
        rangeEnd.SetAttribute(WordXml.W, "id", commentId);
        OfficeXmlNode commentEnd = lastRun.InsertAfterSelf(rangeEnd);
        OfficeXmlNode referenceRun = paragraph.CreateElement(WordXml.W, "r");
        OfficeXmlNode reference = referenceRun.CreateElement(WordXml.W, "commentReference");
        reference.SetAttribute(WordXml.W, "id", commentId);
        referenceRun.Append(reference);
        commentEnd.InsertAfterSelf(referenceRun);
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
