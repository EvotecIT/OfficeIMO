namespace OfficeIMO.Rtf.Writing;

internal static partial class RtfDocumentWriter {
    private static void WriteNoteSettings(StringBuilder builder, RtfNoteSettings settings) {
        if (!settings.HasAnyValue) return;

        AppendOptionalTwips(builder, @"\ftnstart", settings.FootnoteStartNumber);
        WriteFootnoteRestart(builder, settings.FootnoteRestart);
        WriteFootnoteNumberFormat(builder, settings.FootnoteNumberFormat);
        WriteFootnotePlacement(builder, settings.FootnotePlacement);

        AppendOptionalTwips(builder, @"\aftnstart", settings.EndnoteStartNumber);
        WriteEndnoteRestart(builder, settings.EndnoteRestart);
        WriteEndnoteNumberFormat(builder, settings.EndnoteNumberFormat);
        WriteEndnotePlacement(builder, settings.EndnotePlacement);
    }

    private static void WriteFootnoteRestart(StringBuilder builder, RtfNoteNumberRestart? restart) {
        if (!restart.HasValue) return;

        builder.Append(restart.Value switch {
            RtfNoteNumberRestart.EachPage => @"\ftnrstpg",
            RtfNoteNumberRestart.EachSection => @"\ftnrestart",
            _ => @"\ftnrstcont"
        });
    }

    private static void WriteEndnoteRestart(StringBuilder builder, RtfNoteNumberRestart? restart) {
        if (!restart.HasValue) return;

        builder.Append(restart.Value == RtfNoteNumberRestart.EachSection
            ? @"\aftnrestart"
            : @"\aftnrstcont");
    }

    private static void WriteFootnoteNumberFormat(StringBuilder builder, RtfNoteNumberFormat? format) {
        if (!format.HasValue) return;

        builder.Append(format.Value switch {
            RtfNoteNumberFormat.LowerLetter => @"\ftnnalc",
            RtfNoteNumberFormat.UpperLetter => @"\ftnnauc",
            RtfNoteNumberFormat.LowerRoman => @"\ftnnrlc",
            RtfNoteNumberFormat.UpperRoman => @"\ftnnruc",
            _ => @"\ftnnar"
        });
    }

    private static void WriteEndnoteNumberFormat(StringBuilder builder, RtfNoteNumberFormat? format) {
        if (!format.HasValue) return;

        builder.Append(format.Value switch {
            RtfNoteNumberFormat.LowerLetter => @"\aftnnalc",
            RtfNoteNumberFormat.UpperLetter => @"\aftnnauc",
            RtfNoteNumberFormat.LowerRoman => @"\aftnnrlc",
            RtfNoteNumberFormat.UpperRoman => @"\aftnnruc",
            _ => @"\aftnnar"
        });
    }

    private static void WriteFootnotePlacement(StringBuilder builder, RtfFootnotePlacement? placement) {
        if (!placement.HasValue) return;

        builder.Append(placement.Value switch {
            RtfFootnotePlacement.BeneathText => @"\ftntj",
            RtfFootnotePlacement.SectionEnd => @"\endnotes",
            RtfFootnotePlacement.DocumentEnd => @"\enddoc",
            _ => @"\ftnbj"
        });
    }

    private static void WriteEndnotePlacement(StringBuilder builder, RtfEndnotePlacement? placement) {
        if (!placement.HasValue) return;

        builder.Append(placement.Value switch {
            RtfEndnotePlacement.DocumentEnd => @"\aenddoc",
            RtfEndnotePlacement.PageBottom => @"\aftnbj",
            RtfEndnotePlacement.BeneathText => @"\aftntj",
            _ => @"\aendnotes"
        });
    }

    private static void WriteDetachedNotes(StringBuilder builder, RtfDocument document, HashSet<RtfNote> referencedNotes, RtfWriteContext context) {
        foreach (RtfNote note in document.Notes) {
            if (!referencedNotes.Contains(note)) {
                WriteNote(builder, note, context);
            }
        }
    }

    private static void WriteNote(StringBuilder builder, RtfNote note, RtfWriteContext context) {
        builder.Append(@"{\");
        builder.Append(note.Kind switch {
            RtfNoteKind.Annotation => "annotation",
            _ => "footnote"
        });
        if (note.Kind == RtfNoteKind.Endnote) builder.Append(@"\ftnalt");
        if (note.Kind == RtfNoteKind.Annotation) {
            WriteAnnotationMetadata(builder, note, context.UnicodeSkipCount);
            builder.Append(@"\chatn");
        }

        for (int index = 0; index < note.Paragraphs.Count; index++) {
            RtfParagraph paragraph = note.Paragraphs[index];
            // Native readers reserve the first note character for its reference marker.
            // Preserve an existing marker when reopening generated RTF instead of duplicating it.
            bool needsReference = index == 0 && note.Kind != RtfNoteKind.Annotation &&
                !(paragraph.Inlines.FirstOrDefault() is RtfGeneratedText generated && generated.Kind == RtfGeneratedTextKind.NoteReference);
            WriteParagraph(builder, note.Paragraphs[index], context,
                terminateParagraph: note.Kind == RtfNoteKind.Annotation || index < note.Paragraphs.Count - 1,
                prefix: needsReference ? new RtfGeneratedText(RtfGeneratedTextKind.NoteReference) : null);
        }

        builder.Append('}');
    }

    private static void WriteAnnotationMetadata(StringBuilder builder, RtfNote note, int unicodeSkipCount) {
        WriteIgnorableTextDestination(builder, "atnid", note.Id, unicodeSkipCount);
        WriteIgnorableTextDestination(builder, "atnauthor", note.Author, unicodeSkipCount);
        WriteIgnorableTimestampDestination(builder, "atntime", note.Created);
    }

    private static void WriteIgnorableTextDestination(StringBuilder builder, string name, string? value, int unicodeSkipCount) {
        if (string.IsNullOrEmpty(value)) return;
        builder.Append(@"{\*\");
        builder.Append(name);
        builder.Append(' ');
        builder.Append(EscapeText(value!, unicodeSkipCount));
        builder.Append('}');
    }

    private static void WriteIgnorableTimestampDestination(StringBuilder builder, string name, DateTime? value) {
        if (!value.HasValue) return;

        DateTime timestamp = value.Value;
        builder.Append(@"{\*\");
        builder.Append(name);
        AppendOptionalTwips(builder, @"\yr", timestamp.Year);
        AppendOptionalTwips(builder, @"\mo", timestamp.Month);
        AppendOptionalTwips(builder, @"\dy", timestamp.Day);
        AppendOptionalTwips(builder, @"\hr", timestamp.Hour);
        AppendOptionalTwips(builder, @"\min", timestamp.Minute);
        AppendOptionalTwips(builder, @"\sec", timestamp.Second);
        builder.Append('}');
    }
}
