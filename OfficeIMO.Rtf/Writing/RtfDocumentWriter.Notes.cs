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

    private static HashSet<RtfNote> CollectReferencedNotes(RtfDocument document) {
        var notes = new HashSet<RtfNote>();
        foreach (IRtfBlock block in document.Blocks) {
            if (block is RtfParagraph paragraph) {
                AddReferencedNotes(paragraph, notes);
            } else if (block is RtfTable table) {
                foreach (RtfTableRow row in table.Rows) {
                    foreach (RtfTableCell cell in row.Cells) {
                        foreach (RtfParagraph cellParagraph in cell.Paragraphs) {
                            AddReferencedNotes(cellParagraph, notes);
                        }
                    }
                }
            }
        }

        foreach (RtfHeaderFooter headerFooter in document.HeaderFooters) {
            foreach (RtfParagraph paragraph in headerFooter.Paragraphs) {
                AddReferencedNotes(paragraph, notes);
            }
        }

        return notes;
    }

    private static void AddReferencedNotes(RtfParagraph paragraph, HashSet<RtfNote> notes) {
        foreach (RtfRun run in paragraph.Runs) {
            if (run.Note != null) {
                notes.Add(run.Note);
            }
        }
    }

    private static void WriteDetachedNotes(StringBuilder builder, RtfDocument document, HashSet<RtfNote> referencedNotes, int? defaultLanguageId, int unicodeSkipCount) {
        foreach (RtfNote note in document.Notes) {
            if (!referencedNotes.Contains(note)) {
                WriteNote(builder, note, defaultLanguageId, unicodeSkipCount);
            }
        }
    }

    private static void WriteNote(StringBuilder builder, RtfNote note, int? defaultLanguageId, int unicodeSkipCount) {
        builder.Append(@"{\");
        builder.Append(note.Kind switch {
            RtfNoteKind.Annotation => "annotation",
            RtfNoteKind.Endnote => "endnote",
            _ => "footnote"
        });
        if (note.Kind == RtfNoteKind.Annotation) {
            WriteAnnotationMetadata(builder, note, unicodeSkipCount);
            builder.Append(@"\chatn");
        }

        foreach (RtfParagraph paragraph in note.Paragraphs) {
            WriteParagraph(builder, paragraph, defaultLanguageId, unicodeSkipCount);
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
