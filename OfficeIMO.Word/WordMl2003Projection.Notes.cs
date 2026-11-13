using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

internal sealed partial class WordMl2003Projection {
    private long _nextFootnoteId = 1;
    private long _nextEndnoteId = 1;

    /// <summary>Moves an inline WordML note into its native part without moving its body reference.</summary>
    private void AppendNote(Run referenceRun, XElement source) {
        _token.ThrowIfCancellationRequested();
        CheckAttributes(source, "type suppressRef");
        string? type = (string?)source.Attribute(Word + "type");
        string? suppress = (string?)source.Attribute(Word + "suppressRef");
        if (_story != _main || type != null && type != "normal" ||
            suppress != null && suppress is not ("on" or "off" or "1" or "0" or "true" or "false")) {
            Omit(source, "note outside the body, special note type or unsupported reference state");
            return;
        }

        bool endnote = source.Name == Word + "endnote";
        bool customMark = suppress != null && IsTrue(suppress);
        if (customMark) {
            string identity = source.Name + "|custom-reference";
            if (!_omissions.ContainsKey(identity)) {
                _omissions.Add(identity, new OfficeDiagnostic("WORDML2003_NOTE_REFERENCE_APPROXIMATION", OfficeDiagnosticSeverity.Warning,
                    "Custom WordML 2003 note reference appearance is unqualified; note content and reference suppression are retained.",
                    location: source.Name.ToString(), format: "WordML 2003", lossKind: OfficeConversionLossKind.Approximation));
            }
        }
        OpenXmlPart previous = _story;
        try {
            if (endnote) {
                EndnotesPart? part = _main.EndnotesPart;
                if (part == null) {
                    ReserveRelationship(_main, createsPart: true);
                    part = _main.AddNewPart<EndnotesPart>();
                    part.Endnotes = new Endnotes();
                }
                _story = part;
                long id = _nextEndnoteId++;
                var note = new Endnote(ProjectBlocks(source)) { Id = id };
                if (!note.HasChildren) note.AppendChild(new Paragraph());
                part.Endnotes!.AppendChild(note);
                var reference = new EndnoteReference { Id = id };
                if (customMark) reference.CustomMarkFollows = true;
                referenceRun.AppendChild(reference);
            } else {
                FootnotesPart? part = _main.FootnotesPart;
                if (part == null) {
                    ReserveRelationship(_main, createsPart: true);
                    part = _main.AddNewPart<FootnotesPart>();
                    part.Footnotes = new Footnotes();
                }
                _story = part;
                long id = _nextFootnoteId++;
                var note = new Footnote(ProjectBlocks(source)) { Id = id };
                if (!note.HasChildren) note.AppendChild(new Paragraph());
                part.Footnotes!.AppendChild(note);
                var reference = new FootnoteReference { Id = id };
                if (customMark) reference.CustomMarkFollows = true;
                referenceRun.AppendChild(reference);
            }
        } finally {
            _story = previous;
        }
    }

    /// <summary>WordML uses footnoteRef for the automatic mark inside either kind of note.</summary>
    private void AppendNoteMark(Run run, XElement source) {
        CheckAttributes(source, "");
        if (_story is FootnotesPart) run.AppendChild(new FootnoteReferenceMark());
        else if (_story is EndnotesPart) run.AppendChild(new EndnoteReferenceMark());
        else Omit(source, "note reference mark outside a note story");
    }
}
