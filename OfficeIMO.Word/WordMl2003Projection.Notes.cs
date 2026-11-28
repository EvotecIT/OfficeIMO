using System.Globalization;
using System.Xml.Linq;
using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.Word;

internal sealed partial class WordMl2003Projection {
    private long _nextFootnoteId = 1;
    private long _nextEndnoteId = 1;

    /// <summary>Moves an inline WordML note into its native part without moving its body reference.</summary>
    private void AppendNote(int referenceRun, XElement source) {
        _token.ThrowIfCancellationRequested();
        CheckAttributes(source, "type suppressRef");
        string? type = (string?)source.Attribute(Word + "type");
        string? suppress = (string?)source.Attribute(Word + "suppressRef");
        if (_story.PartName != MainPart || type != null && type != "normal" ||
            suppress != null && suppress is not ("on" or "off" or "1" or "0" or "true" or "false")) {
            Omit(source, "note outside the body, special note type or unsupported reference state");
            return;
        }

        string kind = source.Name.LocalName;
        bool endnote = kind == "endnote";
        bool customMark = suppress != null && IsTrue(suppress);
        if (customMark) {
            string identity = source.Name + "|custom-reference";
            if (!_omissions.ContainsKey(identity)) {
                _omissions.Add(identity, new OfficeDiagnostic("WORDML2003_NOTE_REFERENCE_APPROXIMATION", OfficeDiagnosticSeverity.Warning,
                    "Custom WordML 2003 note reference appearance is unqualified; note content and reference suppression are retained.",
                    location: source.Name.ToString(), format: "WordML 2003", lossKind: OfficeConversionLossKind.Approximation));
            }
        }
        Story previous = _story;
        string name = kind + "s";
        string path = "/word/" + name + ".xml";
        OfficeXmlTree notes = _package.GetParsedXml(path) ?? AddMainPart(name, name, out _);
        string id = (endnote ? _nextEndnoteId++ : _nextFootnoteId++).ToString(CultureInfo.InvariantCulture);
        try {
            _story = new Story(path, notes);
            int note = AppendWord(notes, notes.Root, kind);
            notes.SetAttribute(note, W, "id", id);
            foreach (int block in ProjectBlocks(source)) notes.Append(note, block);
            if (notes.FirstChildOf(note) == NoNode) AppendWord(notes, note, "p");
        } finally { _story = previous; }

        OfficeXmlTree document = previous.Tree;
        int reference = AppendWord(document, referenceRun, kind + "Reference");
        document.SetAttribute(reference, W, "id", id);
        if (customMark) document.SetAttribute(reference, W, "customMarkFollows", "true");
    }

    /// <summary>WordML uses footnoteRef for the automatic mark inside either kind of note.</summary>
    private void AppendNoteMark(int run, XElement source) {
        CheckAttributes(source, "");
        if (_story.PartName == "/word/footnotes.xml") AppendWord(_story.Tree, run, "footnoteRef");
        else if (_story.PartName == "/word/endnotes.xml") AppendWord(_story.Tree, run, "endnoteRef");
        else Omit(source, "note reference mark outside a note story");
    }
}
