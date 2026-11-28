using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.Word.OpenXml;

namespace OfficeIMO.Word;

public partial class WordParagraph {
    private WordField? _authoredField;
    private SimpleField? _simpleFieldValue;

    // Internal SDK callers that format the result run can explicitly resolve the field.
    internal SimpleField? _simpleField {
        get {
            if (_ownedTree != null && _authoredField?.CurrentNodes().SimpleField.Exists == true) ResolveOwnedToSdk();
            return _simpleFieldValue;
        }
        set => _simpleFieldValue = value;
    }

    internal static WordParagraph CreateOwnedField(WordDocument document, OfficeXmlNode paragraph,
        OfficeXmlNode simple, List<OfficeXmlNode>? runs) {
        WordParagraph wrapper = CreateOwned(document, paragraph.Tree!, paragraph.Node);
        var field = new WordField(document, paragraph, simple, runs);
        field.SetAuthoringParagraph(wrapper);
        wrapper._authoredField = field;
        wrapper._ownedFieldProjection = true;
        return wrapper;
    }

    private bool _ownedFieldProjection;

    internal OfficeXmlNode EditParagraphForField() {
        if (OwnedModel is { } model) {
            WordOwnedEdits.PrepareBodyEdit(model);
            return new OfficeXmlNode(_ownedTree!, _ownedParagraphNode);
        }
        return OfficeXmlNode.FromSdk(_paragraph);
    }

    internal void RememberAuthoredField(WordField field) {
        _authoredField = field;
        if (_ownedTree == null) UpdateAuthoredSdkField(field);
    }

    internal void UpdateAuthoredSdkField(WordField field) {
        if (!ReferenceEquals(_authoredField, field) || _ownedTree != null) return;
        (OfficeXmlNode simple, List<OfficeXmlNode> runs) = field.CurrentNodes();
        _simpleField = (SimpleField?)simple.SdkElement;
        _runs = _simpleField == null ? runs.Select(run => (Run)run.SdkElement!).ToList() : null;
    }

    internal static WordParagraph WrapSdkField(WordDocument document, OfficeXmlNode paragraph, OfficeXmlNode simple,
        List<OfficeXmlNode> runs) => simple.SdkElement is SimpleField field
        ? new WordParagraph(document, (Paragraph)paragraph.SdkElement!, field)
        : new WordParagraph(document, (Paragraph)paragraph.SdkElement!, runs.Select(run => (Run)run.SdkElement!).ToList());

    internal static WordField WrapSdkField(WordDocument document, Paragraph paragraph, SimpleField? simple, List<Run>? runs) =>
        new WordField(document, OfficeXmlNode.FromSdk(paragraph), OfficeXmlNode.FromSdk(simple),
            runs?.Select(OfficeXmlNode.FromSdk).ToList(),
            runs == null ? null : replacement => {
                runs.Clear();
                runs.Add((Run)replacement.SdkElement!);
            });
}
