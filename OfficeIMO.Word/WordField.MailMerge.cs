using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Word {
    public partial class WordField {
        /// <summary>Replaces the current field with plain text.</summary>
        /// <param name="text">Text to insert in place of the field.</param>
        public void ReplaceWithText(string text) {
            var model = ResolveNodes();
            OfficeXmlNode reference = _simpleField.Exists ? _simpleField : _runs.FirstOrDefault();
            if (reference.IsNone) return;
            OfficeXmlNode source = _simpleField.Exists ? _simpleField.Child(W, "r") : FindReplacementSourceRun();

            OfficeXmlNode replacement = _paragraph.CreateElement(W, "r");
            OfficeXmlNode properties = source.Child(W, "rPr");
            if (properties.Exists) replacement.Append(properties.Clone());
            OfficeXmlNode content = replacement.CreateElement(W, "t");
            content.SetInnerText(text ?? string.Empty);
            content.SetAttribute(OfficeXmlNamespaces.Xml, "space", "preserve");
            replacement.Append(content);
            replacement = _paragraph.InsertBefore(replacement, reference);

            if (_simpleField.Exists) {
                _simpleField.Remove();
            } else {
                foreach (OfficeXmlNode run in _runs) run.Remove();
                _replaceSdkRuns?.Invoke(replacement);
            }
            _runs.Clear();
            _runs.Add(replacement);
            _authoringParagraph?.UpdateAuthoredSdkField(this);
            model?.MarkChanged();
        }

        private OfficeXmlNode FindReplacementSourceRun() {
            bool afterSeparator = false;
            foreach (OfficeXmlNode run in _runs) {
                string? type = run.Child(W, "fldChar").Attribute(W, "fldCharType");
                if (type == "end") break;
                if (type == "separate") afterSeparator = true;
                else if (afterSeparator) return run;
            }
            return _runs.FirstOrDefault(run => run.Child(W, "rPr").Exists) is { Exists: true } formatted
                ? formatted : _runs[0];
        }
    }
}
