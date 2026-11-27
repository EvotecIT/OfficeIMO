using System.Runtime.CompilerServices;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Word {
    public partial class WordTableOfContent {
        private string? _retainedOwnedPlaceholderText;
        private const string PlaceholderPrefix = "OfficeIMO_TocPlaceholder_";

        private sealed class RetainedPlaceholder {
            internal OfficeXmlNode Part;
        }

        // Query APIs can return the same control through several wrappers, including after disposal.
        private static readonly ConditionalWeakTable<object, RetainedPlaceholder> RetainedPlaceholders = new();

        private string GetPlaceholderText() {
            const int W = TocNamespace;
            OfficeXmlNode placeholder = GetPlaceholderPart();
            if (placeholder.Exists) return ReadPlaceholderText(placeholder);
            if (_document.IsDisposed && _retainedOwnedPlaceholderText != null) return _retainedOwnedPlaceholderText;

            foreach (OfficeXmlNode paragraph in ContentNode.Children(W, "p")) {
                foreach (OfficeXmlNode field in paragraph.Children(W, "fldSimple")) {
                    if (IsGeneratedListInstruction(field.Attribute(W, "instr"))) return FirstText(field).InnerText;
                }
            }
            return string.Empty;
        }

        private OfficeXmlNode GetPlaceholderPart() {
            const int W = TocNamespace;
            if (_document.IsDisposed && SdtNode.SdkElement is { } live) {
                return RetainedPlaceholders.TryGetValue(live, out var retained) ? retained.Part : default;
            }

            string? name = SdtNode.Child(W, "sdtPr").Child(W, "placeholder").Child(W, "docPart").Attribute(W, "val");
            if (name != null) {
                foreach (OfficeXmlNode part in GetGlossaryRoot(create: false).Child(W, "docParts").Children(W, "docPart")) {
                    if (PlaceholderName(part) == name) return RetainPlaceholderPart(part);
                }
            }
            return RetainPlaceholderPart(default);
        }

        private OfficeXmlNode RetainPlaceholderPart(OfficeXmlNode part) {
            if (SdtNode.SdkElement is { } element) {
                RetainedPlaceholders.GetValue(element, _ => new RetainedPlaceholder()).Part = part;
                _retainedOwnedPlaceholderText = null;
            } else {
                _retainedOwnedPlaceholderText = part.Exists ? ReadPlaceholderText(part) : null;
            }
            return part;
        }

        private static string ReadPlaceholderText(OfficeXmlNode part) =>
            string.Concat(WordFieldInventory.EnumerateFieldContent(part.Child(TocNamespace, "docPartBody"))
                .Where(node => node.Is(TocNamespace, "t")).Select(node => node.InnerText));

        private void SetPlaceholderText(string value) {
            const int W = TocNamespace;
            PreservePlaceholderText(value);
            OfficeXmlNode content = ContentNode;
            foreach (OfficeXmlNode paragraph in content.Children(W, "p")) {
                if (IsGeneratedResultStyle(paragraph.Child(W, "pPr").Child(W, "pStyle").Attribute(W, "val"))) return;
            }

            OfficeXmlNode text = default;
            foreach (OfficeXmlNode field in WordFieldInventory.EnumerateFieldContent(content)) {
                if (!field.Is(W, "fldSimple") || !IsGeneratedListInstruction(field.Attribute(W, "instr"))) continue;
                text = FirstText(field);
                break;
            }

            if (text.IsNone) {
                foreach (OfficeXmlNode paragraph in content.Children(W, "p")) {
                    if (!WordFieldInventory.EnumerateFieldContent(paragraph)
                        .Any(code => code.Is(W, "instrText") && IsGeneratedListInstruction(code.InnerText))) continue;
                    text = FirstText(paragraph);
                    if (text.Exists) break;
                }
            }
            if (text.Exists) text.SetInnerText(value ?? string.Empty);
        }

        private static OfficeXmlNode FirstText(OfficeXmlNode root) =>
            WordFieldInventory.EnumerateFieldContent(root).FirstOrDefault(node => node.Is(TocNamespace, "t"));

        private static bool IsGeneratedListInstruction(string? instruction) {
            WordFieldType? type = string.IsNullOrWhiteSpace(instruction) ? null : WordFieldInventory.ParseInstruction(instruction!).FieldType;
            return type == WordFieldType.TOC || type == WordFieldType.Index;
        }

        private static string? PlaceholderName(OfficeXmlNode part) =>
            part.Child(TocNamespace, "docPartPr").Child(TocNamespace, "name").Attribute(TocNamespace, "val");

        private void PreservePlaceholderText(string value) {
            // Word's glossary keeps the preference when the field result is replaced.
            const int W = TocNamespace;
            OfficeXmlNode part = GetPlaceholderPart();
            if (PlaceholderName(part)?.StartsWith(PlaceholderPrefix, StringComparison.Ordinal) != true) {
                OfficeXmlNode glossary = GetGlossaryRoot(create: true);
                OfficeXmlNode parts = glossary.Child(W, "docParts");
                if (parts.IsNone) parts = AppendTocChild(glossary, "docParts");
                part = AppendTocChild(parts, "docPart");
                string name = PlaceholderPrefix + Guid.NewGuid().ToString("N");
                OfficeXmlNode properties = AppendTocChild(part, "docPartPr");
                AppendTocChild(properties, "name").SetAttribute(W, "val", name);
                OfficeXmlNode types = AppendTocChild(properties, "types");
                AppendTocChild(types, "type").SetAttribute(W, "val", "bbPlcHdr");

                OfficeXmlNode sdt = SdtNode.GetOrPrependChild(W, "sdtPr");
                sdt.RemoveChildren(W, "placeholder");
                AppendTocChild(AppendTocChild(sdt, "placeholder"), "docPart").SetAttribute(W, "val", name);
            }
            OfficeXmlNode oldBody = part.Child(W, "docPartBody");
            OfficeXmlNode newBody = part.InsertBefore(part.CreateElement(W, "docPartBody"), oldBody);
            if (oldBody.Exists) oldBody.Remove();
            OfficeXmlNode paragraph = AppendTocChild(newBody, "p");
            paragraph.Append(CreateTocTextRun(paragraph, "t", value ?? string.Empty, preserveSpace: true));
            RetainPlaceholderPart(part);
            RecordOwnedTocEdit();
        }

        internal void CopyGeneratedRefreshStateTo(WordTableOfContent clone) {
            clone._entriesWrittenOnSave = _entriesWrittenOnSave;
        }
    }
}
