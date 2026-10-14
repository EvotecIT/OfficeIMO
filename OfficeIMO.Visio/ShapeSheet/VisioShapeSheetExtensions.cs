using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    /// <summary>Typed access to source-preserved Visio ShapeSheet sections.</summary>
    public static class VisioShapeSheetExtensions {
        /// <summary>Returns source-preserved ShapeSheet sections, including native text rows with current typed edits.</summary>
        public static IReadOnlyList<VisioShapeSheetSection> GetShapeSheetSections(
            this VisioShape shape) {
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            return ReadSections(shape.PreservedNonGeometrySections, shape.CharacterSectionSource,
                shape.ParagraphSectionSource, shape.TextStyle);
        }

        /// <summary>Returns source-preserved ShapeSheet sections, including native text rows with current typed edits.</summary>
        public static IReadOnlyList<VisioShapeSheetSection> GetShapeSheetSections(
            this VisioConnector connector) {
            if (connector == null) throw new ArgumentNullException(nameof(connector));
            return ReadSections(connector.PreservedNonGeometrySections, connector.CharacterSectionSource,
                connector.ParagraphSectionSource, connector.TextStyle);
        }

        /// <summary>Sets a typed ShapeSheet section without disturbing other preserved sections.</summary>
        /// <remarks>Supported fields in a single Character or Paragraph row synchronize the text style.
        /// Complex text rows remain native-owned and are not replaced by global typed formatting.</remarks>
        public static VisioShape SetShapeSheetSection(this VisioShape shape,
            VisioShapeSheetSection section) {
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            SetSection(shape.PreservedNonGeometrySections, shape.PreservedShapeChildren,
                entry => entry.RawElement, entry => entry.Token, element => new VisioShape.PreservedShapeChildEntry(element), section);
            if (TextSectionKind(section.Name) is bool character) {
                bool modeled = VisioDocument.ModelShapeSheetTextSection(shape, section, character);
                TransferTextSectionOwnership(shape.PreservedNonGeometrySections, shape.PreservedShapeChildren,
                    entry => entry.RawElement, entry => entry.Token,
                    token => new VisioShape.PreservedShapeChildEntry(token), character, modeled);
            }
            shape.NativeCellMetadata?.ForgetNullConditions(section.AssignedValueAddresses());
            return shape;
        }

        /// <summary>Sets a typed ShapeSheet section without disturbing other preserved sections.</summary>
        /// <remarks>Supported fields in a single Character or Paragraph row synchronize the text style.
        /// Complex text rows remain native-owned and are not replaced by global typed formatting.</remarks>
        public static VisioConnector SetShapeSheetSection(this VisioConnector connector,
            VisioShapeSheetSection section) {
            if (connector == null) throw new ArgumentNullException(nameof(connector));
            SetSection(connector.PreservedNonGeometrySections, connector.PreservedShapeChildren,
                entry => entry.RawElement, entry => entry.Token, element => new VisioConnector.PreservedShapeChildEntry(element), section);
            if (TextSectionKind(section.Name) is bool character) {
                bool modeled = VisioDocument.ModelShapeSheetTextSection(connector, section, character);
                TransferTextSectionOwnership(connector.PreservedNonGeometrySections, connector.PreservedShapeChildren,
                    entry => entry.RawElement, entry => entry.Token,
                    token => new VisioConnector.PreservedShapeChildEntry(token), character, modeled);
            }
            connector.NativeCellMetadata?.ForgetNullConditions(section.AssignedValueAddresses());
            return connector;
        }

        /// <summary>Removes a source-preserved section and clears any typed formatting owned by that text section.</summary>
        public static bool RemoveShapeSheetSection(this VisioShape shape, string name) {
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            bool removed = RemoveSection(shape.PreservedNonGeometrySections, shape.PreservedShapeChildren, entry => entry.RawElement, name);
            if (TextSectionKind(name) is bool character) {
                removed |= (character ? shape.CharacterSectionSource : shape.ParagraphSectionSource) != null;
                if (!removed) return false;
                if (character) { shape.CharacterSectionSource = null; shape.HasModeledCharSection = false; }
                else { shape.ParagraphSectionSource = null; shape.HasModeledParaSection = false; }
                VisioDocument.ClearShapeSheetTextFormatting(shape.TextStyle, character);
                TransferTextSectionOwnership(shape.PreservedNonGeometrySections, shape.PreservedShapeChildren,
                    entry => entry.RawElement, entry => entry.Token,
                    token => new VisioShape.PreservedShapeChildEntry(token), character, modeled: false);
            }
            return removed;
        }

        /// <summary>Removes a source-preserved section and clears any typed formatting owned by that text section.</summary>
        public static bool RemoveShapeSheetSection(this VisioConnector connector, string name) {
            if (connector == null) throw new ArgumentNullException(nameof(connector));
            bool removed = RemoveSection(connector.PreservedNonGeometrySections, connector.PreservedShapeChildren, entry => entry.RawElement, name);
            if (TextSectionKind(name) is bool character) {
                removed |= (character ? connector.CharacterSectionSource : connector.ParagraphSectionSource) != null;
                if (!removed) return false;
                if (character) { connector.CharacterSectionSource = null; connector.HasModeledCharSection = false; }
                else { connector.ParagraphSectionSource = null; connector.HasModeledParaSection = false; }
                VisioDocument.ClearShapeSheetTextFormatting(connector.TextStyle, character);
                TransferTextSectionOwnership(connector.PreservedNonGeometrySections, connector.PreservedShapeChildren,
                    entry => entry.RawElement, entry => entry.Token,
                    token => new VisioConnector.PreservedShapeChildEntry(token), character, modeled: false);
            }
            return removed;
        }

        private static IReadOnlyList<VisioShapeSheetSection> ReadSections(
            IEnumerable<XElement> sections, VisioTextSectionSource? character,
            VisioTextSectionSource? paragraph, VisioTextStyle? style) {
            var retained = sections.ToList();
            foreach (bool kind in new[] { true, false }) {
                VisioTextSectionSource? source = kind ? character : paragraph;
                if (source == null || retained.Any(section => TextSectionKind((string?)section.Attribute("N")) == kind)) continue;
                XElement? current = VisioDocument.GetRenderTextSection(retained, source, style, kind);
                if (current != null) retained.Add(current);
            }
            return retained
            .Where(element => string.Equals(element.Name.LocalName, "Section",
                StringComparison.OrdinalIgnoreCase))
            .Select(element => new VisioShapeSheetSection(element)).ToList();
        }

        private static bool? TextSectionKind(string? name) =>
            string.Equals(name, "Character", StringComparison.OrdinalIgnoreCase) || string.Equals(name, "Char", StringComparison.OrdinalIgnoreCase) ? true :
            string.Equals(name, "Paragraph", StringComparison.OrdinalIgnoreCase) || string.Equals(name, "Para", StringComparison.OrdinalIgnoreCase) ? false : (bool?)null;

        private static bool SectionNamesMatch(string? left, string? right) =>
            string.Equals(left, right, StringComparison.OrdinalIgnoreCase) ||
            (TextSectionKind(left) is bool kind && TextSectionKind(right) == kind);

        private static void TransferTextSectionOwnership<T>(IList<XElement> sections, IList<T> childOrder,
            Func<T, XElement?> rawElement, Func<T, string?> token, Func<string, T> createEntry,
            bool character, bool modeled) {
            string textToken = character ? "Section:Char" : "Section:Para";
            int? position = null;
            for (int index = childOrder.Count - 1; index >= 0; index--) {
                bool raw = TextSectionKind((string?)rawElement(childOrder[index])?.Attribute("N")) == character;
                if (string.Equals(token(childOrder[index]), textToken, StringComparison.OrdinalIgnoreCase) || (modeled && raw)) {
                    position = index;
                    childOrder.RemoveAt(index);
                }
            }
            if (!modeled) return;
            for (int index = sections.Count - 1; index >= 0; index--)
                if (TextSectionKind((string?)sections[index].Attribute("N")) == character) sections.RemoveAt(index);
            if (position.HasValue) childOrder.Insert(position.Value, createEntry(textToken));
        }

        private static void SetSection<T>(IList<XElement> sections, IList<T> childOrder,
            Func<T, XElement?> rawElement, Func<T, string?> token, Func<XElement, T> createEntry,
            VisioShapeSheetSection section) {
            if (section == null) throw new ArgumentNullException(nameof(section));
            string? name = section.Name;
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Section name cannot be empty.", nameof(section));
            for (int index = 0; index < sections.Count; index++) {
                if (SectionNamesMatch((string?)sections[index].Attribute("N"), name)) {
                    XElement previous = sections[index];
                    XElement replacement = section.ToXElement();
                    sections[index] = replacement;
                    for (int childIndex = 0; childIndex < childOrder.Count; childIndex++) {
                        if (MatchesSection(rawElement(childOrder[childIndex]), previous)) {
                            childOrder[childIndex] = createEntry(replacement);
                            return;
                        }
                    }
                    InsertSection(childOrder, rawElement, token, createEntry, replacement);
                    return;
                }
            }
            XElement added = section.ToXElement();
            sections.Add(added);
            InsertSection(childOrder, rawElement, token, createEntry, added);
        }

        private static bool RemoveSection<T>(IList<XElement> sections, IList<T> childOrder,
            Func<T, XElement?> rawElement, string name) {
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Section name cannot be empty.", nameof(name));
            for (int index = sections.Count - 1; index >= 0; index--) {
                if (SectionNamesMatch((string?)sections[index].Attribute("N"), name)) {
                    XElement removed = sections[index];
                    sections.RemoveAt(index);
                    for (int childIndex = childOrder.Count - 1; childIndex >= 0; childIndex--) {
                        if (MatchesSection(rawElement(childOrder[childIndex]), removed)) childOrder.RemoveAt(childIndex);
                    }
                    return true;
                }
            }
            return false;
        }

        private static bool MatchesSection(XElement? element, XElement section) =>
            element?.Name == VisioShapeSheetSection.VisioNamespace + "Section" &&
            string.Equals((string?)element.Attribute("N"), (string?)section.Attribute("N"), StringComparison.OrdinalIgnoreCase) &&
            (string?)element.Attribute("IX") == (string?)section.Attribute("IX");

        private static void InsertSection<T>(IList<T> childOrder, Func<T, XElement?> rawElement,
            Func<T, string?> token, Func<XElement, T> createEntry, XElement section) {
            // Authored shapes use the normal section collection. Loaded shapes also keep
            // their native child order, so place new sections before terminal content.
            if (childOrder.Count == 0) return;
            int index = 0;
            for (; index < childOrder.Count; index++) {
                string? childToken = token(childOrder[index]);
                if (childToken == "Text" || childToken == "Shapes" ||
                    rawElement(childOrder[index])?.Name.LocalName == "ForeignData") break;
            }
            childOrder.Insert(index, createEntry(section));
        }
    }
}
