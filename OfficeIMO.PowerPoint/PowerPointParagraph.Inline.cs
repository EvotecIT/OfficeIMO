using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    /// <summary>Kind of authored inline node in a DrawingML paragraph.</summary>
    public enum PowerPointParagraphInlineKind {
        /// <summary>A normal formatted text run.</summary>
        Run,
        /// <summary>An explicit line break.</summary>
        LineBreak,
        /// <summary>A dynamic DrawingML text field.</summary>
        Field
    }

    /// <summary>An ordered typed inline node from a PowerPoint paragraph.</summary>
    public sealed class PowerPointParagraphInline {
        internal PowerPointParagraphInline(PowerPointTextRun run) {
            Kind = PowerPointParagraphInlineKind.Run;
            Run = run;
            _text = run.Text;
        }

        internal PowerPointParagraphInline(PowerPointParagraphInlineKind kind, string text,
            string? fieldId = null, string? fieldType = null, PowerPointTextRun? run = null) {
            Kind = kind;
            _text = text;
            FieldId = fieldId;
            FieldType = fieldType;
            Run = run;
        }

        /// <summary>Node kind.</summary>
        public PowerPointParagraphInlineKind Kind { get; }
        /// <summary>Displayed text contributed by the node; a line break contributes a newline.</summary>
        public string Text => Run?.Text ?? _text;
        private readonly string _text;
        /// <summary>Formatted text properties for a normal run or dynamic field.</summary>
        public PowerPointTextRun? Run { get; }
        /// <summary>DrawingML field identifier when <see cref="Kind"/> is <see cref="PowerPointParagraphInlineKind.Field"/>.</summary>
        public string? FieldId { get; }
        /// <summary>DrawingML field type when <see cref="Kind"/> is <see cref="PowerPointParagraphInlineKind.Field"/>.</summary>
        public string? FieldType { get; }
    }

    public partial class PowerPointParagraph {
        /// <summary>
        /// Gets normal runs, explicit line breaks, and dynamic fields in authored paragraph order.
        /// </summary>
        public IReadOnlyList<PowerPointParagraphInline> InlineNodes {
            get {
                var result = new List<PowerPointParagraphInline>();
                foreach (OfficeXmlNode child in Element.Children()) {
                    if (child.NamespaceId != OfficeXmlNamespaces.DrawingML) continue;
                    switch (child.LocalName) {
                        case "r":
                            result.Add(new PowerPointParagraphInline(RunFor(child)));
                            break;
                        case "br":
                            result.Add(new PowerPointParagraphInline(PowerPointParagraphInlineKind.LineBreak, Environment.NewLine));
                            break;
                        case "fld":
                            result.Add(new PowerPointParagraphInline(
                                PowerPointParagraphInlineKind.Field,
                                child.ChildText(OfficeXmlNamespaces.DrawingML, "t") ?? child.InnerText,
                                child.Attribute("id"),
                                child.Attribute("type"),
                                RunFor(child)));
                            break;
                    }
                }
                return result;
            }
        }

        /// <summary>Adds an explicit DrawingML line break in paragraph order.</summary>
        public PowerPointParagraph AddLineBreak() {
            OfficeXmlNode paragraph = Element;
            paragraph.InsertBefore(paragraph.CreateElement(OfficeXmlNamespaces.DrawingML, "br"), paragraph.Child(OfficeXmlNamespaces.DrawingML, "endParaRPr"));
            MarkEdited();
            return this;
        }

        /// <summary>Adds a dynamic DrawingML field with its current displayed text.</summary>
        public PowerPointParagraph AddField(string displayText, string fieldType, string? fieldId = null) {
            return AddField(displayText, fieldType, fieldId, configure: null);
        }

        /// <summary>Adds a formatted dynamic DrawingML field with its current displayed text.</summary>
        public PowerPointParagraph AddField(string displayText, string fieldType, string? fieldId,
            Action<PowerPointTextRun>? configure) {
            ArgumentException.ThrowIfNullOrWhiteSpace(fieldType);
            OfficeXmlNode paragraph = Element;
            OfficeXmlNode field = paragraph.CreateElement(OfficeXmlNamespaces.DrawingML, "fld");
            field.SetAttribute("id", string.IsNullOrWhiteSpace(fieldId) ? Guid.NewGuid().ToString("B").ToUpperInvariant() : fieldId);
            field.SetAttribute("type", fieldType);
            field.Append(field.CreateElement(OfficeXmlNamespaces.DrawingML, "t")).SetInnerText(PowerPointTextContent.ToSingleLine(displayText));
            field = paragraph.InsertBefore(field, paragraph.Child(OfficeXmlNamespaces.DrawingML, "endParaRPr"));
            MarkEdited();
            if (configure == null) return this;
            PowerPointTextRun run = RunFor(field);
            try {
                configure(run);
            } catch {
                // A failed configuration leaves no field behind, so a retry doesn't add a second one.
                OfficeXmlNode added = run.Element;
                if (!added.Parent.IsNone) PowerPointTextRun.DiscardContent(added, OwnerPartRef, SlidePartRef);
                throw;
            }
            return this;
        }
    }
}
