using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;

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
                foreach (OpenXmlElement child in Paragraph.ChildElements) {
                    if (child is A.Run run) {
                        result.Add(new PowerPointParagraphInline(
                            new PowerPointTextRun(run, _slidePart, _ownerPart)));
                    } else if (child is A.Break) {
                        result.Add(new PowerPointParagraphInline(
                            PowerPointParagraphInlineKind.LineBreak, Environment.NewLine));
                    } else if (child is A.Field textField) {
                        result.Add(new PowerPointParagraphInline(
                            PowerPointParagraphInlineKind.Field,
                            textField.Text?.Text ?? textField.InnerText ?? string.Empty,
                            textField.Id?.Value,
                            textField.Type?.Value,
                            new PowerPointTextRun(textField, _slidePart, _ownerPart)));
                    }
                }
                return result;
            }
        }

        /// <summary>Adds an explicit DrawingML line break in paragraph order.</summary>
        public PowerPointParagraph AddLineBreak() {
            var lineBreak = new A.Break();
            A.EndParagraphRunProperties? endProps =
                Paragraph.GetFirstChild<A.EndParagraphRunProperties>();
            if (endProps != null) Paragraph.InsertBefore(lineBreak, endProps);
            else Paragraph.Append(lineBreak);
            return this;
        }

        /// <summary>Adds a dynamic DrawingML field with its current displayed text.</summary>
        public PowerPointParagraph AddField(string displayText, string fieldType, string? fieldId = null) {
            return AddField(displayText, fieldType, fieldId, configure: null);
        }

        /// <summary>Adds a formatted dynamic DrawingML field with its current displayed text.</summary>
        public PowerPointParagraph AddField(string displayText, string fieldType, string? fieldId,
            Action<PowerPointTextRun>? configure) {
            if (string.IsNullOrWhiteSpace(fieldType)) {
                throw new ArgumentException("Field type cannot be empty.", nameof(fieldType));
            }
            var field = new A.Field(new A.Text(displayText ?? string.Empty)) {
                Id = string.IsNullOrWhiteSpace(fieldId)
                    ? Guid.NewGuid().ToString("B").ToUpperInvariant()
                    : fieldId,
                Type = fieldType
            };
            if (configure != null) {
                configure(new PowerPointTextRun(field, _slidePart, _ownerPart));
            }
            A.EndParagraphRunProperties? endProps =
                Paragraph.GetFirstChild<A.EndParagraphRunProperties>();
            if (endProps != null) Paragraph.InsertBefore(field, endProps);
            else Paragraph.Append(field);
            return this;
        }
    }
}
