using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using A19 = DocumentFormat.OpenXml.Office2019.Drawing;

namespace OfficeIMO.PowerPoint {
    public abstract partial class PowerPointShape {
        private const string DecorativeExtensionUri = "{C183D7F6-B498-43B3-948B-1728B52AA6E4}";

        /// <summary>Gets or sets the concise accessibility title stored in the shape's non-visual properties.</summary>
        public string? Title {
            get => GetNonVisualDrawingProperties(create: false)?.Title?.Value;
            set {
                NonVisualDrawingProperties drawing = GetNonVisualDrawingProperties(create: true)
                    ?? throw new NotSupportedException("This shape type does not expose non-visual drawing properties.");
                drawing.Title = string.IsNullOrWhiteSpace(value) ? null : value!.Trim();
            }
        }

        /// <summary>Gets or sets the detailed accessibility description. This is the OOXML value also exposed by <see cref="AltText"/>.</summary>
        public string? Description {
            get => AltText;
            set => AltText = string.IsNullOrWhiteSpace(value) ? null : value!.Trim();
        }

        /// <summary>Gets or sets whether assistive technology should treat this shape as decorative.</summary>
        public bool Decorative {
            get {
                NonVisualDrawingProperties? drawing = GetNonVisualDrawingProperties(create: false);
                return drawing?.Descendants<A19.Decorative>().Any(item => item.Val?.Value != false) == true;
            }
            set {
                NonVisualDrawingProperties drawing = GetNonVisualDrawingProperties(create: true)
                    ?? throw new NotSupportedException("This shape type does not expose non-visual drawing properties.");
                A.NonVisualDrawingPropertiesExtensionList? list =
                    drawing.GetFirstChild<A.NonVisualDrawingPropertiesExtensionList>();
                A.NonVisualDrawingPropertiesExtension? extension = list?
                    .Elements<A.NonVisualDrawingPropertiesExtension>()
                    .FirstOrDefault(item => item.GetFirstChild<A19.Decorative>() != null);
                if (value) {
                    if (list == null) {
                        list = new A.NonVisualDrawingPropertiesExtensionList();
                        drawing.Append(list);
                    }
                    if (extension == null) {
                        extension = new A.NonVisualDrawingPropertiesExtension { Uri = DecorativeExtensionUri };
                        list.Append(extension);
                    }
                    extension.RemoveAllChildren<A19.Decorative>();
                    extension.Append(new A19.Decorative { Val = true });
                } else if (extension != null) {
                    extension.Remove();
                    if (list != null && !list.Elements<A.NonVisualDrawingPropertiesExtension>().Any()) list.Remove();
                }
            }
        }

        /// <summary>Zero-based shape-tree order used by PowerPoint accessibility and drawing-order tools.</summary>
        public int ReadingOrder => DrawingOrder;

        /// <summary>Moves the shape to a zero-based shape-tree reading order within its current parent.</summary>
        public void MoveToReadingOrder(int order) {
            OpenXmlElement? parent = Element.Parent;
            if (parent == null) throw new InvalidOperationException("The shape must be attached before changing reading order.");
            List<OpenXmlElement> shapes = parent.ChildElements.Where(IsDrawingElement).ToList();
            if (order < 0 || order >= shapes.Count) throw new ArgumentOutOfRangeException(nameof(order));
            if (shapes[order] == Element) return;
            Element.Remove();
            shapes.Remove(Element);
            if (order >= shapes.Count) parent.Append(Element);
            else parent.InsertBefore(Element, shapes[order]);
        }

        /// <summary>Gets the first explicit BCP 47 language tag found in this shape's text.</summary>
        public string? Language => EnumerateLanguageValues().FirstOrDefault();

        /// <summary>Gets the distinct explicit language tags found in this shape's text.</summary>
        public IReadOnlyList<string> Languages => EnumerateLanguageValues()
            .Distinct(StringComparer.OrdinalIgnoreCase).ToList().AsReadOnly();

        /// <summary>Applies a BCP 47 language tag to every text run and paragraph end marker in this shape.</summary>
        public PowerPointShape SetLanguage(string language) {
            string normalized = PowerPointTextRun.NormalizeLanguage(language)
                ?? throw new ArgumentException("Language cannot be empty.", nameof(language));
            foreach (A.Run run in Element.Descendants<A.Run>()) {
                (run.RunProperties ??= new A.RunProperties()).Language = normalized;
            }
            foreach (A.Paragraph paragraph in Element.Descendants<A.Paragraph>()) {
                A.EndParagraphRunProperties? properties = paragraph.GetFirstChild<A.EndParagraphRunProperties>();
                if (properties == null) {
                    properties = new A.EndParagraphRunProperties();
                    paragraph.Append(properties);
                }
                properties.Language = normalized;
            }
            foreach (A.DefaultRunProperties properties in Element.Descendants<A.DefaultRunProperties>()) {
                properties.Language = normalized;
            }
            return this;
        }

        private IEnumerable<string> EnumerateLanguageValues() {
            foreach (A.RunProperties properties in Element.Descendants<A.RunProperties>()) {
                if (!string.IsNullOrWhiteSpace(properties.Language?.Value)) yield return properties.Language!.Value!;
            }
            foreach (A.EndParagraphRunProperties properties in Element.Descendants<A.EndParagraphRunProperties>()) {
                if (!string.IsNullOrWhiteSpace(properties.Language?.Value)) yield return properties.Language!.Value!;
            }
            foreach (A.DefaultRunProperties properties in Element.Descendants<A.DefaultRunProperties>()) {
                if (!string.IsNullOrWhiteSpace(properties.Language?.Value)) yield return properties.Language!.Value!;
            }
        }
    }
}
