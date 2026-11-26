using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    public abstract partial class PowerPointShape {
        private const string DecorativeExtensionUri = "{C183D7F6-B498-43B3-948B-1728B52AA6E4}";
        private const string DecorativeNamespace = "http://schemas.microsoft.com/office/drawing/2017/decorative";

        /// <summary>Gets or sets the concise accessibility title stored in the shape's non-visual properties.</summary>
        public string? Title {
            get => NonVisualDrawingElement(create: false).Attribute("title");
            set {
                RequireNonVisualDrawing().SetAttribute("title", string.IsNullOrWhiteSpace(value) ? null : value!.Trim());
                OwnerPartRef.MarkEdited();
            }
        }

        /// <summary>Gets or sets whether assistive technology should treat this shape as decorative.</summary>
        public bool Decorative {
            get {
                OfficeXmlNode drawing = NonVisualDrawingElement(create: false);
                return !drawing.IsNone && FindDecorative(drawing, all: true).Any(item => OfficeXmlViewValues.ParseBoolean(item.Attribute("val")) != false);
            }
            set {
                OfficeXmlNode drawing = RequireNonVisualDrawing();
                OfficeXmlNode list = drawing.Child(OfficeXmlNamespaces.DrawingML, "extLst");
                OfficeXmlNode extension = default;
                foreach (OfficeXmlNode candidate in list.Children(OfficeXmlNamespaces.DrawingML, "ext")) {
                    if (!candidate.Child(DecorativeNamespace, "decorative").IsNone) {
                        extension = candidate;
                        break;
                    }
                }
                if (value) {
                    if (list.IsNone) list = drawing.Append(drawing.CreateElement(OfficeXmlNamespaces.DrawingML, "extLst"));
                    if (extension.IsNone) {
                        extension = list.Append(list.CreateElement(OfficeXmlNamespaces.DrawingML, "ext"));
                        extension.SetAttribute("uri", DecorativeExtensionUri);
                    }
                    foreach (OfficeXmlNode existing in extension.Children().ToList()) {
                        if (existing.Is(DecorativeNamespace, "decorative")) existing.Remove();
                    }
                    OfficeXmlNode decorative = extension.CreateElement(DecorativeNamespace, "decorative", "adec");
                    decorative.SetAttribute("val", "1");
                    extension.Append(decorative);
                } else if (!extension.IsNone) {
                    extension.Remove();
                    if (list.Children(OfficeXmlNamespaces.DrawingML, "ext").Count == 0) list.Remove();
                }
                OwnerPartRef.MarkEdited();
            }
        }

        private static List<OfficeXmlNode> FindDecorative(OfficeXmlNode element, bool all) {
            var found = new List<OfficeXmlNode>();
            Collect(element);
            return found;

            void Collect(OfficeXmlNode parent) {
                foreach (OfficeXmlNode child in parent.Children()) {
                    if (child.Is(DecorativeNamespace, "decorative")) {
                        found.Add(child);
                        if (!all) return;
                    }
                    Collect(child);
                }
            }
        }

        /// <summary>Zero-based shape-tree order used by PowerPoint accessibility and drawing-order tools.</summary>
        public int ReadingOrder => DrawingOrder;

        /// <summary>Moves the shape to a zero-based shape-tree reading order within its current parent.</summary>
        public void MoveToReadingOrder(int order) {
            OfficeXmlNode element = XmlElement;
            OfficeXmlNode parent = element.Parent;
            if (parent.IsNone) throw new InvalidOperationException("The shape must be attached before changing reading order.");
            List<OfficeXmlNode> shapes = parent.Children().ToList().Where(IsDrawingElement).ToList();
            ArgumentOutOfRangeException.ThrowIfNegative(order);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(order, shapes.Count);
            if (shapes[order].Equals(element)) return;
            element.Remove();
            shapes.Remove(element);
            if (order >= shapes.Count) parent.Append(element);
            else parent.InsertBefore(element, shapes[order]);
            OwnerPartRef.MarkEdited();
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
            OfficeXmlNode element = XmlElement;
            PowerPointTextContent.FormatRuns(element, properties => properties.SetAttribute("lang", normalized));
            foreach (OfficeXmlNode paragraph in element.Descendants(OfficeXmlNamespaces.DrawingML, "p")) {
                OfficeXmlNode properties = paragraph.Child(OfficeXmlNamespaces.DrawingML, "endParaRPr");
                if (properties.IsNone) properties = paragraph.Append(paragraph.CreateElement(OfficeXmlNamespaces.DrawingML, "endParaRPr"));
                properties.SetAttribute("lang", normalized);
            }
            foreach (OfficeXmlNode properties in element.Descendants(OfficeXmlNamespaces.DrawingML, "defRPr")) {
                properties.SetAttribute("lang", normalized);
            }
            OwnerPartRef.MarkEdited();
            return this;
        }

        private IEnumerable<string> EnumerateLanguageValues() {
            OfficeXmlNode element = XmlElement;
            foreach (string name in new[] { "rPr", "endParaRPr", "defRPr" }) {
                foreach (OfficeXmlNode properties in element.Descendants(OfficeXmlNamespaces.DrawingML, name)) {
                    string? language = properties.Attribute("lang");
                    if (!string.IsNullOrWhiteSpace(language)) yield return language!;
                }
            }
        }
    }
}
