using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.OpenXml.Internal;
using System.Globalization;

namespace OfficeIMO.Word {
    // Provides access to section property settings.
    public partial class WordSection {
        /// <summary>
        /// Gets or sets the equal-width column gap as a physical length.
        /// Null omits the stored gap; layout uses its existing 36-point default.
        /// Storage rounds to whole signed 32-bit twips. Native DOC saving accepts
        /// only gaps from zero through 65535 twips and writes its default for an omitted gap
        /// when a column count is set.
        /// Individual unequal-column gaps are controlled by ColumnDefinitions.
        /// </summary>
        /// <exception cref="OverflowException">The rounded gap exceeds the signed 32-bit twip range.</exception>
        public OfficeLength? ColumnsSpace {
            get => int.TryParse(SectionNode.Child(SectionNamespace, "cols").Attribute(SectionNamespace, "space"), out int spacing)
                ? OfficeLength.Points(spacing / 20D) : (OfficeLength?)null;
            set {
                int? spacing = checked((int?)(value?.ToTwips()));
                EditSectionColumns().SetAttribute(SectionNamespace, "space", spacing?.ToString(CultureInfo.InvariantCulture));
            }
        }

        /// <summary>Gets or sets the number of columns in the section.</summary>
        public int? ColumnCount {
            get {
                OfficeXmlNode columns = SectionNode.Child(SectionNamespace, "cols");
                if (columns.Attribute(SectionNamespace, "num") is { } token)
                    return short.Parse(token, CultureInfo.InvariantCulture);
                int count = HasExplicitColumnWidths(columns) ? CountDeclaredColumns(columns) : 0;
                return count > 0 ? count : (int?)null;
            }
            set {
                OfficeXmlNode columns = SectionNode.Child(SectionNamespace, "cols");
                if (value.HasValue && HasExplicitColumnWidths(columns)) {
                    int count = CountDeclaredColumns(columns);
                    if (count != 0 && count != value.Value)
                        throw new InvalidOperationException("Replace or clear ColumnDefinitions before changing the explicit section column count.");
                }
                EditSectionColumns().SetAttribute(SectionNamespace, "num", ((short?)value)?.ToString(CultureInfo.InvariantCulture));
            }
        }

        /// <summary>Gets or sets whether a separator line is shown between section columns.</summary>
        public bool HasColumnSeparator {
            get => SectionNode.Child(SectionNamespace, "cols").Attribute(SectionNamespace, "sep") is "true" or "1" or "on";
            set => EditSectionColumns().SetAttribute(SectionNamespace, "sep", value ? "true" : "false");
        }

        /// <summary>
        /// Gets or sets the footnote properties for the section.
        /// </summary>
        internal FootnoteProperties FootnoteProperties {
            get {
                var fp = _sectionProperties.GetFirstChild<FootnoteProperties>();
                if (fp == null) {
                    fp = new FootnoteProperties();
                    _sectionProperties.InsertAt(fp, 0);
                }
                return fp;
            }
            set {
                var existing = _sectionProperties.GetFirstChild<FootnoteProperties>();
                existing?.Remove();
                if (value != null) {
                    _sectionProperties.InsertAt(value, 0);
                }
            }
        }

        /// <summary>
        /// Gets or sets the endnote properties for the section.
        /// </summary>
        internal EndnoteProperties EndnoteProperties {
            get {
                var ep = _sectionProperties.GetFirstChild<EndnoteProperties>();
                if (ep == null) {
                    ep = new EndnoteProperties();
                    _sectionProperties.InsertAt(ep, 0);
                }
                return ep;
            }
            set {
                var existing = _sectionProperties.GetFirstChild<EndnoteProperties>();
                existing?.Remove();
                if (value != null) {
                    var refNode = _sectionProperties.Elements<FooterReference>().Cast<OpenXmlElement>()
                        .Concat(_sectionProperties.Elements<HeaderReference>()).LastOrDefault();
                    if (refNode != null) {
                        _sectionProperties.InsertAfter(value, refNode);
                    } else {
                        _sectionProperties.InsertAt(value, 0);
                    }
                }
            }
        }

        /// <summary>
        /// Gets or sets the page numbering configuration for the section.
        /// </summary>
        internal PageNumberType PageNumberType {
            get {
                var pn = _sectionProperties.GetFirstChild<PageNumberType>();
                if (pn == null) {
                    pn = new PageNumberType();
                    _sectionProperties.AddChild(pn, true);
                }
                return pn;
            }
            set {
                var existing = _sectionProperties.GetFirstChild<PageNumberType>();
                existing?.Remove();
                if (value != null) {
                    _sectionProperties.AddChild(value, true);
                }
            }
        }
    }
}
