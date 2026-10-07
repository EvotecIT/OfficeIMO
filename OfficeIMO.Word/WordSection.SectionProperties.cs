using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Provides access to section property settings.
    /// </summary>
    public partial class WordSection {
        /// <summary>
        /// Gets or sets the page orientation of the section.
        /// </summary>
        public OfficePageOrientation PageOrientation {
            get => WordPageSizes.GetOrientation(_sectionProperties).ToOfficeEnum();
            set => WordPageSizes.SetOrientation(_sectionProperties, value.ToOpenXml());
        }
        /// <summary>
        /// Gets or sets spacing between section columns.
        /// </summary>
        public int? ColumnsSpace {
            get {
                Columns? columns = _sectionProperties.GetFirstChild<Columns>();
                if (columns == null) {
                    return null;
                }

                return int.TryParse(columns.Space?.Value, out int spacing) ? spacing : null;
            }
            set {
                Columns? columns = _sectionProperties.GetFirstChild<Columns>();
                if (columns == null) {
                    columns = new Columns();
                    _sectionProperties.Append(columns);
                }
                columns.Space = value.HasValue ? new DocumentFormat.OpenXml.StringValue(value.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)) : null;
            }
        }

        /// <summary>
        /// Gets or sets the number of columns in the section.
        /// </summary>
        public int? ColumnCount {
            get {
                Columns? columns = _sectionProperties.GetFirstChild<Columns>();
                if (columns == null) {
                    return null;
                }

                if (columns.ColumnCount != null) {
                    return (int)columns.ColumnCount!.Value!;
                }

                return columns.EqualWidth?.Value == false && columns.Elements<Column>().Any()
                    ? columns.Elements<Column>().Count() : null;
            }
            set {
                Columns? columns = _sectionProperties.GetFirstChild<Columns>();
                if (value.HasValue && columns?.EqualWidth?.Value == false &&
                    columns.Elements<Column>().Any() && columns.Elements<Column>().Count() != value.Value) {
                    throw new InvalidOperationException("Replace or clear ColumnDefinitions before changing the explicit section column count.");
                }
                if (columns == null) {
                    columns = new Columns();
                    _sectionProperties.Append(columns);
                }
                columns.ColumnCount = value.HasValue ? (Int16Value)value.Value : null;
            }
        }

        /// <summary>
        /// Gets or sets whether a separator line is shown between section columns.
        /// </summary>
        public bool HasColumnSeparator {
            get {
                Columns? columns = _sectionProperties.GetFirstChild<Columns>();
                return columns?.Separator?.Value ?? false;
            }
            set {
                Columns? columns = _sectionProperties.GetFirstChild<Columns>();
                if (columns == null) {
                    columns = new Columns();
                    _sectionProperties.Append(columns);
                }

                columns.Separator = value;
            }
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
