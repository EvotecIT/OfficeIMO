using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Provides access to section property settings.
    /// </summary>
    public partial class WordSection {
        /// <summary>
        /// Gets or sets the page orientation of the section.
        /// </summary>
        public PageOrientationValues PageOrientation {
            get => WordPageSizes.GetOrientation(_sectionProperties);
            set => WordPageSizes.SetOrientation(_sectionProperties, value);
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

                if (columns.Space != null) {
                    return int.Parse(columns.Space!.Value!);
                }

                return null;
            }
            set {
                Columns? columns = _sectionProperties.GetFirstChild<Columns>();
                if (columns == null) {
                    columns = new Columns();
                    _sectionProperties.Append(columns);
                }
                columns.Space = value?.ToString();
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

                return null;
            }
            set {
                Columns? columns = _sectionProperties.GetFirstChild<Columns>();
                if (columns == null) {
                    columns = new Columns();
                    _sectionProperties.Append(columns);
                }
                if (value != null) columns.ColumnCount = (Int16Value)value.Value;
            }
        }

        /// <summary>
        /// Gets or sets the footnote properties for the section.
        /// </summary>
        public FootnoteProperties FootnoteProperties {
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
        public EndnoteProperties EndnoteProperties {
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
        public PageNumberType PageNumberType {
            get {
                var pn = _sectionProperties.GetFirstChild<PageNumberType>();
                if (pn == null) {
                    pn = new PageNumberType();
                    _sectionProperties.InsertAt(pn, 0);
                }
                return pn;
            }
            set {
                var existing = _sectionProperties.GetFirstChild<PageNumberType>();
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
    }
}
