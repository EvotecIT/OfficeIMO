using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordSection {
        public PageOrientationValues PageOrientation {
            get => WordPageSizes.GetOrientation(_sectionProperties);
            set => WordPageSizes.SetOrientation(_sectionProperties, value);
        }
        public int? ColumnsSpace {
            get {
                Columns columns = _sectionProperties.GetFirstChild<Columns>();
                if (columns == null) {
                    return null;
                }

                if (columns.Space != null) {
                    return int.Parse(columns.Space);
                }

                return null;
            }
            set {
                Columns columns = _sectionProperties.GetFirstChild<Columns>();
                if (columns == null) {
                    columns = new Columns();
                    _sectionProperties.Append(columns);
                }
                columns.Space = value.ToString();
            }
        }

        public int? ColumnCount {
            get {
                Columns columns = _sectionProperties.GetFirstChild<Columns>();
                if (columns == null) {
                    return null;
                }

                if (columns.ColumnCount != null) {
                    return int.Parse(columns.ColumnCount);
                }

                return null;
            }
            set {
                Columns columns = _sectionProperties.GetFirstChild<Columns>();
                if (columns == null) {
                    columns = new Columns();
                    _sectionProperties.Append(columns);
                }
                if (value != null) columns.ColumnCount = (Int16Value)value.Value;
            }
        }

        public FootnoteProperties FootnoteProperties {
            get {
                return _sectionProperties.GetFirstChild<FootnoteProperties>();
            }
            set {
                var existing = _sectionProperties.GetFirstChild<FootnoteProperties>();
                existing?.Remove();
                if (value != null) {
                    _sectionProperties.Append(value);
                }
            }
        }

        public EndnoteProperties EndnoteProperties {
            get {
                return _sectionProperties.GetFirstChild<EndnoteProperties>();
            }
            set {
                var existing = _sectionProperties.GetFirstChild<EndnoteProperties>();
                existing?.Remove();
                if (value != null) {
                    _sectionProperties.Append(value);
                }
            }
        }
    }
}
