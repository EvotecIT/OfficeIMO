using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO {
    public partial class WordSection {
        public PageOrientationValues PageOrientation {
            get {
                var pageSize = _sectionProperties.GetFirstChild<PageSize>();
                if (pageSize == null) {
                    return PageOrientationValues.Portrait;
                }

                if (pageSize.Orient != null) {
                    return pageSize.Orient.Value;
                }

                return PageOrientationValues.Portrait;
            }
            set {
                var pageSize = _sectionProperties.Descendants<PageSize>().FirstOrDefault();
                if (pageSize == null) {
                    // we need to setup default values for A4 
                    pageSize = PageSizes.A4;
                    _sectionProperties.Append(pageSize);
                }
                if (this.PageOrientation != value) {
                    // changing orientation is not enough, we need to change width with height and vice versa
                    var width = pageSize.Width;
                    var height = pageSize.Height;
                    pageSize.Width = height;
                    pageSize.Height = width;

                    pageSize.Orient = value;
                }
            }
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
    }
}
