using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordTableStyleDetails {

        private readonly WordTable _table;
        private readonly TableProperties _tableProperties;

        /// <summary>
        ///     Constructor for Table Positioning
        /// </summary>
        /// <param name="table"></param>
        internal WordTableStyleDetails(WordTable table) {
            _table = table;
            _tableProperties = table._tableProperties;
        }

        public string MarginDefaultTopWidth {
            get {
                if (_tableProperties != null && _tableProperties.TableStyle != null) {
                    var tableCellMarginDefault = _tableProperties.TableStyle.OfType<TableCellMarginDefault>().FirstOrDefault();
                    if (tableCellMarginDefault != null) {
                        var tableCellMarginTop = tableCellMarginDefault.OfType<TopMargin>().FirstOrDefault();
                        if (tableCellMarginTop != null) {
                            return tableCellMarginTop.Width.Value;
                        }
                    }
                }
                return null;
            }
            set {
                _table.CheckTableProperties();

            }
        }
    }
}
