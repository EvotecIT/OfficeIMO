using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordTableCell {
        /// <summary>
        /// Gets or sets whether text wraps within the cell.
        /// </summary>
        public bool WrapText {
            get {
                return !IsTextLayoutEnabled(CurrentTableCellProperties?.GetFirstChild<NoWrap>());
            }
            set {
                AddTableCellProperties();
                var current = _tableCellProperties!.GetFirstChild<NoWrap>();
                if (value) {
                    current?.Remove();
                } else {
                    if (current == null) {
                        _tableCellProperties.Append(new NoWrap());
                        NormalizeTableCellPropertiesOrder();
                    } else {
                        current.Val = OnOffOnlyValues.On;
                    }
                }
            }
        }

        /// <summary>
        /// Gets or sets whether text is compressed to fit within the cell width.
        /// </summary>
        public bool FitText {
            get {
                var tcPr = _tableCell.GetFirstChild<TableCellProperties>();
                return IsTextLayoutEnabled(tcPr?.GetFirstChild<TableCellFitText>());
            }
            set {
                AddTableCellProperties();
                var current = _tableCellProperties!.GetFirstChild<TableCellFitText>();
                if (value) {
                    if (current == null) {
                        _tableCellProperties.Append(new TableCellFitText { Val = OnOffOnlyValues.On });
                        NormalizeTableCellPropertiesOrder();
                    } else {
                        current.Val = OnOffOnlyValues.On;
                    }
                } else {
                    current?.Remove();
                }
            }
        }

        /// <summary>
        /// Gets or sets whether the empty cell mark is hidden for this cell.
        /// </summary>
        public bool HideMark {
            get {
                var tcPr = _tableCell.GetFirstChild<TableCellProperties>();
                return IsTextLayoutEnabled(tcPr?.GetFirstChild<HideMark>());
            }
            set {
                AddTableCellProperties();
                var current = _tableCellProperties!.GetFirstChild<HideMark>();
                if (value) {
                    if (current == null) {
                        _tableCellProperties.Append(new HideMark { Val = OnOffOnlyValues.On });
                        NormalizeTableCellPropertiesOrder();
                    } else {
                        current.Val = OnOffOnlyValues.On;
                    }
                } else {
                    current?.Remove();
                }
            }
        }

        private static bool IsTextLayoutEnabled(OnOffOnlyType? setting) =>
            setting != null && (setting.Val == null || setting.Val.Value == OnOffOnlyValues.On);
    }
}
