using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;
public partial class WordTableStyleDetails {
    private readonly WordTable _table;
    private readonly TableProperties _tableProperties;

    /// <summary>
    /// Constructor for Table Style Details
    /// </summary>
    /// <param name="table"></param>
    internal WordTableStyleDetails(WordTable table) {
        _table = table;
        _tableProperties = table._tableProperties;
    }

    public Int16? MarginDefaultTopWidth {
        get {
            if (_tableProperties != null && _tableProperties.TableCellMarginDefault?.TopMargin?.Width != null) {
                return Int16.Parse(_tableProperties.TableCellMarginDefault.TopMargin.Width.Value);
            }
            return null;
        }
        set {
            _table.CheckTableProperties();

            if (_tableProperties.TableCellMarginDefault == null) {
                _tableProperties.TableCellMarginDefault = new TableCellMarginDefault();
            }

            if (value != null) {
                if (_tableProperties.TableCellMarginDefault.TopMargin == null) {
                    _tableProperties.TableCellMarginDefault.TopMargin = new TopMargin();
                }
                _tableProperties.TableCellMarginDefault.TopMargin.Width = value.ToString();
                _tableProperties.TableCellMarginDefault.TopMargin.Type = TableWidthUnitValues.Dxa;
            } else {
                if (_tableProperties.TableCellMarginDefault.TopMargin != null) {
                    _tableProperties.TableCellMarginDefault.TopMargin.Remove();
                }
            }
        }
    }

    public Int16? MarginDefaultBottomWidth {
        get {
            if (_tableProperties?.TableCellMarginDefault?.BottomMargin?.Width != null) {
                return Int16.Parse(_tableProperties.TableCellMarginDefault.BottomMargin.Width.Value);
            }
            return null;
        }
        set {
            _table.CheckTableProperties();

            if (_tableProperties.TableCellMarginDefault == null) {
                _tableProperties.TableCellMarginDefault = new TableCellMarginDefault();
            }

            if (value != null) {
                if (_tableProperties.TableCellMarginDefault.BottomMargin == null) {
                    _tableProperties.TableCellMarginDefault.BottomMargin = new BottomMargin();
                }
                _tableProperties.TableCellMarginDefault.BottomMargin.Width = value.ToString();
                _tableProperties.TableCellMarginDefault.BottomMargin.Type = TableWidthUnitValues.Dxa;
            } else {
                if (_tableProperties.TableCellMarginDefault.BottomMargin != null) {
                    _tableProperties.TableCellMarginDefault.BottomMargin.Remove();
                }
            }
        }
    }

    public Int16? MarginDefaultLeftWidth {
        get {
            if (_tableProperties?.TableCellMarginDefault?.TableCellLeftMargin?.Width != null) {
                return _tableProperties.TableCellMarginDefault.TableCellLeftMargin.Width;
            }
            return null;
        }
        set {
            _table.CheckTableProperties();

            if (_tableProperties.TableCellMarginDefault == null) {
                _tableProperties.TableCellMarginDefault = new TableCellMarginDefault();
            }

            if (value != null) {
                if (_tableProperties.TableCellMarginDefault.TableCellLeftMargin == null) {
                    _tableProperties.TableCellMarginDefault.TableCellLeftMargin = new TableCellLeftMargin();
                }
                _tableProperties.TableCellMarginDefault.TableCellLeftMargin.Width = value;
                _tableProperties.TableCellMarginDefault.TableCellLeftMargin.Type = TableWidthValues.Dxa;
            } else {
                if (_tableProperties.TableCellMarginDefault.TableCellLeftMargin != null) {
                    _tableProperties.TableCellMarginDefault.TableCellLeftMargin.Remove();
                }
            }
        }
    }

    public Int16? MarginDefaultRightWidth {
        get {
            if (_tableProperties?.TableCellMarginDefault?.TableCellRightMargin?.Width != null) {
                return Int16.Parse(_tableProperties.TableCellMarginDefault.TableCellRightMargin.Width);
            }
            return null;
        }
        set {
            _table.CheckTableProperties();

            if (_tableProperties.TableCellMarginDefault == null) {
                _tableProperties.TableCellMarginDefault = new TableCellMarginDefault();
            }

            if (value != null) {
                if (_tableProperties.TableCellMarginDefault.TableCellRightMargin == null) {
                    _tableProperties.TableCellMarginDefault.TableCellRightMargin = new TableCellRightMargin();
                }
                _tableProperties.TableCellMarginDefault.TableCellRightMargin.Width = value;
                _tableProperties.TableCellMarginDefault.TableCellRightMargin.Type = TableWidthValues.Dxa;
            } else {
                if (_tableProperties.TableCellMarginDefault.TableCellRightMargin != null) {
                    _tableProperties.TableCellMarginDefault.TableCellRightMargin.Remove();
                }
            }
        }
    }

    public Int16? CellSpacing {
        get {
            if (_tableProperties?.TableCellSpacing?.Width != null) {
                return Int16.Parse(_tableProperties.TableCellSpacing.Width.Value);
            }
            return null;
        }
        set {
            _table.CheckTableProperties();

            if (value != null) {
                if (_tableProperties.TableCellSpacing == null) {
                    _tableProperties.TableCellSpacing = new TableCellSpacing();
                }
                _tableProperties.TableCellSpacing.Width = value.ToString();
                _tableProperties.TableCellSpacing.Type = TableWidthUnitValues.Dxa;
            } else {
                if (_tableProperties.TableCellSpacing != null) {
                    _tableProperties.TableCellSpacing.Remove();
                }
            }
        }
    }

    public TableBorders TableBorders {
        get {
            return _tableProperties?.TableBorders;
        }
        set {
            _table.CheckTableProperties();
            if (value != null) {
                if (_tableProperties.TableBorders != null) {
                    _tableProperties.TableBorders.Remove();
                }
                _tableProperties.TableBorders = value;
            } else {
                if (_tableProperties.TableBorders != null) {
                    _tableProperties.TableBorders.Remove();
                }
            }
        }
    }

    public TableCellMarginDefault TableCellMarginDefault {
        get {
            return _tableProperties?.TableCellMarginDefault;
        }
        set {
            _table.CheckTableProperties();
            if (value != null) {
                if (_tableProperties.TableCellMarginDefault == null) {
                    _tableProperties.TableCellMarginDefault = value;
                } else {
                    _tableProperties.TableCellMarginDefault.Remove();
                    _tableProperties.TableCellMarginDefault = value;
                }
            } else if (_tableProperties.TableCellMarginDefault != null) {
                _tableProperties.TableCellMarginDefault.Remove();
            }
        }
    }


}