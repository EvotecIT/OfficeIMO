using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;

namespace OfficeIMO.Word;
/// <summary>
/// Manages table style settings and borders.
/// </summary>
public partial class WordTableStyleDetails {
    private readonly WordTable _table;
    private readonly TableProperties _tableProperties;

    /// <summary>
    /// Constructor for Table Style Details
    /// </summary>
    /// <param name="table"></param>
    internal WordTableStyleDetails(WordTable table) {
        _table = table ?? throw new ArgumentNullException(nameof(table));
        _table.CheckTableProperties();
        _tableProperties = table._tableProperties ?? throw new InvalidOperationException("Table properties could not be initialized.");
    }

    /// <summary>
    /// Gets or sets the nonnegative default top cell margin; null removes the direct override.
    /// Values round to whole twips and must fit 0 through 32767 twips.
    /// </summary>
    public global::OfficeIMO.OfficeLength? MarginDefaultTop {
        get => ReadNonnegativeTwips(_tableProperties.TableCellMarginDefault?.TopMargin?.Width?.InnerText);
        set {
            short? native = ToMarginTwips(value, nameof(value));
            _table.CheckTableProperties();

            if (_tableProperties.TableCellMarginDefault == null) {
                _tableProperties.TableCellMarginDefault = new TableCellMarginDefault();
            }

            if (native.HasValue) {
                if (_tableProperties.TableCellMarginDefault.TopMargin == null) {
                    _tableProperties.TableCellMarginDefault.TopMargin = new TopMargin();
                }
                _tableProperties.TableCellMarginDefault.TopMargin.Width = native.Value.ToString(CultureInfo.InvariantCulture);
                _tableProperties.TableCellMarginDefault.TopMargin.Type = TableWidthUnitValues.Dxa;
            } else {
                if (_tableProperties.TableCellMarginDefault.TopMargin != null) {
                    _tableProperties.TableCellMarginDefault.TopMargin.Remove();
                }
            }
        }
    }

    /// <summary>
    /// Gets or sets the nonnegative default bottom cell margin; null removes the direct override.
    /// Values round to whole twips and must fit 0 through 32767 twips.
    /// </summary>
    public global::OfficeIMO.OfficeLength? MarginDefaultBottom {
        get => ReadNonnegativeTwips(_tableProperties.TableCellMarginDefault?.BottomMargin?.Width?.InnerText);
        set {
            short? native = ToMarginTwips(value, nameof(value));
            _table.CheckTableProperties();

            if (_tableProperties.TableCellMarginDefault == null) {
                _tableProperties.TableCellMarginDefault = new TableCellMarginDefault();
            }

            if (native.HasValue) {
                if (_tableProperties.TableCellMarginDefault.BottomMargin == null) {
                    _tableProperties.TableCellMarginDefault.BottomMargin = new BottomMargin();
                }
                _tableProperties.TableCellMarginDefault.BottomMargin.Width = native.Value.ToString(CultureInfo.InvariantCulture);
                _tableProperties.TableCellMarginDefault.BottomMargin.Type = TableWidthUnitValues.Dxa;
            } else {
                if (_tableProperties.TableCellMarginDefault.BottomMargin != null) {
                    _tableProperties.TableCellMarginDefault.BottomMargin.Remove();
                }
            }
        }
    }

    /// <summary>
    /// Gets or sets the nonnegative default left cell margin; null removes the direct override.
    /// Values round to whole twips and must fit 0 through 32767 twips.
    /// </summary>
    public global::OfficeIMO.OfficeLength? MarginDefaultLeft {
        get => ReadNonnegativeTwips(_tableProperties.TableCellMarginDefault?.TableCellLeftMargin?.Width?.InnerText);
        set {
            short? native = ToMarginTwips(value, nameof(value));
            _table.CheckTableProperties();

            if (_tableProperties.TableCellMarginDefault == null) {
                _tableProperties.TableCellMarginDefault = new TableCellMarginDefault();
            }

            if (native != null) {
                if (_tableProperties.TableCellMarginDefault.TableCellLeftMargin == null) {
                    _tableProperties.TableCellMarginDefault.TableCellLeftMargin = new TableCellLeftMargin();
                }
                _tableProperties.TableCellMarginDefault.TableCellLeftMargin.Width = native;
                _tableProperties.TableCellMarginDefault.TableCellLeftMargin.Type = TableWidthValues.Dxa;
            } else {
                if (_tableProperties.TableCellMarginDefault.TableCellLeftMargin != null) {
                    _tableProperties.TableCellMarginDefault.TableCellLeftMargin.Remove();
                }
            }
        }
    }

    /// <summary>
    /// Gets or sets the nonnegative default right cell margin; null removes the direct override.
    /// Values round to whole twips and must fit 0 through 32767 twips.
    /// </summary>
    public global::OfficeIMO.OfficeLength? MarginDefaultRight {
        get => ReadNonnegativeTwips(_tableProperties.TableCellMarginDefault?.TableCellRightMargin?.Width?.InnerText);
        set {
            short? native = ToMarginTwips(value, nameof(value));
            _table.CheckTableProperties();

            if (_tableProperties.TableCellMarginDefault == null) {
                _tableProperties.TableCellMarginDefault = new TableCellMarginDefault();
            }

            if (native != null) {
                if (_tableProperties.TableCellMarginDefault.TableCellRightMargin == null) {
                    _tableProperties.TableCellMarginDefault.TableCellRightMargin = new TableCellRightMargin();
                }
                _tableProperties.TableCellMarginDefault.TableCellRightMargin.Width = native;
                _tableProperties.TableCellMarginDefault.TableCellRightMargin.Type = TableWidthValues.Dxa;
            } else {
                if (_tableProperties.TableCellMarginDefault.TableCellRightMargin != null) {
                    _tableProperties.TableCellMarginDefault.TableCellRightMargin.Remove();
                }
            }
        }
    }

    /// <summary>
    /// Gets or sets nonnegative cell spacing; null removes the direct override.
    /// Values round to whole twips and must fit 0 through 32767 twips.
    /// </summary>
    public global::OfficeIMO.OfficeLength? CellSpacing {
        get => ReadNonnegativeTwips(_tableProperties.TableCellSpacing?.Width?.Value);
        set {
            short? native = ToMarginTwips(value, nameof(value));
            _table.CheckTableProperties();

            if (native.HasValue) {
                if (_tableProperties.TableCellSpacing == null) {
                    _tableProperties.TableCellSpacing = new TableCellSpacing();
                }
                _tableProperties.TableCellSpacing.Width = native.Value.ToString(CultureInfo.InvariantCulture);
                _tableProperties.TableCellSpacing.Type = TableWidthUnitValues.Dxa;
            } else {
                if (_tableProperties.TableCellSpacing != null) {
                    _tableProperties.TableCellSpacing.Remove();
                }
            }
        }
    }

    /// <summary>
    /// Gets or sets signed table indentation; null removes the direct override.
    /// Values round to whole twips within the signed 32-bit native range.
    /// </summary>
    public global::OfficeIMO.OfficeLength? TableIndentation {
        get => ReadIndentation(_tableProperties.TableIndentation?.Width?.Value);
        set {
            int? native = ToIndentationTwips(value, nameof(value));
            _table.CheckTableProperties();

            if (native.HasValue) {
                if (_tableProperties.TableIndentation == null) {
                    _tableProperties.TableIndentation = new TableIndentation();
                }

                _tableProperties.TableIndentation.Width = native.Value;
                _tableProperties.TableIndentation.Type = TableWidthUnitValues.Dxa;
            } else if (_tableProperties.TableIndentation != null) {
                _tableProperties.TableIndentation.Remove();
            }
        }
    }

    /// <summary>
    /// Get or set the table borders
    /// </summary>
    internal TableBorders? TableBorders {
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

    /// <summary>
    /// Get or set the table cell margin default
    /// </summary>
    internal TableCellMarginDefault? TableCellMarginDefault {
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

    private static global::OfficeIMO.OfficeLength? ReadNonnegativeTwips(string? value) {
        return short.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out short twips) && twips >= 0
            ? global::OfficeIMO.OfficeLength.Emus((long)twips * 635) : (global::OfficeIMO.OfficeLength?)null;
    }

    private static global::OfficeIMO.OfficeLength? ReadIndentation(int? value) => value.HasValue
        ? global::OfficeIMO.OfficeLength.Emus((long)value.Value * 635) : (global::OfficeIMO.OfficeLength?)null;

    private static short? ToMarginTwips(global::OfficeIMO.OfficeLength? value, string parameter) {
        if (!value.HasValue) return null;
        long twips = value.Value.ToTwips();
        if (value.Value < global::OfficeIMO.OfficeLength.Zero || twips > short.MaxValue)
            throw new ArgumentOutOfRangeException(parameter, "The length must fit 0 through 32767 twips.");
        return (short)twips;
    }

    private static int? ToIndentationTwips(global::OfficeIMO.OfficeLength? value, string parameter) {
        if (!value.HasValue) return null;
        long twips = value.Value.ToTwips();
        if (twips < int.MinValue || twips > int.MaxValue)
            throw new ArgumentOutOfRangeException(parameter, "Indentation must fit Word's signed 32-bit twip range.");
        return (int)twips;
    }
}
