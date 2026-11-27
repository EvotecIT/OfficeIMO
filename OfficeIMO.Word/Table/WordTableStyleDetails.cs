using System.Globalization;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.Word.OpenXml;

namespace OfficeIMO.Word;
/// <summary>
/// Manages table style settings and borders.
/// </summary>
public partial class WordTableStyleDetails {
    private const int W = OfficeXmlNamespaces.WordprocessingML;

    // Word writes w:left before w:start and w:right before w:end when both are present, although the schema lists
    // w:start first; the table keeps Word's order.
    private static readonly OfficeXmlChildOrder MarginOrder = new(
        new[] { W, W, W, W, W, W },
        new[] { "top", "left", "start", "bottom", "right", "end" });

    private readonly WordTable _table;
    private readonly OfficeXmlNode _properties;

    /// <summary>
    /// Constructor for Table Style Details
    /// </summary>
    /// <param name="table"></param>
    internal WordTableStyleDetails(WordTable table) {
        _table = table ?? throw new ArgumentNullException(nameof(table));
        _table.CheckTableProperties();
        _properties = OfficeXmlNode.FromSdk(table._tableProperties ?? throw new InvalidOperationException("Table properties could not be initialized."));
    }

    /// <summary>
    /// Gets or sets the nonnegative default top cell margin; null removes the direct override.
    /// Values round to whole twips and must fit 0 through 32767 twips.
    /// </summary>
    public global::OfficeIMO.OfficeLength? MarginDefaultTop {
        get => GetMargin("top");
        set => SetMargin("top", value);
    }

    /// <summary>
    /// Gets or sets the nonnegative default bottom cell margin; null removes the direct override.
    /// Values round to whole twips and must fit 0 through 32767 twips.
    /// </summary>
    public global::OfficeIMO.OfficeLength? MarginDefaultBottom {
        get => GetMargin("bottom");
        set => SetMargin("bottom", value);
    }

    /// <summary>
    /// Gets or sets the nonnegative default left cell margin; null removes the direct override.
    /// Values round to whole twips and must fit 0 through 32767 twips.
    /// </summary>
    public global::OfficeIMO.OfficeLength? MarginDefaultLeft {
        get => GetMargin("left");
        set => SetMargin("left", value);
    }

    /// <summary>
    /// Gets or sets the nonnegative default right cell margin; null removes the direct override.
    /// Values round to whole twips and must fit 0 through 32767 twips.
    /// </summary>
    public global::OfficeIMO.OfficeLength? MarginDefaultRight {
        get => GetMargin("right");
        set => SetMargin("right", value);
    }

    /// <summary>
    /// Gets or sets nonnegative cell spacing; null removes the direct override.
    /// Values round to whole twips and must fit 0 through 32767 twips.
    /// </summary>
    public global::OfficeIMO.OfficeLength? CellSpacing {
        get => ReadNonnegativeTwips(_properties.Child(W, "tblCellSpacing").Attribute(W, "w"));
        set {
            short? native = ToMarginTwips(value, nameof(value));
            _table.CheckTableProperties();
            SetWidth("tblCellSpacing", native?.ToString(CultureInfo.InvariantCulture));
        }
    }

    /// <summary>
    /// Gets or sets signed table indentation; null removes the direct override.
    /// Values round to whole twips within the signed 32-bit native range.
    /// </summary>
    public global::OfficeIMO.OfficeLength? TableIndentation {
        get => ReadIndentation(OfficeXmlViewValues.ParseInt32(_properties.Child(W, "tblInd").Attribute(W, "w")));
        set {
            int? native = ToIndentationTwips(value, nameof(value));
            _table.CheckTableProperties();
            SetWidth("tblInd", native?.ToString(CultureInfo.InvariantCulture));
        }
    }

    // Only the table's own w:tblCellMar counts, not one inside extension markup.
    private global::OfficeIMO.OfficeLength? GetMargin(string side) =>
        ReadNonnegativeTwips(_properties.Child(W, "tblCellMar").Child(W, side).Attribute(W, "w"));

    private void SetMargin(string side, global::OfficeIMO.OfficeLength? value) {
        short? native = ToMarginTwips(value, nameof(value));
        _table.CheckTableProperties();
        // The margins element is added even when a margin is cleared, as the table always has.
        OfficeXmlNode margins = _properties.GetOrAddChild(W, "tblCellMar", WTblPrView.Order);
        if (native.HasValue) {
            OfficeXmlNode margin = margins.GetOrAddChild(W, side, MarginOrder);
            margin.SetAttribute(W, "w", native.Value.ToString(CultureInfo.InvariantCulture));
            margin.SetAttribute(W, "type", "dxa");
        } else {
            margins.Child(W, side).Remove();
        }
    }

    private void SetWidth(string localName, string? twips) {
        if (twips == null) {
            _properties.Child(W, localName).Remove();
            return;
        }
        OfficeXmlNode width = _properties.GetOrAddChild(W, localName, WTblPrView.Order);
        width.SetAttribute(W, "w", twips);
        width.SetAttribute(W, "type", "dxa");
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
