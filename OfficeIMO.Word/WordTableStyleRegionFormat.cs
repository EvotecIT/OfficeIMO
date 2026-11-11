using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

/// <summary>A region of a table that a table style can format separately.</summary>
public enum WordTableStyleRegion {
    /// <summary>Every cell of the table.</summary>
    WholeTable,
    /// <summary>The header row. Tables show it when their first-row look is on.</summary>
    FirstRow,
    /// <summary>The total row. Tables show it when their last-row look is on.</summary>
    LastRow,
    /// <summary>The first column.</summary>
    FirstColumn,
    /// <summary>The last column.</summary>
    LastColumn,
    /// <summary>Odd rows when rows are banded.</summary>
    BandedRows,
    /// <summary>Odd columns when columns are banded.</summary>
    BandedColumns
}

/// <summary>The text and cell formatting a table style applies to one <see cref="WordTableStyleRegion"/>.</summary>
/// <remarks>As on <see cref="WordStyle"/>, <see langword="null"/> inherits and assigning it removes the declaration.</remarks>
public sealed class WordTableStyleRegionFormat {
    private readonly WordStyle _style;

    internal WordTableStyleRegionFormat(WordStyle style, WordTableStyleRegion region) {
        _style = style;
        Region = region;
    }

    /// <summary>Gets the region this format applies to.</summary>
    public WordTableStyleRegion Region { get; }

    /// <summary>Gets or sets bold text.</summary>
    public bool? Bold {
        get => WordStyle.ReadOnOff(ReadRun()?.GetFirstChild<Bold>());
        set => SetRun<Bold>(value.HasValue ? (value.Value ? new Bold() : new Bold { Val = false }) : null);
    }

    /// <summary>Gets or sets italic text.</summary>
    public bool? Italic {
        get => WordStyle.ReadOnOff(ReadRun()?.GetFirstChild<Italic>());
        set => SetRun<Italic>(value.HasValue ? (value.Value ? new Italic() : new Italic { Val = false }) : null);
    }

    /// <summary>Gets or sets the text colour.</summary>
    public OfficeColor? Color {
        get => Helpers.ParseNativeColor(WordRunFormatting.ResolveColorHex(ReadRun()?.GetFirstChild<Color>(), _style.Document));
        set => SetRun<Color>(value.HasValue ? new Color { Val = value.Value.ToRgbHex() } : null);
    }

    /// <summary>Gets or sets the cell background colour.</summary>
    public OfficeColor? FillColor {
        get => Helpers.ParseNativeColor(WordRunFormatting.ResolveShadingFillColorHex(ReadCell()?.GetFirstChild<Shading>(), _style.Document));
        set {
            _style.EnsureEditable();
            OpenXmlCompositeElement cell = EnsureCell();
            cell.RemoveAllChildren<Shading>();
            if (value.HasValue) cell.AddChild(new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = value.Value.ToRgbHex() }, true);
            Prune();
        }
    }

    private TableStyleOverrideValues? OverrideType => Region switch {
        WordTableStyleRegion.WholeTable => null,
        WordTableStyleRegion.FirstRow => TableStyleOverrideValues.FirstRow,
        WordTableStyleRegion.LastRow => TableStyleOverrideValues.LastRow,
        WordTableStyleRegion.FirstColumn => TableStyleOverrideValues.FirstColumn,
        WordTableStyleRegion.LastColumn => TableStyleOverrideValues.LastColumn,
        WordTableStyleRegion.BandedRows => TableStyleOverrideValues.Band1Horizontal,
        WordTableStyleRegion.BandedColumns => TableStyleOverrideValues.Band1Vertical,
        _ => throw new ArgumentOutOfRangeException(nameof(Region))
    };

    private TableStyleProperties? FindRegion() {
        TableStyleOverrideValues? type = OverrideType;
        return type == null ? null : _style.Element.Elements<TableStyleProperties>().FirstOrDefault(region => region.Type?.Value == type.Value);
    }

    private OpenXmlCompositeElement? ReadRun() => OverrideType == null
        ? _style.Element.StyleRunProperties
        : FindRegion()?.GetFirstChild<RunPropertiesBaseStyle>();

    private OpenXmlCompositeElement? ReadCell() => OverrideType == null
        ? _style.Element.GetFirstChild<StyleTableCellProperties>()
        : FindRegion()?.GetFirstChild<TableStyleConditionalFormattingTableCellProperties>();

    private TableStyleProperties EnsureRegion() {
        TableStyleProperties? region = FindRegion();
        if (region != null) return region;
        region = new TableStyleProperties { Type = OverrideType!.Value };
        _style.Element.AppendChild(region);
        return region;
    }

    private OpenXmlCompositeElement EnsureRun() {
        if (OverrideType == null) return _style.Element.StyleRunProperties ??= new StyleRunProperties();
        return GetOrAdd<RunPropertiesBaseStyle>(EnsureRegion());
    }

    private static T GetOrAdd<T>(OpenXmlCompositeElement parent) where T : OpenXmlCompositeElement, new() {
        T? child = parent.GetFirstChild<T>();
        if (child != null) return child;
        child = new T();
        parent.AddChild(child, true);
        return child;
    }

    private OpenXmlCompositeElement EnsureCell() {
        if (OverrideType == null) return GetOrAdd<StyleTableCellProperties>(_style.Element);
        return GetOrAdd<TableStyleConditionalFormattingTableCellProperties>(EnsureRegion());
    }

    private void SetRun<T>(T? element) where T : OpenXmlElement {
        _style.EnsureEditable();
        OpenXmlCompositeElement run = EnsureRun();
        run.RemoveAllChildren<T>();
        if (element != null) run.AddChild(element, true);
        if (typeof(T) == typeof(Bold) || typeof(T) == typeof(Italic)) {
            // Word reads complex-script emphasis separately; keep both scripts in step.
            if (typeof(T) == typeof(Bold)) { run.RemoveAllChildren<BoldComplexScript>(); if (element is Bold b) run.AddChild(new BoldComplexScript { Val = b.Val }, true); }
            else { run.RemoveAllChildren<ItalicComplexScript>(); if (element is Italic i) run.AddChild(new ItalicComplexScript { Val = i.Val }, true); }
        }
        Prune();
    }

    private void Prune() {
        foreach (OpenXmlCompositeElement container in new OpenXmlCompositeElement?[] { ReadRun(), ReadCell() }.OfType<OpenXmlCompositeElement>())
            if (!container.HasChildren) container.Remove();
        if (FindRegion() is { HasChildren: false } region) region.Remove();
    }
}
