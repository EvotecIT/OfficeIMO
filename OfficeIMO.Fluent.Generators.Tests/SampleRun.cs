using OfficeIMO.Fluent;

namespace OfficeIMO.Fluent.Generators.Tests.Sample;

/// <summary>A small text-run receiver shaped like the engines' run models.</summary>
[Fluent]
public class SampleRun {
    /// <summary>Counts writes through the canonical font-size mutator.</summary>
    internal int FontSizeWrites { get; private set; }

    /// <summary>Null inherits from the paragraph style.</summary>
    public bool? Bold { get; set; }

    public bool IsHidden { get; set; }

    public string? FontFamily { get; set; }

    public double? FontSize { get; private set; }

    public SamplePageSetup PageSetup { get; } = new SamplePageSetup();

    public SampleBorder? Border { get; set; }

    public Func<string, string>? TextTransform { get; set; }

    [NoFluent]
    public string? Note { get; set; }

    /// <summary>The single write path for the font size: the fluent method and the engine both go through it.</summary>
    [FluentMutator("FontSize")]
    internal void SetFontSizeCore(double points) {
        if (points <= 0 || points > 1638) throw new ArgumentOutOfRangeException(nameof(points), points, "Font size must be between 0 and 1638 points.");
        FontSize = points;
        FontSizeWrites++;
    }
}

/// <summary>A derived receiver: chains on it must keep this type.</summary>
public sealed class HeadingRun : SampleRun {
    public int Level { get; set; }
}

public sealed class SamplePageSetup {
    public double Width { get; set; } = 595;
    public bool Landscape { get; set; }
}

public sealed class SampleBorder {
    public string Style { get; set; } = "single";
    public double Width { get; set; } = 0.5;
}
