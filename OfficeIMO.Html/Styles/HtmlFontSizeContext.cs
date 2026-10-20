namespace OfficeIMO.Html;

/// <summary>CSS size provenance retained across inheritance and generic-family transitions.</summary>
internal readonly struct HtmlFontSizeContext {
    internal HtmlFontSizeContext(double pixels, bool isAbsolute, string absoluteKeyword, bool soleGenericMonospace) {
        Pixels = pixels;
        IsAbsolute = isAbsolute;
        AbsoluteKeyword = absoluteKeyword;
        SoleGenericMonospace = soleGenericMonospace;
    }

    internal double Pixels { get; }
    internal bool IsAbsolute { get; }
    internal string AbsoluteKeyword { get; }
    internal bool SoleGenericMonospace { get; }

    /// <summary>A relative transformation retains the parent absolute status and replaces its keyword identity.</summary>
    internal HtmlFontSizeContext WithRelativeSize(double pixels) =>
        new HtmlFontSizeContext(pixels, IsAbsolute, string.Empty, SoleGenericMonospace);
}

/// <summary>One Browser rendering operation's CSS defaults, separate from physical font selection.</summary>
internal readonly struct HtmlRenderingFontDefaults {
    internal HtmlRenderingFontDefaults(string standardFamily, double standardMedium, double fixedMedium, bool quirksMode) {
        StandardFamily = standardFamily;
        StandardMedium = standardMedium;
        FixedMedium = fixedMedium;
        QuirksMode = quirksMode;
    }

    internal string StandardFamily { get; }
    internal double StandardMedium { get; }
    internal double FixedMedium { get; }
    internal bool QuirksMode { get; }
    internal double Medium(bool monospace) => monospace ? FixedMedium : StandardMedium;

    // Browser lookup belongs here so qualified preference tables do not change Document resolution.
    internal bool TryResolveKeyword(string value, double inherited, bool monospace, out double result) =>
        HtmlBrowserFontSizeKeywords.TryResolve(value, inherited, Medium(monospace), QuirksMode, out result);

    internal HtmlFontSizeContext Initial {
        get {
            bool monospace = HtmlFontSizePolicy.IsSoleGenericMonospace(StandardFamily);
            return new HtmlFontSizeContext(Medium(monospace), false, "medium", monospace);
        }
    }

    // The Browser preset keeps its fixed medium independent of the configured standard medium.
    internal static HtmlRenderingFontDefaults? ForRendering(HtmlRenderOptions options, bool quirksMode = false) =>
        options.UserAgentStyles == HtmlRenderUserAgentStyleMode.Browser
            ? new HtmlRenderingFontDefaults(options.DefaultFontFamily, options.DefaultFontSize, 13D, quirksMode)
            : null;
}
