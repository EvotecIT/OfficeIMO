namespace OfficeIMO.Html;

/// <summary>One destination's bounded native editable-layout projection contract.</summary>
public sealed class HtmlEditableLayoutCapabilityContract {
    internal HtmlEditableLayoutCapabilityContract(
        HtmlConversionTarget target,
        string nativeRegions,
        string nativeGeometry,
        string nativePaintAndEffects,
        string diagnosticBoundary) {
        Target = target;
        NativeRegions = nativeRegions;
        NativeGeometry = nativeGeometry;
        NativePaintAndEffects = nativePaintAndEffects;
        DiagnosticBoundary = diagnosticBoundary;
    }

    /// <summary>Destination format family.</summary>
    public HtmlConversionTarget Target { get; }
    /// <summary>HTML formatting contexts eligible for native editable projection.</summary>
    public string NativeRegions { get; }
    /// <summary>Destination-native geometry representation.</summary>
    public string NativeGeometry { get; }
    /// <summary>Destination-native paint and picture-effect representation.</summary>
    public string NativePaintAndEffects { get; }
    /// <summary>Stable simplify-or-omit boundary.</summary>
    public string DiagnosticBoundary { get; }
}

/// <summary>Executable native editable-layout contracts shared by documentation and adapters.</summary>
public static class HtmlEditableLayoutCapabilityContracts {
    private static readonly IReadOnlyList<HtmlEditableLayoutCapabilityContract> Contracts = Array.AsReadOnly(new[] {
        new HtmlEditableLayoutCapabilityContract(
            HtmlConversionTarget.Word,
            "bounded positioned and floating regions",
            "page-relative DrawingML text-box anchors with exact offsets, size, wrap, and z-order",
            "solid fill plus policy-approved inline pictures with native crop and alpha",
            "paged print regions stay in semantic flow when page ownership cannot be mapped; extra background layers, CSS shadows, and policy-rejected pictures are diagnosed"),
        new HtmlEditableLayoutCapabilityContract(
            HtmlConversionTarget.Rtf,
            "bounded positioned and floating regions",
            "page-anchored RTF paragraph frames with exact offsets, size, and wrap controls",
            "solid frame background plus embedded PNG/JPEG pictures",
            "paged print regions stay in semantic flow when page ownership cannot be mapped; background image layers, picture crop/alpha, shadows, explicit stacking metadata, and unsupported pictures are diagnosed"),
        new HtmlEditableLayoutCapabilityContract(
            HtmlConversionTarget.Excel,
            "bounded positioned, floating, flex, and grid regions",
            "editable merged-cell regions plus absolute DrawingML picture anchors",
            "cell fills, supported background/image layers, and picture alpha",
            "worksheet ownership, cell shadows, and unsupported image/effect types are diagnosed"),
        new HtmlEditableLayoutCapabilityContract(
            HtmlConversionTarget.PowerPoint,
            "bounded positioned, floating, flex, and grid regions",
            "editable slide text boxes and DrawingML pictures in rendered geometry",
            "solid fills, supported background/image layers, picture alpha, and one native outer shadow",
            "additional shadow layers and unsupported image/effect types are diagnosed")
    });

    /// <summary>Gets every native editable-layout contract in stable destination order.</summary>
    public static IReadOnlyList<HtmlEditableLayoutCapabilityContract> All => Contracts;
}
