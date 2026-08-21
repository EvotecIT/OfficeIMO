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
            "page-relative DrawingML text-box anchors with bounded native offsets and size, wrap, and z-order when forced page-break ownership is unambiguous",
            "solid fill plus policy-approved inline pictures with native crop and alpha",
            "semantic-rich including progress/meter value controls, ruby/MathML and language-scoped text, generated-content, visible multi-block, raw-comment-bearing, bookmark-target, paint-hidden, padded or margined, rounded, root/descendant border or outline, explicitly repeated, authored, or inherited text-alignment/indentation/typography, mixed inline text/picture, multi-column, nested-placement including positioned pictures, aligned or explicitly placed flex/grid, multi-child flex/grid, clipped or scrolling, region or descendant paint-effect, forced-page-break-owned, paged print, multi-page continuous, unrendered-picture, and external-stylesheet-owned regions stay in flow when native ownership would flatten, reorder, reveal, mispaint, break links, or lose page ownership; projection metadata is isolated from authored CSS, while extra background layers, CSS shadows, and policy-rejected pictures are diagnosed"),
        new HtmlEditableLayoutCapabilityContract(
            HtmlConversionTarget.Rtf,
            "bounded positioned and floating regions",
            "first-page RTF paragraph frames with bounded signed native offsets, negative-capable controls, size, and wrap controls when forced page-break ownership is unambiguous",
            "solid frame background plus embedded PNG/JPEG pictures",
            "semantic-rich including progress/meter value controls, ruby/MathML and language-scoped text, generated-content, visible multi-block, raw-comment-bearing, bookmark-target, paint-hidden, padded or margined, rounded, root/descendant border or outline, explicitly repeated, authored, or inherited text-alignment/indentation/typography, mixed inline text/picture, multi-column, nested-placement including positioned pictures, aligned or explicitly placed flex/grid, multi-child flex/grid, clipped or scrolling, region or descendant paint-effect, forced-page-break-owned, paged print, and multi-page continuous regions stay in flow when native ownership would flatten, reorder, reveal, mispaint, break links, or lose page ownership; projection metadata is isolated from authored CSS, while background image layers, picture crop/alpha, shadows, explicit stacking metadata, and unsupported pictures are diagnosed"),
        new HtmlEditableLayoutCapabilityContract(
            HtmlConversionTarget.Excel,
            "bounded positioned, floating, and default-aligned single-content flex and grid regions",
            "editable merged-cell regions plus absolute DrawingML picture anchors without empty-sheet sentinel collisions",
            "cell fills, foreground DrawingML pictures, and picture alpha",
            "semantic-rich including progress/meter value controls, ruby/MathML and language-scoped text, generated-content, visible multi-block, raw-comment-bearing, bookmark-target, paint-hidden, padded or margined, rounded, explicitly repeated, authored, or inherited text-alignment/indentation/typography, mixed inline text/picture, multi-column, nested-placement, aligned or explicitly placed flex/grid, multi-child flex/grid, clipped or scrolling, root/descendant border or outline, and region or descendant paint-effect regions stay in flow; background image layers are omitted to keep editable cell text visible, while worksheet ownership, negative-coordinate clamping, bounds, cell shadows, and unsupported image/effect types are diagnosed"),
        new HtmlEditableLayoutCapabilityContract(
            HtmlConversionTarget.PowerPoint,
            "bounded positioned, floating, and default-aligned single-content flex and grid regions",
            "editable slide text boxes and DrawingML pictures in section-local rendered geometry with failed provisional shape reservations reclaimed and retried",
            "solid fills, supported background/image layers, picture alpha, and one approximated native outer shadow",
            "semantic-rich including progress/meter value controls, ruby/MathML and language-scoped text, generated-content, visible multi-block, raw-comment-bearing, bookmark-target, paint-hidden, padded or margined, rounded, explicitly repeated, authored, or inherited text-alignment/indentation/typography, mixed inline text/picture, multi-column, nested-placement, aligned or explicitly placed flex/grid, multi-child flex/grid, clipped or scrolling, root/descendant border or outline, and region or descendant paint-effect regions stay in flow; explicit section/article containers remain as semantic slide owners while their bounded content projects; collision bounds include only successfully imported native shapes, while collision simplification, every CSS shadow parameter approximation, omitted non-image background layers, additional shadow layers, and unsupported image/effect types are diagnosed")
    });

    /// <summary>Gets every native editable-layout contract in stable destination order.</summary>
    public static IReadOnlyList<HtmlEditableLayoutCapabilityContract> All => Contracts;
}