using System.Text;

namespace OfficeIMO.Pdf;

/// <summary>Describes the managed renderer's treatment of a PDF feature.</summary>
public enum PdfRenderSupportLevel {
    /// <summary>The feature is projected without a known intentional simplification.</summary>
    Supported,
    /// <summary>The feature is projected with a documented simplification.</summary>
    Simplified,
    /// <summary>The feature is skipped because the managed renderer does not implement it.</summary>
    Unsupported
}

/// <summary>One stable entry in the generated managed-renderer capability manifest.</summary>
public sealed class PdfRenderCapability {
    internal PdfRenderCapability(string id, string category, string feature, PdfRenderSupportLevel supportLevel, string message) {
        Id = id;
        Category = category;
        Feature = feature;
        SupportLevel = supportLevel;
        Message = message;
    }

    /// <summary>Stable machine-readable capability and diagnostic identifier.</summary>
    public string Id { get; }
    /// <summary>Broad capability category.</summary>
    public string Category { get; }
    /// <summary>PDF feature, operator, or resource covered by this entry.</summary>
    public string Feature { get; }
    /// <summary>Current managed-renderer support level.</summary>
    public PdfRenderSupportLevel SupportLevel { get; }
    /// <summary>Stable explanation of the renderer behavior.</summary>
    public string Message { get; }
}

/// <summary>A page-specific occurrence of a simplified or unsupported renderer capability.</summary>
public sealed class PdfRenderCapabilityDiagnostic {
    internal PdfRenderCapabilityDiagnostic(PdfRenderCapability capability, string? subject = null) {
        Capability = capability;
        Subject = subject;
    }

    /// <summary>Manifest entry that defines this diagnostic.</summary>
    public PdfRenderCapability Capability { get; }
    /// <summary>Stable capability identifier.</summary>
    public string Code => Capability.Id;
    /// <summary>Support level associated with the diagnostic.</summary>
    public PdfRenderSupportLevel SupportLevel => Capability.SupportLevel;
    /// <summary>Operator or resource name that triggered this occurrence, when available.</summary>
    public string? Subject { get; }
    /// <summary>Human-readable diagnostic text.</summary>
    public string Message => string.IsNullOrWhiteSpace(Subject)
        ? Capability.Message
        : Capability.Message + " Subject: " + Subject + ".";
}

/// <summary>Generated manifest for the dependency-free managed PDF page renderer.</summary>
public sealed class PdfRenderCapabilityManifest {
    private readonly string _schema = "officeimo.pdf.render-capabilities.v1";

    internal PdfRenderCapabilityManifest(IReadOnlyList<PdfRenderCapability> entries) {
        Entries = entries;
    }

    /// <summary>Schema identifier for serialized manifests.</summary>
    public string Schema => _schema;
    /// <summary>Capabilities in stable identifier order.</summary>
    public IReadOnlyList<PdfRenderCapability> Entries { get; }

    /// <summary>Serializes the generated manifest without adding a JSON runtime dependency.</summary>
    public string ToJson() {
        var builder = new StringBuilder(2048);
        builder.Append("{\"schema\":\"").Append(Schema).Append("\",\"entries\":[");
        for (int i = 0; i < Entries.Count; i++) {
            if (i > 0) builder.Append(',');
            PdfRenderCapability entry = Entries[i];
            builder.Append("{\"id\":\"").Append(EscapeJson(entry.Id))
                .Append("\",\"category\":\"").Append(EscapeJson(entry.Category))
                .Append("\",\"feature\":\"").Append(EscapeJson(entry.Feature))
                .Append("\",\"supportLevel\":\"").Append(entry.SupportLevel)
                .Append("\",\"message\":\"").Append(EscapeJson(entry.Message)).Append("\"}");
        }

        return builder.Append("]}").ToString();
    }

    private static string EscapeJson(string value) {
        var builder = new StringBuilder(value.Length + 8);
        for (int i = 0; i < value.Length; i++) {
            char character = value[i];
            switch (character) {
                case '"': builder.Append("\\\""); break;
                case '\\': builder.Append("\\\\"); break;
                case '\b': builder.Append("\\b"); break;
                case '\f': builder.Append("\\f"); break;
                case '\n': builder.Append("\\n"); break;
                case '\r': builder.Append("\\r"); break;
                case '\t': builder.Append("\\t"); break;
                default:
                    if (character < 0x20) builder.Append("\\u").Append(((int)character).ToString("x4", System.Globalization.CultureInfo.InvariantCulture));
                    else builder.Append(character);
                    break;
            }
        }

        return builder.ToString();
    }
}

/// <summary>Stable capability registry shared by the manifest and per-page diagnostics.</summary>
public static class PdfRenderCapabilities {
    internal const string UnknownOperatorId = "render.operator.unsupported";
    internal const string MiterLimitId = "render.operator.miter-limit-simplified";
    internal const string RenderingIntentId = "render.operator.rendering-intent-simplified";
    internal const string FlatnessId = "render.operator.flatness-simplified";
    internal const string MarkedPointId = "render.operator.marked-point-simplified";
    internal const string Type3MetricsId = "render.operator.type3-metrics-unsupported";
    internal const string FontSubstitutionId = "render.resource.font-substitution";
    internal const string ColorSpaceId = "render.resource.colorspace-unsupported";
    internal const string TilingPatternId = "render.resource.tiling-pattern";
    internal const string BlendModeId = "render.resource.blend-mode";
    internal const string SoftMaskId = "render.resource.soft-mask";
    internal const string UnsupportedTilingPatternId = "render.resource.tiling-pattern-unsupported";
    internal const string UnsupportedBlendModeId = "render.resource.blend-mode-unsupported";
    internal const string UnsupportedSoftMaskId = "render.resource.soft-mask-unsupported";
    internal const string XObjectId = "render.resource.xobject-unsupported";
    internal const string SynthesizedAnnotationAppearanceId = "render.annotation.appearance-synthesized";
    internal const string AnnotationAppearanceId = "render.resource.annotation-appearance-missing";
    internal const string OptionalImageCodecId = "render.resource.image-codec-optional";

    private static readonly IReadOnlyList<PdfRenderCapability> Registry = BuildRegistry();
    private static readonly Dictionary<string, PdfRenderCapability> ById = Registry.ToDictionary(static entry => entry.Id, StringComparer.Ordinal);
    private static readonly PdfRenderCapabilityManifest GeneratedManifest = new PdfRenderCapabilityManifest(Registry);

    /// <summary>Current generated renderer capability manifest.</summary>
    public static PdfRenderCapabilityManifest Current => GeneratedManifest;

    internal static PdfRenderCapability Get(string id) => ById[id];

    private static System.Collections.ObjectModel.ReadOnlyCollection<PdfRenderCapability> BuildRegistry() {
        var entries = new[] {
            Entry("render.annotation.appearance", "annotation", "Annotation appearance streams", PdfRenderSupportLevel.Supported, "Annotation appearance streams are projected when present."),
            Entry(SynthesizedAnnotationAppearanceId, "annotation", "Synthesized annotation appearances", PdfRenderSupportLevel.Simplified, "A bounded appearance was synthesized from supported annotation geometry and style because the annotation did not provide a usable appearance stream."),
            Entry(AnnotationAppearanceId, "annotation", "Annotations without appearance streams", PdfRenderSupportLevel.Unsupported, "The annotation is skipped because it has no usable appearance stream."),
            Entry("render.colorspace.device", "color", "DeviceGray, DeviceRGB, and DeviceCMYK", PdfRenderSupportLevel.Supported, "Device color spaces are projected to shared Drawing colors."),
            Entry("render.colorspace.calibrated", "color", "CalGray, CalRGB, and Lab color spaces", PdfRenderSupportLevel.Simplified, "Calibrated colors are projected through managed device-color approximations."),
            Entry("render.colorspace.icc", "color", "ICCBased color spaces", PdfRenderSupportLevel.Simplified, "ICCBased colors are projected through their declared component count without applying the embedded profile."),
            Entry("render.colorspace.indexed-image", "color", "Indexed image color spaces", PdfRenderSupportLevel.Supported, "Indexed image palettes are decoded and projected through shared Drawing images."),
            Entry(ColorSpaceId, "color", "Indexed content painting, Separation, DeviceN, and NChannel color spaces", PdfRenderSupportLevel.Unsupported, "The selected content-paint color space is not projected by the managed renderer."),
            Entry("render.content.path", "operator", "Path construction, painting, clipping, and graphics-state operators", PdfRenderSupportLevel.Supported, "Core path and graphics-state operators are projected."),
            Entry(FlatnessId, "operator", "i flatness operator", PdfRenderSupportLevel.Simplified, "The flatness value is accepted but the shared Drawing renderer selects its own curve tolerance."),
            Entry(MarkedPointId, "operator", "MP and DP marked-point operators", PdfRenderSupportLevel.Simplified, "Marked-point metadata has no visual projection and is ignored."),
            Entry(MiterLimitId, "operator", "M miter-limit operator", PdfRenderSupportLevel.Simplified, "The miter limit is accepted but the shared Drawing renderer uses its own miter behavior."),
            Entry(RenderingIntentId, "operator", "ri rendering-intent operator", PdfRenderSupportLevel.Simplified, "The rendering intent is accepted but color conversion uses the managed renderer defaults."),
            Entry(Type3MetricsId, "operator", "d0 and d1 Type 3 glyph metric operators", PdfRenderSupportLevel.Unsupported, "Type 3 glyph programs are not executed by the managed renderer."),
            Entry(UnknownOperatorId, "operator", "Unknown content-stream operator", PdfRenderSupportLevel.Unsupported, "The content-stream operator is not recognized by the managed renderer and is skipped."),
            Entry("render.resource.extgstate-alpha", "resource", "ExtGState alpha, line width, dash, cap, and join", PdfRenderSupportLevel.Supported, "Supported ExtGState painting values are projected."),
            Entry(BlendModeId, "resource", "ExtGState blend modes", PdfRenderSupportLevel.Supported, "Standard separable and nonseparable blend modes are projected through shared Drawing effect groups."),
            Entry(UnsupportedBlendModeId, "resource", "Unknown ExtGState blend mode", PdfRenderSupportLevel.Unsupported, "The declared blend mode is not recognized by the managed renderer."),
            Entry("render.resource.embedded-truetype", "resource", "Embedded TrueType and TrueType-outline OpenType fonts", PdfRenderSupportLevel.Supported, "Supported embedded font programs are retained in the shared Drawing scene and used for managed glyph rendering."),
            Entry(FontSubstitutionId, "resource", "Missing or unsupported PDF font programs", PdfRenderSupportLevel.Simplified, "Text is decoded and positioned, but glyph outlines are rendered through a mapped system font when no supported embedded TrueType program is available."),
            Entry(SoftMaskId, "resource", "ExtGState soft masks", PdfRenderSupportLevel.Supported, "Alpha and luminosity Form XObject soft masks are projected through shared Drawing vector masks."),
            Entry(UnsupportedSoftMaskId, "resource", "Unsupported ExtGState soft mask", PdfRenderSupportLevel.Unsupported, "The soft mask does not contain a usable transparency-group Form XObject."),
            Entry("render.resource.form-xobject", "resource", "Form XObjects", PdfRenderSupportLevel.Supported, "Form XObjects are recursively projected with bounded nesting."),
            Entry("render.resource.image-xobject", "resource", "Image XObjects and inline images", PdfRenderSupportLevel.Supported, "Decoded image resources are projected through shared Drawing image elements."),
            Entry(OptionalImageCodecId, "resource", "JPEG 2000 and optional raster codecs", PdfRenderSupportLevel.Simplified, "The image is projected, but PNG output requires a caller-supplied IOfficeRasterImageCodec for formats outside the managed raster engine."),
            Entry("render.resource.shading-pattern", "resource", "Axial and radial shading patterns", PdfRenderSupportLevel.Supported, "Supported axial and radial shadings are projected to shared Drawing gradients."),
            Entry(TilingPatternId, "resource", "Tiling pattern fills", PdfRenderSupportLevel.Supported, "Colored and basic uncolored tiling fills are projected through shared Drawing vector patterns; stroked and text patterns remain outside this capability."),
            Entry(UnsupportedTilingPatternId, "resource", "Unsupported tiling pattern", PdfRenderSupportLevel.Unsupported, "The tiling pattern is malformed or exceeds the managed renderer's bounded pattern contract."),
            Entry(XObjectId, "resource", "Unsupported XObject subtype", PdfRenderSupportLevel.Unsupported, "The XObject subtype is not projected by the managed renderer.")
        };
        Array.Sort(entries, static (left, right) => string.CompareOrdinal(left.Id, right.Id));
        return Array.AsReadOnly(entries);
    }

    private static PdfRenderCapability Entry(string id, string category, string feature, PdfRenderSupportLevel supportLevel, string message) =>
        new PdfRenderCapability(id, category, feature, supportLevel, message);
}
