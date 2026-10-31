namespace OfficeIMO.Pdf;

public sealed partial class PdfLoadOptions {
    /// <summary>Snapshots all load settings and resource limits. The caller-supplied AES provider is retained.</summary>
    public PdfLoadOptions Clone() => Snapshot(null, null);

    /// <summary>Copies the complete policy, optionally tightening its input ceiling and overriding the open password.</summary>
    /// <remarks>A larger input ceiling does not relax this policy. A null password retains the existing password.</remarks>
    public PdfLoadOptions Clone(long? maximumInputBytes, string? passwordOverride) =>
        Snapshot(maximumInputBytes.HasValue ? System.Math.Min(Limits.MaxInputBytes, maximumInputBytes.Value) : (long?)null, passwordOverride);

    internal PdfLoadOptions Snapshot(long? maximumInputBytes, string? passwordOverride, PdfPermissionPolicy? permissionPolicyOverride = null) {
        PdfLoadOptions effective = Resolve(this);
        return new PdfLoadOptions {
            ParsingMode = effective.ParsingMode,
            Profile = effective.Profile,
            IncludeImages = effective.IncludeImages,
            IncludeFormWidgets = effective.IncludeFormWidgets,
            PageSelection = effective.PageSelection,
            LayoutOptions = CloneLayoutOptions(effective.LayoutOptions),
            Pipeline = ClonePipelineOptions(effective.Pipeline),
            Limits = effective.Limits.CloneWithMaximumInputBytes(maximumInputBytes),
            Password = passwordOverride ?? effective.Password,
            AesCryptographyProvider = effective.AesCryptographyProvider,
            PermissionPolicy = permissionPolicyOverride ?? effective.PermissionPolicy,
            PreferToUnicode = effective.PreferToUnicode,
            UseWinAnsiFallback = effective.UseWinAnsiFallback,
            AdjustKerningFromTJ = effective.AdjustKerningFromTJ,
            IncludeArtifactText = effective.IncludeArtifactText
        };
    }
}
