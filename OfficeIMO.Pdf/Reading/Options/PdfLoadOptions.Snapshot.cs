namespace OfficeIMO.Pdf;

public sealed partial class PdfLoadOptions {
    /// <summary>Snapshots all load settings and resource limits. The caller-supplied AES provider is retained.</summary>
    public PdfLoadOptions Clone() => Snapshot(null, null);

    internal PdfLoadOptions Snapshot(long? maximumInputBytes, string? passwordOverride) {
        PdfLoadOptions effective = Resolve(this);
        return new PdfLoadOptions {
            ParsingMode = effective.ParsingMode,
            Limits = effective.Limits.CloneWithMaximumInputBytes(maximumInputBytes),
            Password = passwordOverride ?? effective.Password,
            AesCryptographyProvider = effective.AesCryptographyProvider,
            PermissionPolicy = effective.PermissionPolicy,
            PreferToUnicode = effective.PreferToUnicode,
            UseWinAnsiFallback = effective.UseWinAnsiFallback,
            AdjustKerningFromTJ = effective.AdjustKerningFromTJ,
            IncludeArtifactText = effective.IncludeArtifactText
        };
    }
}
