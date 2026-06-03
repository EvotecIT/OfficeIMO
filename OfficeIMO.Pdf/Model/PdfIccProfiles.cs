namespace OfficeIMO.Pdf;

/// <summary>
/// Provides built-in ICC profiles used by generated PDF output intents.
/// </summary>
public static class PdfIccProfiles {
    private const string SrgbIec6196621ResourceName = "OfficeIMO.Pdf.Resources.sRGB_IEC61966-2-1_no_black_scaling.icc";

    /// <summary>Canonical output condition identifier for the built-in sRGB IEC61966-2.1 profile.</summary>
    public const string SrgbIec6196621OutputConditionIdentifier = "sRGB IEC61966-2.1";

    /// <summary>
    /// Returns the built-in ICC sRGB IEC61966-2.1 profile bytes.
    /// </summary>
    /// <remarks>
    /// Source profile: International Color Consortium registry,
    /// sRGB_IEC61966-2-1_no_black_scaling.icc. The returned array is a fresh copy.
    /// </remarks>
    public static byte[] SrgbIec6196621 {
        get {
            using (Stream? stream = typeof(PdfIccProfiles).Assembly.GetManifestResourceStream(SrgbIec6196621ResourceName)) {
                if (stream == null) {
                    throw new InvalidOperationException("The built-in sRGB IEC61966-2.1 ICC profile resource is missing.");
                }

                using (var memory = new MemoryStream()) {
                    stream.CopyTo(memory);
                    return memory.ToArray();
                }
            }
        }
    }
}
