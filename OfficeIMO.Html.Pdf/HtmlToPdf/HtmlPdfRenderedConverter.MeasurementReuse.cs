using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

internal static partial class HtmlPdfRenderedConverter {
    private const int MaximumReusedTextMeasurements = 4096;
    private const int MaximumReusedMeasurementTextLength = 256;

    /// <summary>Reuses short built-in measurements within one prepared conversion and font configuration.</summary>
    /// <remarks>Dynamic host-face selection and caller shaping providers retain their ordinary callback behavior.</remarks>
    private static Func<string, OfficeFontInfo, OfficeFontFaceDescriptor, TValue?> ReuseTextMeasurement<TValue>(
        HtmlPdfExportOptions options,
        PdfCore.PdfCreateOptions measurementOptions,
        Func<string, OfficeFontInfo, OfficeFontFaceDescriptor, TValue?> measure) where TValue : struct {
        if (options.TextShapingProvider != null
            || options.ResourcePolicy.AllowSystemFontEmbedding && options.ResourcePolicy.AllowDocumentFontEmbedding) {
            return measure;
        }

        var results = new Dictionary<(string Text, OfficeFontInfo Font, OfficeFontFaceDescriptor Face), TValue>();
        PdfCore.PdfCreateOptions? fontSnapshot = null;
        return (text, font, descriptor) => {
            if (text.Length > MaximumReusedMeasurementTextLength
                || font.FamilyName.Length > MaximumReusedMeasurementTextLength) {
                return measure(text, font, descriptor);
            }

            var key = (text, font, descriptor);
            if (results.TryGetValue(key, out TValue found) && !RefreshFontSnapshot()) return found;

            TValue? value = measure(text, font, descriptor);
            // Fallback planning may have registered a face while measuring. Never
            // keep earlier decisions across that change, even within one request.
            RefreshFontSnapshot();
            if (value.HasValue) {
                // A long table measures more distinct strings than one bounded generation holds.
                // Start a new generation rather than stop reusing, so later width passes still hit.
                if (results.Count >= MaximumReusedTextMeasurements) results.Clear();
                results[key] = value.Value;
            }
            return value;
        };

        bool RefreshFontSnapshot() {
            if (fontSnapshot != null && fontSnapshot.HasSameMeasurementFonts(measurementOptions)) return false;
            results.Clear();
            fontSnapshot = measurementOptions.CreateTextMeasurementSnapshot();
            return true;
        }
    }
}
