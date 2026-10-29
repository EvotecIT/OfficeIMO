using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfFormFiller {
    /// <summary>Uses the same font encoding and row paint for incremental updates as for fully rewritten fields.</summary>
    internal static PdfStream CreateUpdatedTextWidgetAppearance(Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary resources, PdfDictionary? widgetAppearanceResources, PdfDictionary? widgetPageResources, string defaultAppearance, string value, double width, double height,
        PdfFormFieldStyle style, double fontSize, ChoiceAppearanceData? choiceAppearance,
        ref int nextObjectNumber, int maximumDecodedBytes, CancellationToken cancellationToken) =>
        CreateTextAppearanceStream(objects, resources, widgetAppearanceResources, widgetPageResources, value, width, height, style, defaultAppearance,
            fontSize, null, null, ref nextObjectNumber, choiceAppearance: choiceAppearance,
            maximumDecodedBytes: maximumDecodedBytes, cancellationToken: cancellationToken, addsObjects: false);
}
