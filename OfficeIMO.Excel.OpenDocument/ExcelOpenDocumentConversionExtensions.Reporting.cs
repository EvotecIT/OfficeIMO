using OfficeIMO.OpenDocument;

namespace OfficeIMO.Excel.OpenDocument;

public static partial class ExcelOpenDocumentConversionExtensions {
    private static void AddConverted(OdfConversionReport report, string feature, int count) {
        if (count > 0) report.Add(feature, OdfConversionMappingStatus.Converted, count);
    }

    private static void AddUnsupported(OdfConversionReport report, string feature, int count, string? message) {
        if (count > 0) report.Add(feature, OdfConversionMappingStatus.Unsupported, count, message);
    }

    private static void AddUnmappedOdfFindings(OdfFeatureReport features, OdfConversionReport report,
        int formulas, int validations, int hyperlinks, int annotations, int namedRanges) {
        foreach (OdfFeatureDiagnostic diagnostic in features.Diagnostics) {
            report.Add("source-inspection", OdfConversionMappingStatus.Unsupported, 1,
                diagnostic.Code + " in " + diagnostic.PartPath + ": " + diagnostic.Message);
        }
        int remainingFormulas = formulas, remainingValidations = validations, remainingHyperlinks = hyperlinks;
        int remainingAnnotations = annotations, remainingNamedRanges = namedRanges;
        foreach (OdfFeatureFinding finding in features.Findings) {
            int handled = 0;
            if (finding.Name == "spreadsheet-formulas") handled = Consume(ref remainingFormulas, finding.Count);
            else if (finding.Name == "spreadsheet-validations") handled = Consume(ref remainingValidations, finding.Count);
            else if (finding.Name == "external-links") handled = Consume(ref remainingHyperlinks, finding.Count);
            else if (finding.Name == "annotations") handled = Consume(ref remainingAnnotations, finding.Count);
            else if (finding.Name == "spreadsheet-named-ranges") handled = Consume(ref remainingNamedRanges, finding.Count);
            int remaining = Math.Max(0, finding.Count - handled);
            if (remaining > 0) report.Add("source-" + finding.Name, OdfConversionMappingStatus.Unsupported, remaining,
                "The source feature is not represented by the XLSX conversion surface.");
        }
    }

    private static int Consume(ref int available, int requested) {
        int consumed = Math.Min(available, requested);
        available -= consumed;
        return consumed;
    }
}
