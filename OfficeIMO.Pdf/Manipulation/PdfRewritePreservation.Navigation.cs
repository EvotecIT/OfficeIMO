namespace OfficeIMO.Pdf;

public static partial class PdfRewritePreservation {
    private static void CompareNavigationMetadata(List<PdfRewritePreservationIssue> issues, PdfDocumentInfo original, PdfDocumentInfo rewritten, PdfRewritePreservationOptions options) {
        CompareNamedDestinations(issues, original.NamedDestinations, rewritten.NamedDestinations, options);
        ComparePageLabels(issues, original.PageLabels, rewritten.PageLabels, options);
    }

    private static void CompareNamedDestinations(List<PdfRewritePreservationIssue> issues, IReadOnlyList<PdfNamedDestination> original, IReadOnlyList<PdfNamedDestination> rewritten, PdfRewritePreservationOptions options) {
        if (!options.PreserveNamedDestinations || original.Count != rewritten.Count) {
            return;
        }

        for (int i = 0; i < original.Count; i++) {
            PdfNamedDestination before = original[i];
            PdfNamedDestination after = rewritten[i];
            string prefix = "NamedDestinations[" + i.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]";

            CompareString(issues, prefix + ".Name", before.Name, after.Name);
            CompareNullableInt(issues, prefix + ".PageNumber", before.PageNumber, after.PageNumber);
            CompareNullableDestinationMode(issues, prefix + ".DestinationMode", before.DestinationMode, after.DestinationMode);
            CompareNullableDouble(issues, prefix + ".DestinationTop", before.DestinationTop, after.DestinationTop);
            CompareNullableDouble(issues, prefix + ".DestinationLeft", before.DestinationLeft, after.DestinationLeft);
            CompareNullableDouble(issues, prefix + ".DestinationBottom", before.DestinationBottom, after.DestinationBottom);
            CompareNullableDouble(issues, prefix + ".DestinationRight", before.DestinationRight, after.DestinationRight);
            CompareNullableDouble(issues, prefix + ".DestinationZoom", before.DestinationZoom, after.DestinationZoom);
        }
    }

    private static void ComparePageLabels(List<PdfRewritePreservationIssue> issues, IReadOnlyList<PdfPageLabel> original, IReadOnlyList<PdfPageLabel> rewritten, PdfRewritePreservationOptions options) {
        if (!options.PreservePageLabels || original.Count != rewritten.Count) {
            return;
        }

        for (int i = 0; i < original.Count; i++) {
            PdfPageLabel before = original[i];
            PdfPageLabel after = rewritten[i];
            string prefix = "PageLabels[" + i.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]";

            CompareCounts(issues, prefix + ".StartPageIndex", before.StartPageIndex, after.StartPageIndex, true);
            CompareString(issues, prefix + ".Style", before.Style, after.Style);
            CompareString(issues, prefix + ".Prefix", before.Prefix, after.Prefix);
            CompareNullableInt(issues, prefix + ".StartNumber", before.StartNumber, after.StartNumber);
        }
    }

    private static void CompareNullableDestinationMode(List<PdfRewritePreservationIssue> issues, string feature, PdfOpenActionDestinationMode? expected, PdfOpenActionDestinationMode? actual) {
        if (expected == actual) {
            return;
        }

        issues.Add(CreateIssue(feature, FormatNullableDestinationMode(expected), FormatNullableDestinationMode(actual)));
    }

    private static string FormatNullableDestinationMode(PdfOpenActionDestinationMode? value) {
        return value.HasValue ? value.Value.ToString() : "(null)";
    }

    private static void CompareNullableDouble(List<PdfRewritePreservationIssue> issues, string feature, double? expected, double? actual) {
        if (!expected.HasValue && !actual.HasValue) {
            return;
        }

        if (expected.HasValue && actual.HasValue && Math.Abs(expected.Value - actual.Value) <= GeometryTolerance) {
            return;
        }

        issues.Add(CreateIssue(feature, FormatNullableDouble(expected), FormatNullableDouble(actual)));
    }

    private static string FormatNullableDouble(double? value) {
        return value.HasValue ? FormatDouble(value.Value) : "(null)";
    }
}
