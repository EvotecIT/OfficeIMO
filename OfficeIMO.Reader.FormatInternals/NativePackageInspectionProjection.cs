using System.Globalization;
using System.Linq;

namespace OfficeIMO.Reader.FormatInternals;

internal static class NativePackageInspectionProjection {
    internal static void Add(ReaderInspectionMetadataBuilder builder, OfficePackageSecurityReport? report, string? scope, string? path) {
        var location = new ReaderLocation { Path = path, SourceBlockKind = "package" };
        if (report == null) {
            builder.Coverage("package-security", "not-captured", "The owning import route did not capture structural source-package observations.", location);
            return;
        }
        bool complete = report.ContainerKind != OfficePackageContainerKind.Unknown && !report.Findings.Any(finding =>
            finding.Rule is OfficePackageSecurityRule.PackageSize or OfficePackageSecurityRule.PartCount or
                OfficePackageSecurityRule.PartSize or OfficePackageSecurityRule.TotalUncompressedSize or
                OfficePackageSecurityRule.CompressionRatio or OfficePackageSecurityRule.DuplicatePartName or
                OfficePackageSecurityRule.UnsafePartName or OfficePackageSecurityRule.MalformedPackage or
                OfficePackageSecurityRule.MalformedRelationship);
        builder.Add(OfficeDocumentInspectionCategories.PackageSecurity, "Package structure", report.ContainerKind.ToString(), "container-kind", location,
            attributes: ReaderInspectionMetadataBuilder.Attributes(
                ("scope", scope), ("packageBytes", report.PackageBytes.ToString(CultureInfo.InvariantCulture)),
                ("partCount", report.PartCount.ToString(CultureInfo.InvariantCulture)), ("macroPartCount", report.MacroPartCount.ToString(CultureInfo.InvariantCulture)),
                ("embeddedPayloadPartCount", report.EmbeddedPayloadPartCount.ToString(CultureInfo.InvariantCulture)),
                ("activeXPartCount", report.ActiveXPartCount.ToString(CultureInfo.InvariantCulture)),
                ("externalRelationshipCount", report.ExternalRelationshipCount.ToString(CultureInfo.InvariantCulture)),
                ("digitalSignaturePartCount", report.DigitalSignaturePartCount.ToString(CultureInfo.InvariantCulture)),
                ("totalUncompressedBytes", report.TotalUncompressedBytes.ToString(CultureInfo.InvariantCulture)),
                ("largestPartBytes", report.LargestPartBytes.ToString(CultureInfo.InvariantCulture)),
                ("highestCompressionRatio", report.HighestCompressionRatio.ToString("R", CultureInfo.InvariantCulture)),
                ("validWithinPolicy", report.IsValid ? "true" : "false"),
                ("completeWithinStructuralProfile", complete ? "true" : "false"), ("executed", "false")));
        foreach (OfficePackageSecurityFinding finding in report.Findings) {
            builder.Add(OfficeDocumentInspectionCategories.PackageSecurity, finding.Rule.ToString(), finding.Message, "finding", location, finding.PartName,
                ReaderInspectionMetadataBuilder.Attributes(("severity", finding.Severity.ToString()), ("partName", finding.PartName),
                    ("observedValue", finding.ObservedValue?.ToString("R", CultureInfo.InvariantCulture)), ("limit", finding.Limit?.ToString("R", CultureInfo.InvariantCulture))));
        }
        builder.Coverage("package-security", complete ? "bounded-structural" : "partial-structural", "Shared source-package directory/content-type/relationship profile. Native bound and malformed-part findings identify possible omissions; presence of macro/signature parts does not establish execution or signature validity.", location);
    }
}
