using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint.LegacyPpt;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointPresentation {
        private static void ValidatePackageSecurity(byte[] bytes,
            PowerPointLoadOptions options) {
            if (options.PackageSecurity != null) {
                OfficePackageSecurityInspector.Validate(bytes,
                    options.PackageSecurity);
            }
        }

        private static void ValidateLegacyContentSecurity(
            LegacyPptPresentation legacy, PowerPointLoadOptions options) {
            OfficePackageSecurityOptions? security = options.PackageSecurity;
            if (security == null) return;

            if (security.Macros == OfficePackageContentPolicy.Reject
                && legacy.HasVbaContent) {
                throw new OfficePackageSecurityException(
                    OfficePackageSecurityRule.Macros,
                    "The binary PowerPoint presentation contains a VBA project while macro content is rejected.");
            }
            int embeddedObjectCount = checked(
                (legacy.HasEmbeddedOleContent ? 1 : 0)
                + (legacy.HasLinkedOleContent ? 1 : 0));
            if (security.EmbeddedPayloads ==
                    OfficePackageContentPolicy.Reject
                && embeddedObjectCount > 0) {
                throw new OfficePackageSecurityException(
                    OfficePackageSecurityRule.EmbeddedPayloads,
                    $"The binary PowerPoint presentation contains {embeddedObjectCount} embedded or cached OLE payload(s) while embedded content is rejected.",
                    embeddedObjectCount, 0);
            }
            if (security.ActiveX == OfficePackageContentPolicy.Reject
                && legacy.HasActiveXContent) {
                throw new OfficePackageSecurityException(
                    OfficePackageSecurityRule.ActiveX,
                    "The binary PowerPoint presentation contains ActiveX content while ActiveX content is rejected.",
                    1, 0);
            }
            int externalTargetCount =
                (legacy.HasExternalHyperlinkContent ? 1 : 0)
                + (legacy.HasLinkedOleContent ? 1 : 0)
                + (legacy.HasExternalMediaContent ? 1 : 0)
                + (legacy.HasRunProgramContent ? 1 : 0);
            if (security.ExternalRelationships ==
                    OfficePackageContentPolicy.Reject
                && externalTargetCount > 0) {
                throw new OfficePackageSecurityException(
                    OfficePackageSecurityRule.ExternalRelationships,
                    $"The binary PowerPoint presentation contains {externalTargetCount} external target(s) while external relationships are rejected.",
                    externalTargetCount, 0);
            }
        }

        internal static bool IsExternalLegacyHyperlink(
            LegacyPpt.Model.LegacyPptHyperlink hyperlink) =>
            !hyperlink.IsInternalSlideTarget
            && (!string.IsNullOrWhiteSpace(hyperlink.Target)
                || !string.IsNullOrWhiteSpace(hyperlink.Location));

    }
}
