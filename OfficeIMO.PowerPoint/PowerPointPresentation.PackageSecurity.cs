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
            if (security.EmbeddedPayloads ==
                    OfficePackageContentPolicy.Reject
                && (legacy.HasEmbeddedOleContent
                    || legacy.HasLinkedOleContent)) {
                throw new OfficePackageSecurityException(
                    OfficePackageSecurityRule.EmbeddedPayloads,
                    "The binary PowerPoint presentation contains embedded or cached OLE content while embedded payloads are rejected.");
            }
            if (security.ActiveX == OfficePackageContentPolicy.Reject
                && legacy.HasActiveXContent) {
                throw new OfficePackageSecurityException(
                    OfficePackageSecurityRule.ActiveX,
                    "The binary PowerPoint presentation contains ActiveX content while ActiveX content is rejected.");
            }
            if (security.ExternalRelationships ==
                    OfficePackageContentPolicy.Reject
                && (legacy.HasExternalHyperlinkContent
                    || legacy.HasLinkedOleContent
                    || legacy.HasExternalMediaContent
                    || legacy.HasRunProgramContent)) {
                throw new OfficePackageSecurityException(
                    OfficePackageSecurityRule.ExternalRelationships,
                    "The binary PowerPoint presentation contains external content while external relationships are rejected.");
            }
        }

        internal static bool IsExternalLegacyHyperlink(
            LegacyPpt.Model.LegacyPptHyperlink hyperlink) =>
            !hyperlink.IsInternalSlideTarget
            && (!string.IsNullOrWhiteSpace(hyperlink.Target)
                || !string.IsNullOrWhiteSpace(hyperlink.Location));

    }
}
