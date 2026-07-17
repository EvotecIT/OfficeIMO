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

            bool hasVba = legacy.VbaProject != null
                || HasLegacyDiagnostic(legacy, "PPT-VBA-");
            if (security.Macros == OfficePackageContentPolicy.Reject
                && hasVba) {
                throw new OfficePackageSecurityException(
                    OfficePackageSecurityRule.Macros,
                    "The binary PowerPoint presentation contains a VBA project while macro content is rejected.");
            }
            int embeddedObjectCount = checked(legacy.OleObjects.Count
                + legacy.LinkedOleObjects.Count);
            bool hasEmbeddedObjects = embeddedObjectCount > 0
                || HasLegacyDiagnostic(legacy, "PPT-OLE-");
            if (security.EmbeddedPayloads ==
                    OfficePackageContentPolicy.Reject
                && hasEmbeddedObjects) {
                throw new OfficePackageSecurityException(
                    OfficePackageSecurityRule.EmbeddedPayloads,
                    $"The binary PowerPoint presentation contains {embeddedObjectCount} embedded or cached OLE payload(s) while embedded content is rejected.",
                    Math.Max(1, embeddedObjectCount), 0);
            }
            bool hasActiveX = legacy.ActiveXControls.Count > 0
                || HasLegacyDiagnostic(legacy, "PPT-ACTIVEX-");
            if (security.ActiveX == OfficePackageContentPolicy.Reject
                && hasActiveX) {
                throw new OfficePackageSecurityException(
                    OfficePackageSecurityRule.ActiveX,
                    $"The binary PowerPoint presentation contains {legacy.ActiveXControls.Count} ActiveX control(s) while ActiveX content is rejected.",
                    Math.Max(1, legacy.ActiveXControls.Count), 0);
            }
            int externalTargetCount = legacy.Hyperlinks.Count(
                    IsExternalLegacyHyperlink)
                + legacy.LinkedOleObjects.Count
                + legacy.Media.Count(media =>
                    !string.IsNullOrWhiteSpace(media.Path))
                + CountRunProgramInteractions(legacy)
                + legacy.Diagnostics.Count(diagnostic =>
                    IsPreserveOnlyExternalDiagnostic(diagnostic.Code));
            if (security.ExternalRelationships ==
                    OfficePackageContentPolicy.Reject
                && externalTargetCount > 0) {
                throw new OfficePackageSecurityException(
                    OfficePackageSecurityRule.ExternalRelationships,
                    $"The binary PowerPoint presentation contains {externalTargetCount} external target(s) while external relationships are rejected.",
                    externalTargetCount, 0);
            }
        }

        private static bool HasLegacyDiagnostic(
            LegacyPptPresentation legacy, string prefix) =>
            legacy.Diagnostics.Any(diagnostic => diagnostic.Code.StartsWith(
                prefix, StringComparison.Ordinal));

        internal static bool IsExternalLegacyHyperlink(
            LegacyPpt.Model.LegacyPptHyperlink hyperlink) =>
            !hyperlink.IsInternalSlideTarget
            && (!string.IsNullOrWhiteSpace(hyperlink.Target)
                || !string.IsNullOrWhiteSpace(hyperlink.Location));

        private static bool IsPreserveOnlyExternalDiagnostic(string code) =>
            code.StartsWith("PPT-OLE-LINK-", StringComparison.Ordinal)
            || code.StartsWith("PPT-HYPERLINK-", StringComparison.Ordinal)
            || code.StartsWith("PPT-HYPERLINK9-", StringComparison.Ordinal)
            || code.StartsWith("PPT-MEDIA-MALFORMED", StringComparison.Ordinal)
            || code.StartsWith("PPT-MEDIA-UNKNOWN", StringComparison.Ordinal);

        private static int CountRunProgramInteractions(
            LegacyPptPresentation legacy) => EnumerateLegacyShapes(legacy)
            .Sum(shape => shape.Interactions.Count(interaction =>
                    interaction.Action == LegacyPpt.Model
                        .LegacyPptInteractionAction.RunProgram)
                + shape.TextBody.Interactions.Count(textInteraction =>
                    textInteraction.Interaction.Action == LegacyPpt.Model
                        .LegacyPptInteractionAction.RunProgram));

        private static IEnumerable<LegacyPpt.Model.LegacyPptShape>
            EnumerateLegacyShapes(LegacyPptPresentation legacy) {
            foreach (LegacyPpt.Model.LegacyPptShape shape in legacy.Slides
                         .SelectMany(slide => slide.Shapes)) {
                foreach (LegacyPpt.Model.LegacyPptShape nested in
                         EnumerateLegacyShapes(shape)) yield return nested;
            }
            foreach (LegacyPpt.Model.LegacyPptShape shape in legacy.Slides
                         .Where(slide => slide.NotesPage != null)
                         .SelectMany(slide => slide.NotesPage!.Shapes)) {
                foreach (LegacyPpt.Model.LegacyPptShape nested in
                         EnumerateLegacyShapes(shape)) yield return nested;
            }
            foreach (LegacyPpt.Model.LegacyPptShape shape in legacy.Masters
                         .SelectMany(master => master.Shapes)) {
                foreach (LegacyPpt.Model.LegacyPptShape nested in
                         EnumerateLegacyShapes(shape)) yield return nested;
            }
            foreach (LegacyPpt.Model.LegacyPptShape shape in legacy
                         .NotesMaster?.Shapes
                     ?? Array.Empty<LegacyPpt.Model.LegacyPptShape>()) {
                foreach (LegacyPpt.Model.LegacyPptShape nested in
                         EnumerateLegacyShapes(shape)) yield return nested;
            }
            foreach (LegacyPpt.Model.LegacyPptShape shape in legacy
                         .HandoutMaster?.Shapes
                     ?? Array.Empty<LegacyPpt.Model.LegacyPptShape>()) {
                foreach (LegacyPpt.Model.LegacyPptShape nested in
                         EnumerateLegacyShapes(shape)) yield return nested;
            }
        }

        private static IEnumerable<LegacyPpt.Model.LegacyPptShape>
            EnumerateLegacyShapes(LegacyPpt.Model.LegacyPptShape shape) {
            yield return shape;
            foreach (LegacyPpt.Model.LegacyPptShape child in shape.Children) {
                foreach (LegacyPpt.Model.LegacyPptShape nested in
                         EnumerateLegacyShapes(child)) yield return nested;
            }
        }
    }
}
