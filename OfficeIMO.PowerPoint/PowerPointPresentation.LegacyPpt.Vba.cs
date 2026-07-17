using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointPresentation {
        private static void ProjectLegacyVbaProject(
            PowerPointPresentation presentation,
            LegacyPpt.LegacyPptPresentation legacy) {
            if (legacy.VbaProject == null) return;
            PresentationPart presentationPart = presentation.OpenXmlDocument
                .PresentationPart
                ?? throw new InvalidDataException(
                    "The projected presentation has no presentation part.");
            VbaProjectPart vbaPart = presentationPart.AddNewPart<VbaProjectPart>();
            using (var stream = new MemoryStream(legacy.VbaProject.GetBytes(),
                       writable: false)) {
                vbaPart.FeedData(stream);
            }
        }

        internal static byte[] ConvertProjectedVbaPackageToMacroEnabled(
            byte[] packageBytes, PowerPointLoadOptions loadOptions) {
            ValidatePackageSecurity(packageBytes, loadOptions);
            using var stream = new MemoryStream();
            stream.Write(packageBytes, 0, packageBytes.Length);
            stream.Position = 0;
            using (PresentationDocument document = PresentationDocument.Open(
                       stream, true)) {
                document.ChangeDocumentType(
                    PresentationDocumentType.MacroEnabledPresentation);
                document.Save();
            }
            return stream.ToArray();
        }
    }
}
