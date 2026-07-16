using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint.LegacyPpt.Capabilities;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWritePreflight {
        private const string VbaProjectContentType =
            "application/vnd.ms-office.vbaProject";

        private static void AddPackagePartFindings(
            PowerPointPresentation presentation,
            ICollection<LegacyPptWriteFinding> findings) {
            PresentationDocument document = presentation.OpenXmlDocument;
            OpenXmlPart[] parts = EnumeratePackageParts(document).ToArray();

            if (parts.Any(IsCustomXmlPart)) {
                findings.Add(new LegacyPptWriteFinding(
                    LegacyPptFeature.CustomXml,
                    "PPT-WRITE-CUSTOM-XML",
                    "Open XML custom XML parts have no PowerPoint 97-2003 binary representation."));
            }

            if (parts.Any(IsActiveXPart)) {
                findings.Add(new LegacyPptWriteFinding(
                    LegacyPptFeature.ActiveX,
                    "PPT-WRITE-ACTIVEX",
                    "Open XML ActiveX control parts cannot be converted safely to binary PowerPoint controls."));
            }

            if (parts.Any(IsWebExtensionPart)) {
                findings.Add(new LegacyPptWriteFinding(
                    LegacyPptFeature.UnknownRecordsAndStreams,
                    "PPT-WRITE-WEB-EXTENSION",
                    "Web extensions and task panes have no PowerPoint 97-2003 binary representation."));
            }

            if (parts.OfType<ExtendedPart>().Any(IsExtendedVbaProject)) {
                findings.Add(new LegacyPptWriteFinding(
                    LegacyPptFeature.VbaProjects,
                    "PPT-WRITE-VBA-EXTENDED-PART",
                    "A VBA project stored as an untyped package part cannot be encoded safely by the binary writer."));
            }

            bool hasOpenXmlSignature = document.DigitalSignatureOriginPart != null
                || document.ExtendedFilePropertiesPart?.Properties?
                    .DigitalSignature != null
                || parts.Any(IsDigitalSignaturePart);
            if (hasOpenXmlSignature) {
                findings.Add(new LegacyPptWriteFinding(
                    LegacyPptFeature.DigitalSignatures,
                    "PPT-WRITE-OPENXML-SIGNATURE",
                    "Open XML digital signatures cannot be translated into binary PowerPoint signature carriers."));
            }

            AddUnknownPackagePartFindings(parts, findings);
            AddUnmappedEmbeddedPartFinding(presentation, parts, findings);
            AddExternalRelationshipFindings(document, parts, findings);
        }

        private static IEnumerable<OpenXmlPart> EnumeratePackageParts(
            PresentationDocument document) {
            var visited = new HashSet<OpenXmlPart>();
            var pending = new Stack<OpenXmlPart>(
                document.Parts.Select(pair => pair.OpenXmlPart));
            if (document.DigitalSignatureOriginPart != null) {
                pending.Push(document.DigitalSignatureOriginPart);
            }

            while (pending.Count > 0) {
                OpenXmlPart part = pending.Pop();
                if (!visited.Add(part)) continue;
                yield return part;
                foreach (IdPartPair child in part.Parts) {
                    pending.Push(child.OpenXmlPart);
                }
            }
        }

        private static bool IsCustomXmlPart(OpenXmlPart part) =>
            part is CustomXmlPart
            || part.Uri.OriginalString.IndexOf("/customXml/",
                StringComparison.OrdinalIgnoreCase) >= 0;

        private static bool IsActiveXPart(OpenXmlPart part) =>
            part.ContentType.Equals("application/vnd.ms-office.activeX+xml",
                StringComparison.OrdinalIgnoreCase)
            || part.ContentType.Equals("application/vnd.ms-office.activeX",
                StringComparison.OrdinalIgnoreCase)
            || part.ContentType.Equals("application/vnd.ms-office.activeX.bin",
                StringComparison.OrdinalIgnoreCase);

        private static bool IsWebExtensionPart(OpenXmlPart part) =>
            part.ContentType.Equals(
                "application/vnd.ms-office.webextension+xml",
                StringComparison.OrdinalIgnoreCase)
            || part.ContentType.Equals(
                "application/vnd.ms-office.webextensiontaskpanes+xml",
                StringComparison.OrdinalIgnoreCase);

        private static bool IsExtendedVbaProject(ExtendedPart part) =>
            part.ContentType.Equals(VbaProjectContentType,
                StringComparison.OrdinalIgnoreCase);

        private static bool IsDigitalSignaturePart(OpenXmlPart part) =>
            part.Uri.OriginalString.IndexOf("/_xmlsignatures/",
                StringComparison.OrdinalIgnoreCase) >= 0
            || part.ContentType.IndexOf("digital-signature",
                StringComparison.OrdinalIgnoreCase) >= 0
            || part.ContentType.IndexOf("xmlsignature",
                StringComparison.OrdinalIgnoreCase) >= 0;

        private static void AddUnknownPackagePartFindings(
            IReadOnlyCollection<OpenXmlPart> parts,
            ICollection<LegacyPptWriteFinding> findings) {
            ExtendedPart[] unknown = parts.OfType<ExtendedPart>()
                .Where(part => !IsActiveXPart(part)
                    && !IsWebExtensionPart(part)
                    && !IsExtendedVbaProject(part)
                    && !IsDigitalSignaturePart(part))
                .ToArray();
            if (unknown.Length > 0) {
                findings.Add(new LegacyPptWriteFinding(
                    LegacyPptFeature.UnknownRecordsAndStreams,
                    "PPT-WRITE-EXTENDED-PART",
                    $"The presentation contains {unknown.Length} unrecognized Open XML extended package part(s) that the binary writer cannot preserve."));
            }

            OpenXmlPart[] typed = parts
                .Where(part => part is not ExtendedPart
                    && !IsKnownHandledPart(part)
                    && !IsCustomXmlPart(part)
                    && !IsActiveXPart(part)
                    && !IsWebExtensionPart(part)
                    && !IsDigitalSignaturePart(part)
                    && part is not DigitalSignatureOriginPart
                    && part is not XmlSignaturePart)
                .ToArray();
            if (typed.Length == 0) return;

            findings.Add(new LegacyPptWriteFinding(
                LegacyPptFeature.UnknownRecordsAndStreams,
                "PPT-WRITE-PACKAGE-PART",
                $"The presentation contains {typed.Length} typed Open XML package part(s) outside the binary writer's consumed or explicitly converted contract."));
        }

        private static bool IsKnownHandledPart(OpenXmlPart part) =>
            part is PresentationPart
            or PresentationPropertiesPart
            or ViewPropertiesPart
            or TableStylesPart
            or CoreFilePropertiesPart
            or ExtendedFilePropertiesPart
            or CustomFilePropertiesPart
            or ThumbnailPart
            or SlidePart
            or SlideMasterPart
            or SlideLayoutPart
            or NotesSlidePart
            or NotesMasterPart
            or HandoutMasterPart
            or ThemePart
            or ThemeOverridePart
            or ImagePart
            or ChartPart
            or ChartStylePart
            or ChartColorStylePart
            or DiagramDataPart
            or DiagramLayoutDefinitionPart
            or DiagramColorsPart
            or DiagramStylePart
            or DiagramPersistLayoutPart
            or CommentAuthorsPart
            or SlideCommentsPart
            or PowerPointAuthorsPart
            or PowerPointCommentPart
            or VbaProjectPart
            or EmbeddedObjectPart
            or EmbeddedPackagePart;

        private static void AddUnmappedEmbeddedPartFinding(
            PowerPointPresentation presentation,
            IReadOnlyCollection<OpenXmlPart> parts,
            ICollection<LegacyPptWriteFinding> findings) {
            var mapped = new HashSet<OpenXmlPart>(presentation.Slides
                .SelectMany(slide => slide.EnumerateShapesDeep(
                    slide.Shapes, includeHidden: true))
                .OfType<PowerPointOleObject>()
                .Select(ole => (OpenXmlPart)ole.EmbeddedPart));
            foreach (ChartPart chartPart in parts.OfType<ChartPart>()) {
                foreach (EmbeddedPackagePart workbook in chartPart
                    .GetPartsOfType<EmbeddedPackagePart>()) {
                    mapped.Add(workbook);
                }
            }

            int count = parts.Count(part =>
                part is EmbeddedPackagePart or EmbeddedObjectPart
                && !mapped.Contains(part));
            if (count == 0) return;

            findings.Add(new LegacyPptWriteFinding(
                LegacyPptFeature.UnknownRecordsAndStreams,
                "PPT-WRITE-EMBEDDED-PACKAGE",
                $"The presentation contains {count} embedded package or object part(s) that are not mapped to an editable OLE object or converted chart."));
        }

        private static void AddExternalRelationshipFindings(
            PresentationDocument document,
            IReadOnlyCollection<OpenXmlPart> parts,
            ICollection<LegacyPptWriteFinding> findings) {
            ExternalRelationship[] relationships = document
                .ExternalRelationships
                .Concat(parts.SelectMany(part => part.ExternalRelationships))
                .ToArray();
            if (relationships.Length == 0) return;

            int linkedOle = relationships.Count(relationship =>
                relationship.RelationshipType.IndexOf("oleObject",
                    StringComparison.OrdinalIgnoreCase) >= 0);
            int linkedMedia = relationships.Count(relationship =>
                relationship.RelationshipType.IndexOf("audio",
                    StringComparison.OrdinalIgnoreCase) >= 0
                || relationship.RelationshipType.IndexOf("video",
                    StringComparison.OrdinalIgnoreCase) >= 0
                || relationship.RelationshipType.IndexOf("media",
                    StringComparison.OrdinalIgnoreCase) >= 0);
            int other = relationships.Length - linkedOle - linkedMedia;

            if (linkedOle > 0) {
                findings.Add(new LegacyPptWriteFinding(
                    LegacyPptFeature.LinkedOle,
                    "PPT-WRITE-LINKED-OLE",
                    $"The presentation contains {linkedOle} external OLE relationship(s) whose link and cache semantics cannot be converted safely."));
            }
            if (linkedMedia > 0) {
                findings.Add(new LegacyPptWriteFinding(
                    LegacyPptFeature.LinkedMedia,
                    "PPT-WRITE-LINKED-MEDIA",
                    $"The presentation contains {linkedMedia} external media relationship(s) whose path and playback semantics cannot be converted safely."));
            }
            if (other > 0) {
                findings.Add(new LegacyPptWriteFinding(
                    LegacyPptFeature.UnknownRecordsAndStreams,
                    "PPT-WRITE-EXTERNAL-RELATIONSHIP",
                    $"The presentation contains {other} external package relationship(s) outside supported hyperlink actions."));
            }
        }
    }
}
