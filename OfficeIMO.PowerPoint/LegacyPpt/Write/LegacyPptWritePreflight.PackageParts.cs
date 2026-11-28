using OfficeIMO.PowerPoint.LegacyPpt.Capabilities;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWritePreflight {
        private const string VbaProjectContentType =
            "application/vnd.ms-office.vbaProject";
        private const string RelationshipPrefix = LegacyPptPackageView.RelationshipTypePrefix;
        private const string OfficeRelationshipPrefix = "http://schemas.microsoft.com/office/";
        private const string DigitalSignatureOriginType =
            "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin";
        private const string DigitalSignatureType =
            "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature";
        private const string ThumbnailType =
            "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail";

        // Relationship types of the parts the binary writer consumes or converts.
        private static readonly HashSet<string> HandledRelationshipTypes = new(StringComparer.Ordinal) {
            LegacyPptPackageView.OfficeDocumentRelationshipType,
            RelationshipPrefix + "presProps",
            RelationshipPrefix + "viewProps",
            RelationshipPrefix + "tableStyles",
            LegacyPptPackageView.CorePropertiesRelationshipType,
            LegacyPptPackageView.ExtendedPropertiesRelationshipType,
            LegacyPptPackageView.CustomPropertiesRelationshipType,
            ThumbnailType,
            LegacyPptPackageView.SlideRelationshipType,
            LegacyPptPackageView.SlideMasterRelationshipType,
            LegacyPptPackageView.SlideLayoutRelationshipType,
            LegacyPptPackageView.NotesSlideRelationshipType,
            LegacyPptPackageView.NotesMasterRelationshipType,
            LegacyPptPackageView.HandoutMasterRelationshipType,
            LegacyPptPackageView.ThemeRelationshipType,
            LegacyPptPackageView.ThemeOverrideRelationshipType,
            LegacyPptPackageView.ImageRelationshipType,
            LegacyPptPackageView.ChartRelationshipType,
            OfficeRelationshipPrefix + "2011/relationships/chartStyle",
            OfficeRelationshipPrefix + "2011/relationships/chartColorStyle",
            RelationshipPrefix + "diagramData",
            RelationshipPrefix + "diagramLayout",
            RelationshipPrefix + "diagramColors",
            RelationshipPrefix + "diagramQuickStyle",
            OfficeRelationshipPrefix + "2007/relationships/diagramDrawing",
            LegacyPptPackageView.CommentAuthorsRelationshipType,
            LegacyPptPackageView.CommentsRelationshipType,
            OfficeRelationshipPrefix + "2018/10/relationships/authors",
            OfficeRelationshipPrefix + "2018/10/relationships/comments",
            LegacyPptPackageView.VbaProjectRelationshipType,
            LegacyPptPackageView.OleObjectRelationshipType,
            LegacyPptPackageView.PackageRelationshipType
        };

        // Further relationship types that load as typed parts in a presentation package, which the binary writer has no
        // contract for.
        private static readonly HashSet<string> OtherTypedRelationshipTypes = new(StringComparer.Ordinal) {
            RelationshipPrefix + "customXml",
            RelationshipPrefix + "customXmlProps",
            RelationshipPrefix + "font",
            RelationshipPrefix + "tags",
            RelationshipPrefix + "vmlDrawing",
            RelationshipPrefix + "control",
            RelationshipPrefix + "slideUpdateInfo",
            RelationshipPrefix + "chartUserShapes",
            RelationshipPrefix + "printerSettings",
            RelationshipPrefix + "customProperty",
            DigitalSignatureOriginType,
            DigitalSignatureType,
            OfficeRelationshipPrefix + "2006/relationships/ui/extensibility",
            OfficeRelationshipPrefix + "2007/relationships/ui/extensibility",
            OfficeRelationshipPrefix + "2006/relationships/ui/userCustomization",
            OfficeRelationshipPrefix + "2006/relationships/legacyDocTextInfo",
            OfficeRelationshipPrefix + "2006/relationships/activeXControlBinary",
            OfficeRelationshipPrefix + "2007/relationships/hdphoto",
            OfficeRelationshipPrefix + "2011/relationships/webextensiontaskpanes",
            OfficeRelationshipPrefix + "2011/relationships/webextension",
            OfficeRelationshipPrefix + "2017/06/relationships/model3d",
            OfficeRelationshipPrefix + "2020/02/relationships/classificationlabels"
        };

        /// <summary>A part reached through the package relationships, with the relationship type that first reached it.</summary>
        private readonly struct PackagePart {
            internal PackagePart(LegacyPptPartView part, string relationshipType) {
                Part = part;
                RelationshipType = relationshipType;
            }

            internal LegacyPptPartView Part { get; }
            internal string RelationshipType { get; }
            internal string Name => Part.Name;
            internal string ContentType => Part.ContentType ?? string.Empty;

            internal bool HasType(string relationshipType) =>
                string.Equals(RelationshipType, relationshipType, StringComparison.Ordinal);

            /// <summary>
            /// An untyped extension part: one whose relationship type no typed part claims, or one stored where untyped
            /// parts are added (<c>udata/data…</c>) whatever relationship type it was given.
            /// </summary>
            internal bool IsExtended =>
                !HandledRelationshipTypes.Contains(RelationshipType)
                && !OtherTypedRelationshipTypes.Contains(RelationshipType)
                || Name.IndexOf("/udata/data", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static void AddPackagePartFindings(
            LegacyPptPackageView view, PowerPointDocument presentation,
            ICollection<LegacyPptWriteFinding> findings) {
            PackagePart[] parts = EnumeratePackageParts(view).ToArray();

            if (parts.Any(IsCustomXmlPart)) {
                findings.Add(new LegacyPptWriteFinding(
                    LegacyPptFeature.CustomXml,
                    "ppt-write-custom-xml",
                    "Open XML custom XML parts have no PowerPoint 97-2003 binary representation."));
            }

            if (parts.Any(IsActiveXPart)) {
                findings.Add(new LegacyPptWriteFinding(
                    LegacyPptFeature.ActiveX,
                    "ppt-write-activex",
                    "Open XML ActiveX control parts cannot be converted safely to binary PowerPoint controls."));
            }

            if (parts.Any(IsWebExtensionPart)) {
                findings.Add(new LegacyPptWriteFinding(
                    LegacyPptFeature.UnknownRecordsAndStreams,
                    "ppt-write-web-extension",
                    "Web extensions and task panes have no PowerPoint 97-2003 binary representation."));
            }

            if (parts.Any(IsExtendedVbaProject)) {
                findings.Add(new LegacyPptWriteFinding(
                    LegacyPptFeature.VbaProjects,
                    "ppt-write-vba-extended-part",
                    "A VBA project stored as an untyped package part cannot be encoded safely by the binary writer."));
            }

            bool hasOpenXmlSignature = !view.PackageRelatedPart(DigitalSignatureOriginType).IsNone
                || !view.PackageRelatedPart(LegacyPptPackageView.ExtendedPropertiesRelationshipType)
                    .Root.Child(OfficeIMO.Core.Internal.OpenXml.OfficeXmlNamespaces.ExtendedProperties, "DigSig").IsNone
                || parts.Any(IsDigitalSignaturePart);
            if (hasOpenXmlSignature) {
                findings.Add(new LegacyPptWriteFinding(
                    LegacyPptFeature.DigitalSignatures,
                    "ppt-write-openxml-signature",
                    "Open XML digital signatures cannot be translated into binary PowerPoint signature carriers."));
            }

            AddUnknownPackagePartFindings(parts, findings);
            AddUnmappedEmbeddedPartFinding(view, presentation, parts, findings);
            AddExternalRelationshipFindings(view, parts, findings);
        }

        /// <summary>
        /// The parts reachable from the package relationships through part relationships: internal relationships other
        /// than hyperlinks and audio, video or media references, each part once.
        /// </summary>
        private static IEnumerable<PackagePart> EnumeratePackageParts(
            LegacyPptPackageView view) {
            var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var pending = new Stack<PackagePart>();
            foreach (LegacyPptRelationshipView relationship in view.PackageRelationships) {
                if (IsPartRelationship(relationship)) {
                    pending.Push(new PackagePart(relationship.TargetPart, NormalizedType(relationship)));
                }
            }

            while (pending.Count > 0) {
                PackagePart part = pending.Pop();
                if (!visited.Add(part.Name)) continue;
                yield return part;
                foreach (LegacyPptRelationshipView relationship in part.Part.Relationships) {
                    if (IsPartRelationship(relationship)) {
                        pending.Push(new PackagePart(relationship.TargetPart, NormalizedType(relationship)));
                    }
                }
            }
        }

        private static bool IsPartRelationship(LegacyPptRelationshipView relationship) =>
            !relationship.IsExternal
            && !relationship.TargetPart.IsNone
            && !relationship.IsType(LegacyPptPackageView.HyperlinkRelationshipType)
            && !relationship.IsType(LegacyPptPackageView.AudioRelationshipType)
            && !relationship.IsType(LegacyPptPackageView.VideoRelationshipType)
            && !relationship.IsType(LegacyPptPackageView.MediaRelationshipType);

        private static string NormalizedType(LegacyPptRelationshipView relationship) =>
            OfficeIMO.Core.Internal.OpenXml.OfficeXmlNamespaces.NormalizeRelationshipType(relationship.Type);

        private static bool IsCustomXmlPart(PackagePart part) =>
            part.HasType(RelationshipPrefix + "customXml")
            || part.Name.IndexOf("/customXml/",
                StringComparison.OrdinalIgnoreCase) >= 0;

        private static bool IsActiveXPart(PackagePart part) =>
            part.ContentType.Equals("application/vnd.ms-office.activeX+xml",
                StringComparison.OrdinalIgnoreCase)
            || part.ContentType.Equals("application/vnd.ms-office.activeX",
                StringComparison.OrdinalIgnoreCase)
            || part.ContentType.Equals("application/vnd.ms-office.activeX.bin",
                StringComparison.OrdinalIgnoreCase);

        private static bool IsWebExtensionPart(PackagePart part) =>
            part.ContentType.Equals(
                "application/vnd.ms-office.webextension+xml",
                StringComparison.OrdinalIgnoreCase)
            || part.ContentType.Equals(
                "application/vnd.ms-office.webextensiontaskpanes+xml",
                StringComparison.OrdinalIgnoreCase);

        private static bool IsExtendedVbaProject(PackagePart part) =>
            part.IsExtended
            && part.ContentType.Equals(VbaProjectContentType,
                StringComparison.OrdinalIgnoreCase);

        private static bool IsDigitalSignaturePart(PackagePart part) =>
            part.Name.IndexOf("/_xmlsignatures/",
                StringComparison.OrdinalIgnoreCase) >= 0
            || part.ContentType.IndexOf("digital-signature",
                StringComparison.OrdinalIgnoreCase) >= 0
            || part.ContentType.IndexOf("xmlsignature",
                StringComparison.OrdinalIgnoreCase) >= 0;

        private static void AddUnknownPackagePartFindings(
            IReadOnlyCollection<PackagePart> parts,
            ICollection<LegacyPptWriteFinding> findings) {
            PackagePart[] unknown = parts
                .Where(part => part.IsExtended
                    && !IsActiveXPart(part)
                    && !IsWebExtensionPart(part)
                    && !IsExtendedVbaProject(part)
                    && !IsDigitalSignaturePart(part))
                .ToArray();
            if (unknown.Length > 0) {
                findings.Add(new LegacyPptWriteFinding(
                    LegacyPptFeature.UnknownRecordsAndStreams,
                    "ppt-write-extended-part",
                    $"The presentation contains {unknown.Length} unrecognized Open XML extended package part(s) that the binary writer cannot preserve."));
            }

            PackagePart[] typed = parts
                .Where(part => !part.IsExtended
                    && !HandledRelationshipTypes.Contains(part.RelationshipType)
                    && !IsCustomXmlPart(part)
                    && !IsActiveXPart(part)
                    && !IsWebExtensionPart(part)
                    && !IsDigitalSignaturePart(part)
                    && !part.HasType(DigitalSignatureOriginType)
                    && !part.HasType(DigitalSignatureType))
                .ToArray();
            if (typed.Length == 0) return;

            findings.Add(new LegacyPptWriteFinding(
                LegacyPptFeature.UnknownRecordsAndStreams,
                "ppt-write-package-part",
                $"The presentation contains {typed.Length} typed Open XML package part(s) outside the binary writer's consumed or explicitly converted contract."));
        }

        private static void AddUnmappedEmbeddedPartFinding(
            LegacyPptPackageView view, PowerPointDocument presentation,
            IReadOnlyCollection<PackagePart> parts,
            ICollection<LegacyPptWriteFinding> findings) {
            var mapped = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (PowerPointSlide slide in presentation.Slides) {
                LegacyPptPartView slidePart = view.Part(slide.PartRef.Name);
                foreach (PowerPointOleObject ole in slide.EnumerateShapesDeep(
                             slide.Shapes, includeHidden: true)
                             .OfType<PowerPointOleObject>()) {
                    if (!ole.IsEmbedded) continue;
                    OfficeXmlNode frame = view.Resolve(ole.OwnerPartRef.Name, ole.XmlElement);
                    string? relationshipId = OleObjectElement(frame).Attribute(LegacyPptXml.R, "id");
                    LegacyPptPartView embedded = slidePart.PartById(relationshipId);
                    // An unresolved relationship maps nothing here; the writer reports the OLE object itself.
                    if (embedded.IsNone) continue;
                    mapped.Add(embedded.Name);
                }
            }
            foreach (PackagePart chartPart in parts.Where(part =>
                         part.HasType(LegacyPptPackageView.ChartRelationshipType))) {
                foreach (LegacyPptRelationshipView relationship in chartPart.Part.Relationships) {
                    if (IsPartRelationship(relationship)
                        && relationship.IsType(LegacyPptPackageView.PackageRelationshipType)) {
                        mapped.Add(relationship.TargetPart.Name);
                    }
                }
            }

            int count = parts.Count(part =>
                (part.HasType(LegacyPptPackageView.PackageRelationshipType)
                    || part.HasType(LegacyPptPackageView.OleObjectRelationshipType))
                && !mapped.Contains(part.Name));
            if (count == 0) return;

            findings.Add(new LegacyPptWriteFinding(
                LegacyPptFeature.UnknownRecordsAndStreams,
                "ppt-write-embedded-package",
                $"The presentation contains {count} embedded package or object part(s) that are not mapped to an editable OLE object or converted chart."));
        }

        private static OfficeXmlNode OleObjectElement(OfficeXmlNode frame) =>
            frame.Is(LegacyPptXml.P, "graphicFrame")
                ? frame.Child(LegacyPptXml.A, "graphic")
                    .Child(LegacyPptXml.A, "graphicData")
                    .Child(LegacyPptXml.P, "oleObj")
                : default;

        private static void AddExternalRelationshipFindings(
            LegacyPptPackageView view,
            IReadOnlyCollection<PackagePart> parts,
            ICollection<LegacyPptWriteFinding> findings) {
            LegacyPptRelationshipView[] relationships = view.PackageRelationships
                .Concat(parts.SelectMany(part => part.Part.Relationships))
                .Where(relationship => relationship.IsExternal
                    && !relationship.IsType(LegacyPptPackageView.HyperlinkRelationshipType))
                .ToArray();
            if (relationships.Length == 0) return;

            int linkedOle = relationships.Count(relationship =>
                relationship.Type.IndexOf("oleObject",
                    StringComparison.OrdinalIgnoreCase) >= 0);
            int linkedMedia = relationships.Count(relationship =>
                relationship.Type.IndexOf("audio",
                    StringComparison.OrdinalIgnoreCase) >= 0
                || relationship.Type.IndexOf("video",
                    StringComparison.OrdinalIgnoreCase) >= 0
                || relationship.Type.IndexOf("media",
                    StringComparison.OrdinalIgnoreCase) >= 0);
            int other = relationships.Length - linkedOle - linkedMedia;

            if (linkedOle > 0) {
                findings.Add(new LegacyPptWriteFinding(
                    LegacyPptFeature.LinkedOle,
                    "ppt-write-linked-ole",
                    $"The presentation contains {linkedOle} external OLE relationship(s) whose link and cache semantics cannot be converted safely."));
            }
            if (linkedMedia > 0) {
                findings.Add(new LegacyPptWriteFinding(
                    LegacyPptFeature.LinkedMedia,
                    "ppt-write-linked-media",
                    $"The presentation contains {linkedMedia} external media relationship(s) whose path and playback semantics cannot be converted safely."));
            }
            if (other > 0) {
                findings.Add(new LegacyPptWriteFinding(
                    LegacyPptFeature.UnknownRecordsAndStreams,
                    "ppt-write-external-relationship",
                    $"The presentation contains {other} external package relationship(s) outside supported hyperlink actions."));
            }
        }
    }
}
