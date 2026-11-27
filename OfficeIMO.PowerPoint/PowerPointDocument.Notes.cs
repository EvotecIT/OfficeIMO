using System.Text;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        /// <summary>
        /// A slide's notes slide part, created as the notes model first uses it when the slide has none: on the package
        /// while it holds the slide and the presentation has a notes master, through the SDK otherwise.
        /// </summary>
        internal PowerPointPartRef EnsureNotesSlidePart(PowerPointSlide slide) {
            if (!IsSdkLoaded && slide is { IsPackageBacked: true, RemovedFromPackage: false, PackagePartName: { } slidePart }
                && (GetPackageNotes(slidePart) ?? TryCreatePackageNotes(slidePart)) != null) {
                PowerPointPartRef notes = slide.PartRef.RelatedPart(PowerPointPartRef.NotesSlideRelationshipType);
                if (!notes.IsNone) return notes;
            }
            return PowerPointPartRef.FromSdk(PowerPointUtils.EnsureNotesSlidePart(slide.SdkSlidePart));
        }

        private const string NotesMasterContentType = "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml";
        private const string ThemeContentType = "application/vnd.openxmlformats-officedocument.theme+xml";

        // The notes master the model adds: a body, slide image, date, slide number and footer placeholder, the standard
        // color map and en-US notes styles. Built when needed: a static field here would run in the class's static
        // constructor, which the SDK-bound members share.
        private static byte[] CreateDefaultNotesMaster() => Encoding.UTF8.GetBytes(
            "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
            + "<p:notesMaster xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">"
            + "<p:cSld><p:bg><p:bgPr><a:noFill /></p:bgPr></p:bg><p:spTree>"
            + "<p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"Notes Group Shape\" /><p:cNvGrpSpPr /><p:nvPr /></p:nvGrpSpPr>"
            + "<p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\" /><a:ext cx=\"0\" cy=\"0\" /><a:chOff x=\"0\" y=\"0\" /><a:chExt cx=\"0\" cy=\"0\" /></a:xfrm></p:grpSpPr>"
            + NotesMasterPlaceholder(2, "Notes Placeholder", "body", 1, true)
            + NotesMasterPlaceholder(3, "Slide Image Placeholder", "sldImg", 2, false)
            + NotesMasterPlaceholder(4, "Date Placeholder", "dt", 3, true)
            + NotesMasterPlaceholder(5, "Slide Number Placeholder", "sldNum", 4, true)
            + NotesMasterPlaceholder(6, "Footer Placeholder", "ftr", 5, true)
            + "</p:spTree></p:cSld>"
            + "<p:clrMap bg1=\"lt1\" tx1=\"dk1\" bg2=\"lt2\" tx2=\"dk2\" accent1=\"accent1\" accent2=\"accent2\" accent3=\"accent3\" accent4=\"accent4\" "
            + "accent5=\"accent5\" accent6=\"accent6\" hlink=\"hlink\" folHlink=\"folHlink\" />"
            + "<p:notesStyle>" + string.Concat(Enumerable.Range(1, 9).Select(level => "<a:lvl" + level + "pPr><a:defRPr lang=\"en-US\" /></a:lvl" + level + "pPr>")) + "</p:notesStyle>"
            + "</p:notesMaster>");

        private static string NotesMasterPlaceholder(int id, string name, string type, int index, bool endProperties) =>
            "<p:sp><p:nvSpPr><p:cNvPr id=\"" + id + "\" name=\"" + name + "\" /><p:cNvSpPr /><p:nvPr><p:ph type=\"" + type + "\" idx=\"" + index + "\" /></p:nvPr></p:nvSpPr>"
            + "<p:spPr /><p:txBody><a:bodyPr /><a:lstStyle />" + (endProperties ? "<a:p><a:endParaRPr lang=\"en-US\" /></a:p>" : "<a:p />") + "</p:txBody></p:sp>";

        // Adds the default notes master with a copy of the presentation's theme, relates it to the presentation and lists
        // it there, as the model does when it first needs one. False, having changed nothing, when the presentation has
        // no theme to copy.
        private bool TryAddPackageNotesMaster(OfficeOpcEditablePackage package, PackageSlideList slides, out string? master, out string? relationship) {
            master = relationship = null;
            const string prefix = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix;
            if (FirstRelated(package, slides.Main, prefix + "theme", out _) is not { } theme) return false;
            string masterPart = NextPartName(package, NotesMasterContentType, "/ppt/notesMasters/notesMaster", ".xml");
            package.CreateBinaryPart(masterPart, NotesMasterContentType, CreateDefaultNotesMaster());
            string themePart = NextPartName(package, ThemeContentType, "/ppt/notesMasters/theme/theme", ".xml");
            package.CreateBinaryPart(themePart, ThemeContentType, package.ReadPartBytes(theme));
            package.AddRelationship(masterPart, prefix + "theme", OfficeOpcPartNames.GetRelativeTarget(masterPart, themePart), id: NewRelationshipId());
            relationship = package.AddRelationship(slides.Main, prefix + "notesMaster", OfficeOpcPartNames.GetRelativeTarget(slides.Main, masterPart), id: NewRelationshipId());

            OfficeXmlNode presentation = new OfficeXmlNode(slides.Presentation, slides.Presentation.Root);
            OfficeXmlNode list = presentation.GetOrAddChild(OfficeXmlNamespaces.PresentationML, "notesMasterIdLst", PPresentationView.Order);
            list.Append(list.CreateElement(OfficeXmlNamespaces.PresentationML, "notesMasterId")).SetAttribute(OfficeXmlNamespaces.Relationships, "id", relationship);
            master = masterPart;
            return true;
        }

        // A relationship id as the SDK makes them: R and sixteen hex digits.
        private static string NewRelationshipId() => "R" + Guid.NewGuid().ToString("N").Substring(0, 16);
    }
}
