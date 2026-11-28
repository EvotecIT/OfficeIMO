using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        /// <summary>Creates a native section boundary that the next appended slide will occupy.</summary>
        internal PowerPointSectionInfo StartSectionAtNextSlide(string name) {
            ThrowIfDisposed();
            ArgumentException.ThrowIfNullOrWhiteSpace(name);
            PowerPointPartRef presentation = PresentationPartRef.ForMovingXmlNodes();
            OfficeXmlNode root = presentation.Root;
            List<OfficeXmlNode> slides = SlidesForSections(root);
            OfficeXmlNode list = EnsureSectionList(root, slides);
            OfficeXmlNode section = list.Append(CreateSection(list, name, Array.Empty<uint>()));
            presentation.MarkEdited();
            return BuildSectionInfo(section, BuildSlideIndexMap(slides));
        }

        /// <summary>Restores only section metadata when a collected slide or column is rejected.</summary>
        internal Action CaptureSectionRollback() {
            PowerPointPartRef presentation = PresentationPartRef.ForMovingXmlNodes();
            OfficeXmlNode capturedRoot = presentation.Root;
            if (capturedRoot.IsNone) throw new InvalidOperationException("Presentation is missing.");
            OfficeXmlNode capturedList = capturedRoot.Child(SectionPresentationNamespace, "extLst");
            OfficeXmlNode original = GetSectionExtension(capturedList).Clone();
            return () => {
                OfficeXmlNode root = presentation.Root;
                OfficeXmlNode list = root.Child(SectionPresentationNamespace, "extLst");
                OfficeXmlNode current = GetSectionExtension(list);
                if (original.Exists) {
                    if (list.IsNone) list = root.Append(root.CreateElement(SectionPresentationNamespace, "extLst"));
                    list.InsertBefore(original.Clone(), current);
                    if (current.Exists) current.Remove();
                } else if (current.Exists) {
                    current.Remove();
                    if (capturedList.IsNone && list.ChildNodes().Count == 0) list.Remove();
                }
                presentation.MarkEdited();
            };
        }
    }
}
