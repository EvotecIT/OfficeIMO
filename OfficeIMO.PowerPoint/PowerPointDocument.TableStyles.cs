using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        internal const string TableStylesRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles";
        private const string TableStylesContentType = "application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml";

        /// <summary>
        /// The presentation's table styles part, added with the built-in styles when it has none; none when the
        /// presentation can't be given one.
        /// </summary>
        internal PowerPointPartRef EnsureTableStylesPart() {
            if (IsSdkLoaded) return SdkEnsureTableStylesPart();
            PowerPointPartRef presentation = PresentationPartRef;
            if (presentation.IsNone) return default;
            PowerPointPartRef styles = presentation.RelatedPart(TableStylesRelationshipType);
            if (!styles.IsNone || GetPackage(edit: true) is not { } package) return styles;
            string partName = package.PartExists("/ppt/tableStyles.xml")
                ? NextPartName(package, TableStylesContentType, "/ppt/tableStyles", ".xml")
                : "/ppt/tableStyles.xml";
            package.CreateBinaryPart(partName, TableStylesContentType, PowerPointUtils.TableStylesBytes.Value);
            package.AddRelationship(presentation.Name, TableStylesRelationshipType, OfficeOpcPartNames.GetRelativeTarget(presentation.Name, partName));
            InvalidatePackageRelationships();
            MarkPackageEdited();
            return PowerPointPartRef.Package(this, partName);
        }
    }
}
