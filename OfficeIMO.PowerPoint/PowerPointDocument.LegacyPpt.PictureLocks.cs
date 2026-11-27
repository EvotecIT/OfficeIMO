using OfficeIMO.Drawing.Binary;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        /// <summary>Replaces the <c>p:cNvPicPr</c> of a projected picture with the binary picture's locks.</summary>
        private static void ApplyLegacyPictureEditingProperties(
            OfficeXmlNode picture, LegacyPptShape source) {
            OfficeXmlNode nonVisual = picture.Child(LegacyPptXml.P, "nvPicPr");
            if (nonVisual.IsNone) return;
            WriteLegacyNonVisualPictureDrawingProperties(
                nonVisual.SetChild(LegacyPptXml.P, "cNvPicPr", LegacyPptImportXml.NonVisualPictureOrder), source);
        }

        private static void WriteLegacyNonVisualPictureDrawingProperties(
            OfficeXmlNode target, LegacyPptShape source) {
            target.Set("preferRelativeResize", source.Style.PreferRelativeResize);
            OfficeArtShapeProtection protection =
                OfficeArtShapeProtection.Decode(source.Style.Properties);
            target.AddA("picLocks")
                .Set("noGrp", protection.LockAgainstGrouping)
                .Set("noSelect", protection.LockAgainstSelect)
                .Set("noRot", protection.LockRotation)
                .Set("noChangeAspect", protection.LockAspectRatio)
                .Set("noMove", protection.LockPosition)
                .Set("noEditPoints", protection.LockVertices)
                .Set("noAdjustHandles", protection.LockAdjustHandles)
                .Set("noChangeShapeType", source.Style.LockShapeType)
                .Set("noCrop", protection.LockCropping);
        }
    }
}
