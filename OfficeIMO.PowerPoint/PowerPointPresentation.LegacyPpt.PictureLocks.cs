using DocumentFormat.OpenXml;
using OfficeIMO.Drawing.Binary;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointPresentation {
        private static void ApplyLegacyPictureEditingProperties(
            P.Picture picture, LegacyPptShape source) {
            if (picture.NonVisualPictureProperties == null) return;
            picture.NonVisualPictureProperties
                .NonVisualPictureDrawingProperties =
                CreateLegacyNonVisualPictureDrawingProperties(source);
        }

        private static P.NonVisualPictureDrawingProperties
            CreateLegacyNonVisualPictureDrawingProperties(
                LegacyPptShape source) {
            return new P.NonVisualPictureDrawingProperties(
                CreateLegacyPictureLocks(source)) {
                PreferRelativeResize = ToBooleanValue(
                    source.Style.PreferRelativeResize)
            };
        }

        private static A.PictureLocks CreateLegacyPictureLocks(
            LegacyPptShape source) {
            OfficeArtShapeProtection protection =
                OfficeArtShapeProtection.Decode(source.Style.Properties);
            return new A.PictureLocks {
                NoGrouping = ToBooleanValue(
                    protection.LockAgainstGrouping),
                NoSelection = ToBooleanValue(
                    protection.LockAgainstSelect),
                NoRotation = ToBooleanValue(protection.LockRotation),
                NoChangeAspect = ToBooleanValue(
                    protection.LockAspectRatio),
                NoMove = ToBooleanValue(protection.LockPosition),
                NoEditPoints = ToBooleanValue(protection.LockVertices),
                NoAdjustHandles = ToBooleanValue(
                    protection.LockAdjustHandles),
                NoCrop = ToBooleanValue(protection.LockCropping),
                NoChangeShapeType = ToBooleanValue(
                    source.Style.LockShapeType)
            };
        }

        private static BooleanValue? ToBooleanValue(bool? value) =>
            value.HasValue ? new BooleanValue(value.Value) : null;
    }
}
