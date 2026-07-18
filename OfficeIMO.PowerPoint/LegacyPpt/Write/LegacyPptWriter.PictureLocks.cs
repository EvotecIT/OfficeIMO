using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        private const uint PictureProtectionRewriteMask =
            (1U << 7) | (1U << 23)
            | (1U << 8) | (1U << 24)
            | (1U << 9) | (1U << 25)
            | (1U << 10) | (1U << 26)
            | (1U << 11) | (1U << 27)
            | (1U << 12) | (1U << 28)
            | (1U << 14) | (1U << 30)
            | (1U << 15) | (1U << 31);
        private const uint PictureShapeBooleanRewriteMask =
            (1U << 11) | (1U << 27)
            | (1U << 12) | (1U << 28);

        private static bool TryReadPictureProtectionForWrite(
            PowerPointPicture picture, out uint? propertyValue,
            out string? reason) {
            if (picture == null) throw new ArgumentNullException(
                nameof(picture));
            propertyValue = null;
            if (picture.Element is not P.Picture source) {
                reason = "The picture shape has no DrawingML picture element.";
                return false;
            }
            P.NonVisualPictureDrawingProperties? nonVisual = source
                .NonVisualPictureProperties?
                .NonVisualPictureDrawingProperties;
            if (nonVisual == null) {
                reason = null;
                return true;
            }
            if (nonVisual.ExtendedAttributes.Any()
                || nonVisual.ChildElements.Any(child =>
                    child is not A.PictureLocks)) {
                reason = "Extended non-visual picture properties have no native OfficeArt protection mapping.";
                return false;
            }
            A.PictureLocks[] lockElements = nonVisual
                .Elements<A.PictureLocks>().ToArray();
            if (lockElements.Length > 1) {
                reason = "A picture contains duplicate DrawingML lock elements.";
                return false;
            }
            if (lockElements.Length == 0) {
                reason = null;
                return true;
            }
            A.PictureLocks locks = lockElements[0];
            if (locks.HasChildren || locks.ExtendedAttributes.Any()) {
                reason = "Extended DrawingML picture-lock content has no native OfficeArt protection mapping.";
                return false;
            }
            if (IsTrue(locks.NoResize)) {
                reason = "DrawingML no-resize picture locking has no exact classic OfficeArt protection flag.";
                return false;
            }
            if (IsTrue(locks.NoChangeArrowheads)) {
                reason = "DrawingML arrowhead-change picture locking has no classic OfficeArt picture equivalent.";
                return false;
            }
            uint value = 0;
            AddProtectionBoolean(ref value, locks.NoRotation, 7);
            AddProtectionBoolean(ref value, locks.NoChangeAspect, 8);
            AddProtectionBoolean(ref value, locks.NoMove, 9);
            AddProtectionBoolean(ref value, locks.NoSelection, 10);
            AddProtectionBoolean(ref value, locks.NoCrop, 11);
            AddProtectionBoolean(ref value, locks.NoEditPoints, 12);
            AddProtectionBoolean(ref value, locks.NoAdjustHandles, 14);
            AddProtectionBoolean(ref value, locks.NoGrouping, 15);
            propertyValue = value == 0 ? null : value;
            reason = null;
            return true;
        }

        private static void AddPictureProtectionProperties(
            ICollection<LegacyPptWriterFoptProperty> properties,
            PowerPointPicture picture) {
            if (!TryReadPictureProtectionForWrite(picture,
                    out uint? propertyValue, out string? reason)) {
                throw new NotSupportedException(reason);
            }
            if (propertyValue.HasValue) {
                properties.Add(new LegacyPptWriterFoptProperty(0x007F,
                    propertyValue.Value));
            }
            P.Picture source = (P.Picture)picture.Element;
            P.NonVisualPictureDrawingProperties? nonVisual = source
                .NonVisualPictureProperties?
                .NonVisualPictureDrawingProperties;
            A.PictureLocks? locks = nonVisual?
                .GetFirstChild<A.PictureLocks>();
            uint shapeBooleanValue = 0;
            AddProtectionBoolean(ref shapeBooleanValue,
                nonVisual?.PreferRelativeResize, 11);
            AddProtectionBoolean(ref shapeBooleanValue,
                locks?.NoChangeShapeType, 12);
            if (shapeBooleanValue != 0) {
                MergeBooleanProperty(properties, 0x033F,
                    shapeBooleanValue);
            }
        }

        private static void MergeBooleanProperty(
            ICollection<LegacyPptWriterFoptProperty> properties,
            ushort propertyId, uint value) {
            LegacyPptWriterFoptProperty[] existing = properties
                .Where(property => property.PropertyId == propertyId)
                .ToArray();
            foreach (LegacyPptWriterFoptProperty property in existing) {
                value |= property.Value;
                properties.Remove(property);
            }
            properties.Add(new LegacyPptWriterFoptProperty(propertyId,
                value));
        }

        private static bool IsTrue(BooleanValue? value) =>
            value?.Value == true;

        private static void AddProtectionBoolean(ref uint propertyValue,
            BooleanValue? value, int useBit) {
            if (value == null) return;
            propertyValue |= 1U << useBit;
            if (value.Value) {
                propertyValue |= 1U << checked(useBit + 16);
            }
        }
    }
}
