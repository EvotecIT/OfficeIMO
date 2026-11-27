using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using static OfficeIMO.PowerPoint.LegacyPpt.Write.LegacyPptWriterXml;

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
            OfficeXmlNode source, out uint? propertyValue,
            out string? reason) {
            propertyValue = null;
            if (!source.Is(LegacyPptXml.P, "pic")) {
                reason = "The picture shape has no DrawingML picture element.";
                return false;
            }
            OfficeXmlNode nonVisual = source.Child(LegacyPptXml.P, "nvPicPr")
                .Child(LegacyPptXml.P, "cNvPicPr");
            if (nonVisual.IsNone) {
                reason = null;
                return true;
            }
            if (HasExtendedAttributes(nonVisual, "preferRelativeResize")
                || AnyChildOtherThan(nonVisual, LegacyPptXml.A, "picLocks")) {
                reason = "Extended non-visual picture properties have no native OfficeArt protection mapping.";
                return false;
            }
            int lockCount = nonVisual.Children(LegacyPptXml.A, "picLocks").Count;
            if (lockCount > 1) {
                reason = "A picture contains duplicate DrawingML lock elements.";
                return false;
            }
            if (lockCount == 0) {
                reason = null;
                return true;
            }
            OfficeXmlNode locks = nonVisual.Child(LegacyPptXml.A, "picLocks");
            if (HasChildren(locks) || HasExtendedAttributes(locks,
                    PictureLockAttributes)) {
                reason = "Extended DrawingML picture-lock content has no native OfficeArt protection mapping.";
                return false;
            }
            if (locks.Bool("noResize") == true) {
                reason = "DrawingML no-resize picture locking has no exact classic OfficeArt protection flag.";
                return false;
            }
            if (locks.Bool("noChangeArrowheads") == true) {
                reason = "DrawingML arrowhead-change picture locking has no classic OfficeArt picture equivalent.";
                return false;
            }
            uint value = 0;
            AddProtectionBoolean(ref value, locks, "noRot", 7);
            AddProtectionBoolean(ref value, locks, "noChangeAspect", 8);
            AddProtectionBoolean(ref value, locks, "noMove", 9);
            AddProtectionBoolean(ref value, locks, "noSelect", 10);
            AddProtectionBoolean(ref value, locks, "noCrop", 11);
            AddProtectionBoolean(ref value, locks, "noEditPoints", 12);
            AddProtectionBoolean(ref value, locks, "noAdjustHandles", 14);
            AddProtectionBoolean(ref value, locks, "noGrp", 15);
            propertyValue = value == 0 ? null : value;
            reason = null;
            return true;
        }

        // The attributes CT_PictureLocking defines.
        private static readonly string[] PictureLockAttributes = {
            "noGrp", "noSelect", "noRot", "noChangeAspect", "noMove",
            "noResize", "noEditPoints", "noAdjustHandles",
            "noChangeArrowheads", "noChangeShapeType", "noCrop"
        };

        private static void AddPictureProtectionProperties(
            ICollection<LegacyPptWriterFoptProperty> properties,
            OfficeXmlNode source) {
            if (!TryReadPictureProtectionForWrite(source,
                    out uint? propertyValue, out string? reason)) {
                throw new NotSupportedException(reason);
            }
            if (propertyValue.HasValue) {
                properties.Add(new LegacyPptWriterFoptProperty(0x007F,
                    propertyValue.Value));
            }
            OfficeXmlNode nonVisual = source.Child(LegacyPptXml.P, "nvPicPr")
                .Child(LegacyPptXml.P, "cNvPicPr");
            OfficeXmlNode locks = nonVisual.Child(LegacyPptXml.A, "picLocks");
            uint shapeBooleanValue = 0;
            AddProtectionBoolean(ref shapeBooleanValue, nonVisual,
                "preferRelativeResize", 11);
            AddProtectionBoolean(ref shapeBooleanValue, locks,
                "noChangeShapeType", 12);
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

        // An xsd:boolean attribute that is present sets the use bit; true also sets the value bit.
        private static void AddProtectionBoolean(ref uint propertyValue,
            OfficeXmlNode element, string localName, int useBit) {
            if (element.IsNone || element.Attribute(localName) == null) return;
            propertyValue |= 1U << useBit;
            if (element.Bool(localName) == true) {
                propertyValue |= 1U << checked(useBit + 16);
            }
        }
    }
}
