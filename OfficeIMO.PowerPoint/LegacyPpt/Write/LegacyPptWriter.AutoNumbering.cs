using System.Text;
using static OfficeIMO.PowerPoint.LegacyPpt.Internal.LegacyPptXml;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        private const ushort RecordStyleTextProp9AtomForWrite = 0x0FAC;

        private static bool TryBuildStyleTextProp9Record(
            IReadOnlyList<OfficeXmlNode> paragraphs,
            LegacyPptWriterPictureBulletCatalog pictureBullets,
            out byte[]? record,
            out string? reason) {
            record = null;
            reason = null;
            if (!paragraphs.Any(paragraph => HasStyle9Bullet(paragraph.Child(A, "pPr")))) {
                return true;
            }
            using var payload = new MemoryStream();
            var numberingState = new Dictionary<int, int>();
            foreach (OfficeXmlNode paragraph in paragraphs) {
                OfficeXmlNode properties = paragraph.Child(A, "pPr");
                OfficeXmlNode numbering = properties.Child(A, "buAutoNum");
                OfficeXmlNode pictureBullet = properties.Child(A, "buBlip");
                int? numberingStart = null;
                if (!numbering.IsNone) {
                    int level = properties.Int("lvl") ?? 0;
                    numberingStart = numbering.Int("startAt")
                        ?? (numberingState.TryGetValue(level,
                                out int previous)
                            ? checked(previous + 1)
                            : 1);
                    numberingState[level] = numberingStart.Value;
                }
                if (!TryWriteAutomaticNumberingException9(payload,
                        numbering, numberingStart, pictureBullet,
                        pictureBullets,
                        out reason)) return false;
                // StyleTextProp9 character and special-information
                // exceptions are empty for DrawingML automatic numbering.
                WriteUInt32(payload, 0);
                WriteUInt32(payload, 0);
            }
            record = BuildRecord(version: 0, instance: 0,
                RecordStyleTextProp9AtomForWrite, payload.ToArray());
            return true;
        }

        private static bool TryWriteAutomaticNumberingException9(
            Stream output, OfficeXmlNode numbering,
            int? numberingStart,
            OfficeXmlNode pictureBullet,
            LegacyPptWriterPictureBulletCatalog pictureBullets,
            out string? reason) {
            reason = null;
            if (!numbering.IsNone && !pictureBullet.IsNone) {
                reason = "A paragraph cannot use automatic numbering and a picture bullet at the same time.";
                return false;
            }
            ushort? pictureReference = null;
            if (!pictureBullet.IsNone) {
                if (!pictureBullets.TryGetIndex(pictureBullet,
                        out ushort index)) {
                    reason = "A picture bullet is not present in the bounded PPT9 picture-bullet catalog.";
                    return false;
                }
                pictureReference = index;
            }
            if (numbering.IsNone && !pictureReference.HasValue) {
                WriteUInt32(output, 0);
                return true;
            }
            if (numbering.IsNone) {
                WriteUInt32(output, 1U << 23);
                WriteUInt16(output, pictureReference!.Value);
                return true;
            }
            // Every ST_TextAutonumberScheme token has a binary scheme, so an unmapped token is one outside the schema.
            if (!HasOnlyAttributes(numbering, "type", "startAt")
                || HasChildElements(numbering)
                || !TryMapAutoNumberScheme(numbering.Attribute("type"),
                    out ushort scheme)) {
                reason = "An automatic-numbering bullet has unsupported metadata or no numbering scheme.";
                return false;
            }
            int rawStart = numberingStart ?? 1;
            if (rawStart < 1 || rawStart > short.MaxValue) {
                reason = "An automatic-numbering start value lies outside the classic binary PowerPoint signed 16-bit range.";
                return false;
            }
            uint masks = (1U << 24) | (1U << 25);
            if (pictureReference.HasValue) masks |= 1U << 23;
            WriteUInt32(output, masks);
            if (pictureReference.HasValue) {
                WriteUInt16(output, pictureReference.Value);
            }
            WriteInt16(output, 1);
            WriteUInt16(output, scheme);
            WriteInt16(output, checked((short)rawStart));
            return true;
        }

        internal static byte[] BuildShapePpt9ProgrammableTagsRecord(
            byte[] style9Record) => BuildContainer(RecordProgTags,
                instance: 0,
                new[] { BuildShapePpt9BinaryTagRecord(style9Record) });

        internal static byte[] BuildShapePpt9BinaryTagRecord(
            byte[] style9Record) {
            ArgumentNullException.ThrowIfNull(style9Record);
            byte[] tagName = BuildRecord(version: 0, instance: 0,
                RecordCString, Encoding.Unicode.GetBytes(Ppt9TagName));
            byte[] data = BuildRecord(version: 0, instance: 0,
                RecordBinaryTagDataBlob, style9Record);
            return BuildContainer(RecordProgBinaryTag, instance: 0,
                new[] { tagName, data });
        }

        private static bool HasStyle9Bullet(OfficeXmlNode properties) {
            foreach (OfficeXmlNode child in properties.Children()) {
                if (child.Is(A, "buAutoNum") || child.Is(A, "buBlip")) return true;
            }
            return false;
        }

        private static bool TryMapAutoNumberScheme(
            string? token, out ushort scheme) {
            int index = token == null ? -1 : Array.IndexOf(AutoNumberSchemeTokens, token);
            scheme = index < 0 ? (ushort)0 : checked((ushort)index);
            return index >= 0;
        }

        // ST_TextAutonumberScheme tokens in the order of the binary scheme values.
        private static readonly string[] AutoNumberSchemeTokens = {
            "alphaLcPeriod", "alphaUcPeriod", "arabicParenR", "arabicPeriod",
            "romanLcParenBoth", "romanLcParenR", "romanLcPeriod", "romanUcPeriod",
            "alphaLcParenBoth", "alphaLcParenR", "alphaUcParenBoth", "alphaUcParenR",
            "arabicParenBoth", "arabicPlain", "romanUcParenBoth", "romanUcParenR",
            "ea1ChsPlain", "ea1ChsPeriod", "circleNumDbPlain", "circleNumWdWhitePlain",
            "circleNumWdBlackPlain", "ea1ChtPlain", "ea1ChtPeriod", "arabic1Minus",
            "arabic2Minus", "hebrew2Minus", "ea1JpnKorPlain", "ea1JpnKorPeriod",
            "arabicDbPlain", "arabicDbPeriod", "thaiAlphaPeriod", "thaiAlphaParenR",
            "thaiAlphaParenBoth", "thaiNumPeriod", "thaiNumParenR", "thaiNumParenBoth",
            "hindiAlphaPeriod", "hindiNumPeriod", "ea1JpnChsDbPeriod", "hindiNumParenR",
            "hindiAlpha1Period"
        };
    }
}
