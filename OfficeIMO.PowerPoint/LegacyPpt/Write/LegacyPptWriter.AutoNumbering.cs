using System.Text;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        private const ushort RecordStyleTextProp9AtomForWrite = 0x0FAC;

        private static bool TryBuildStyleTextProp9Record(
            IReadOnlyList<A.Paragraph> paragraphs,
            LegacyPptWriterPictureBulletCatalog pictureBullets,
            out byte[]? record,
            out string? reason) {
            record = null;
            reason = null;
            if (!paragraphs.Any(paragraph => paragraph
                    .ParagraphProperties?.ChildElements.Any(child => child
                        is A.AutoNumberedBullet or A.PictureBullet) == true)) {
                return true;
            }
            using var payload = new MemoryStream();
            var numberingState = new Dictionary<int, int>();
            foreach (A.Paragraph paragraph in paragraphs) {
                A.AutoNumberedBullet? numbering = paragraph
                    .ParagraphProperties?
                    .GetFirstChild<A.AutoNumberedBullet>();
                A.PictureBullet? pictureBullet = paragraph
                    .ParagraphProperties?
                    .GetFirstChild<A.PictureBullet>();
                int? numberingStart = null;
                if (numbering != null) {
                    int level = paragraph.ParagraphProperties?.Level?.Value
                        ?? 0;
                    numberingStart = numbering.StartAt?.Value
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
            Stream output, A.AutoNumberedBullet? numbering,
            int? numberingStart,
            A.PictureBullet? pictureBullet,
            LegacyPptWriterPictureBulletCatalog pictureBullets,
            out string? reason) {
            reason = null;
            if (numbering != null && pictureBullet != null) {
                reason = "A paragraph cannot use automatic numbering and a picture bullet at the same time.";
                return false;
            }
            ushort? pictureReference = null;
            if (pictureBullet != null) {
                if (!pictureBullets.TryGetIndex(pictureBullet,
                        out ushort index)) {
                    reason = "A picture bullet is not present in the bounded PPT9 picture-bullet catalog.";
                    return false;
                }
                pictureReference = index;
            }
            if (numbering == null && !pictureReference.HasValue) {
                WriteUInt32(output, 0);
                return true;
            }
            if (numbering == null) {
                WriteUInt32(output, 1U << 23);
                WriteUInt16(output, pictureReference!.Value);
                return true;
            }
            if (!HasOnlyAttributes(numbering, "type", "startAt")
                || numbering.ChildElements.Count != 0
                || numbering.Type?.HasValue != true) {
                reason = "An automatic-numbering bullet has unsupported metadata or no numbering scheme.";
                return false;
            }
            if (!TryMapAutoNumberScheme(numbering.Type.Value,
                    out ushort scheme)) {
                reason = "An automatic-numbering scheme has no classic binary PowerPoint equivalent.";
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
            if (style9Record == null) throw new ArgumentNullException(
                nameof(style9Record));
            byte[] tagName = BuildRecord(version: 0, instance: 0,
                RecordCString, Encoding.Unicode.GetBytes(Ppt9TagName));
            byte[] data = BuildRecord(version: 0, instance: 0,
                RecordBinaryTagDataBlob, style9Record);
            return BuildContainer(RecordProgBinaryTag, instance: 0,
                new[] { tagName, data });
        }

        private static bool TryMapAutoNumberScheme(
            A.TextAutoNumberSchemeValues value, out ushort scheme) {
            if (value == A.TextAutoNumberSchemeValues.AlphaLowerCharacterPeriod) scheme = 0x0000;
            else if (value == A.TextAutoNumberSchemeValues.AlphaUpperCharacterPeriod) scheme = 0x0001;
            else if (value == A.TextAutoNumberSchemeValues.ArabicParenR) scheme = 0x0002;
            else if (value == A.TextAutoNumberSchemeValues.ArabicPeriod) scheme = 0x0003;
            else if (value == A.TextAutoNumberSchemeValues.RomanLowerCharacterParenBoth) scheme = 0x0004;
            else if (value == A.TextAutoNumberSchemeValues.RomanLowerCharacterParenR) scheme = 0x0005;
            else if (value == A.TextAutoNumberSchemeValues.RomanLowerCharacterPeriod) scheme = 0x0006;
            else if (value == A.TextAutoNumberSchemeValues.RomanUpperCharacterPeriod) scheme = 0x0007;
            else if (value == A.TextAutoNumberSchemeValues.AlphaLowerCharacterParenBoth) scheme = 0x0008;
            else if (value == A.TextAutoNumberSchemeValues.AlphaLowerCharacterParenR) scheme = 0x0009;
            else if (value == A.TextAutoNumberSchemeValues.AlphaUpperCharacterParenBoth) scheme = 0x000A;
            else if (value == A.TextAutoNumberSchemeValues.AlphaUpperCharacterParenR) scheme = 0x000B;
            else if (value == A.TextAutoNumberSchemeValues.ArabicParenBoth) scheme = 0x000C;
            else if (value == A.TextAutoNumberSchemeValues.ArabicPlain) scheme = 0x000D;
            else if (value == A.TextAutoNumberSchemeValues.RomanUpperCharacterParenBoth) scheme = 0x000E;
            else if (value == A.TextAutoNumberSchemeValues.RomanUpperCharacterParenR) scheme = 0x000F;
            else if (value == A.TextAutoNumberSchemeValues.EastAsianSimplifiedChinesePlain) scheme = 0x0010;
            else if (value == A.TextAutoNumberSchemeValues.EastAsianSimplifiedChinesePeriod) scheme = 0x0011;
            else if (value == A.TextAutoNumberSchemeValues.CircleNumberDoubleBytePlain) scheme = 0x0012;
            else if (value == A.TextAutoNumberSchemeValues.CircleNumberWingdingsWhitePlain) scheme = 0x0013;
            else if (value == A.TextAutoNumberSchemeValues.CircleNumberWingdingsBlackPlain) scheme = 0x0014;
            else if (value == A.TextAutoNumberSchemeValues.EastAsianTraditionalChinesePlain) scheme = 0x0015;
            else if (value == A.TextAutoNumberSchemeValues.EastAsianTraditionalChinesePeriod) scheme = 0x0016;
            else if (value == A.TextAutoNumberSchemeValues.Arabic1Minus) scheme = 0x0017;
            else if (value == A.TextAutoNumberSchemeValues.Arabic2Minus) scheme = 0x0018;
            else if (value == A.TextAutoNumberSchemeValues.Hebrew2Minus) scheme = 0x0019;
            else if (value == A.TextAutoNumberSchemeValues.EastAsianJapaneseKoreanPlain) scheme = 0x001A;
            else if (value == A.TextAutoNumberSchemeValues.EastAsianJapaneseKoreanPeriod) scheme = 0x001B;
            else if (value == A.TextAutoNumberSchemeValues.ArabicDoubleBytePlain) scheme = 0x001C;
            else if (value == A.TextAutoNumberSchemeValues.ArabicDoubleBytePeriod) scheme = 0x001D;
            else if (value == A.TextAutoNumberSchemeValues.ThaiAlphaPeriod) scheme = 0x001E;
            else if (value == A.TextAutoNumberSchemeValues.ThaiAlphaParenthesisRight) scheme = 0x001F;
            else if (value == A.TextAutoNumberSchemeValues.ThaiAlphaParenthesisBoth) scheme = 0x0020;
            else if (value == A.TextAutoNumberSchemeValues.ThaiNumberPeriod) scheme = 0x0021;
            else if (value == A.TextAutoNumberSchemeValues.ThaiNumberParenthesisRight) scheme = 0x0022;
            else if (value == A.TextAutoNumberSchemeValues.ThaiNumberParenthesisBoth) scheme = 0x0023;
            else if (value == A.TextAutoNumberSchemeValues.HindiAlphaPeriod) scheme = 0x0024;
            else if (value == A.TextAutoNumberSchemeValues.HindiNumPeriod) scheme = 0x0025;
            else if (value == A.TextAutoNumberSchemeValues.EastAsianJapaneseDoubleBytePeriod) scheme = 0x0026;
            else if (value == A.TextAutoNumberSchemeValues.HindiNumberParenthesisRight) scheme = 0x0027;
            else if (value == A.TextAutoNumberSchemeValues.HindiAlpha1Period) scheme = 0x0028;
            else {
                scheme = 0;
                return false;
            }
            return true;
        }
    }
}
