using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.LegacyDoc.Model {
    internal static class LegacyDocNumberFormatMapper {
        private static readonly NfcMap[] Mappings = {
            new NfcMap(0, NumberFormatValues.Decimal),
            new NfcMap(1, NumberFormatValues.UpperRoman),
            new NfcMap(2, NumberFormatValues.LowerRoman),
            new NfcMap(3, NumberFormatValues.UpperLetter),
            new NfcMap(4, NumberFormatValues.LowerLetter),
            new NfcMap(5, NumberFormatValues.Ordinal),
            new NfcMap(6, NumberFormatValues.CardinalText),
            new NfcMap(7, NumberFormatValues.OrdinalText),
            new NfcMap(8, NumberFormatValues.Hex),
            new NfcMap(9, NumberFormatValues.Chicago),
            new NfcMap(10, NumberFormatValues.IdeographDigital),
            new NfcMap(11, NumberFormatValues.JapaneseCounting),
            new NfcMap(12, NumberFormatValues.Aiueo),
            new NfcMap(13, NumberFormatValues.Iroha),
            new NfcMap(14, NumberFormatValues.DecimalFullWidth),
            new NfcMap(15, NumberFormatValues.DecimalHalfWidth),
            new NfcMap(16, NumberFormatValues.JapaneseLegal),
            new NfcMap(17, NumberFormatValues.JapaneseDigitalTenThousand),
            new NfcMap(18, NumberFormatValues.DecimalEnclosedCircle),
            new NfcMap(19, NumberFormatValues.DecimalFullWidth2),
            new NfcMap(20, NumberFormatValues.AiueoFullWidth),
            new NfcMap(21, NumberFormatValues.IrohaFullWidth),
            new NfcMap(22, NumberFormatValues.DecimalZero),
            new NfcMap(23, NumberFormatValues.Bullet),
            new NfcMap(24, NumberFormatValues.Ganada),
            new NfcMap(25, NumberFormatValues.Chosung),
            new NfcMap(26, NumberFormatValues.DecimalEnclosedFullstop),
            new NfcMap(27, NumberFormatValues.DecimalEnclosedParen),
            new NfcMap(28, NumberFormatValues.DecimalEnclosedCircleChinese),
            new NfcMap(29, NumberFormatValues.IdeographEnclosedCircle),
            new NfcMap(30, NumberFormatValues.IdeographTraditional),
            new NfcMap(31, NumberFormatValues.IdeographZodiac),
            new NfcMap(32, NumberFormatValues.IdeographZodiacTraditional),
            new NfcMap(33, NumberFormatValues.TaiwaneseCounting),
            new NfcMap(34, NumberFormatValues.IdeographLegalTraditional),
            new NfcMap(35, NumberFormatValues.TaiwaneseCountingThousand),
            new NfcMap(36, NumberFormatValues.TaiwaneseDigital),
            new NfcMap(37, NumberFormatValues.ChineseCounting),
            new NfcMap(38, NumberFormatValues.ChineseLegalSimplified),
            new NfcMap(39, NumberFormatValues.ChineseCountingThousand),
            new NfcMap(41, NumberFormatValues.KoreanDigital),
            new NfcMap(42, NumberFormatValues.KoreanCounting),
            new NfcMap(43, NumberFormatValues.KoreanLegal),
            new NfcMap(44, NumberFormatValues.KoreanDigital2),
            new NfcMap(45, NumberFormatValues.Hebrew1),
            new NfcMap(46, NumberFormatValues.ArabicAlpha),
            new NfcMap(47, NumberFormatValues.Hebrew2),
            new NfcMap(48, NumberFormatValues.ArabicAbjad),
            new NfcMap(49, NumberFormatValues.HindiVowels),
            new NfcMap(50, NumberFormatValues.HindiConsonants),
            new NfcMap(51, NumberFormatValues.HindiNumbers),
            new NfcMap(52, NumberFormatValues.HindiCounting),
            new NfcMap(53, NumberFormatValues.ThaiLetters),
            new NfcMap(54, NumberFormatValues.ThaiNumbers),
            new NfcMap(55, NumberFormatValues.ThaiCounting),
            new NfcMap(56, NumberFormatValues.VietnameseCounting),
            new NfcMap(57, NumberFormatValues.NumberInDash),
            new NfcMap(58, NumberFormatValues.RussianLower),
            new NfcMap(59, NumberFormatValues.RussianUpper)
        };

        internal static NumberFormatValues? FromNfc(byte nfc) {
            foreach (NfcMap mapping in Mappings) {
                if (mapping.Nfc == nfc) {
                    return mapping.Format;
                }
            }

            return null;
        }

        internal static byte? ToNfc(NumberFormatValues format) {
            if (format == NumberFormatValues.Bullet) {
                return null;
            }

            foreach (NfcMap mapping in Mappings) {
                if (mapping.Format == format) {
                    return mapping.Nfc;
                }
            }

            return null;
        }

        private readonly struct NfcMap {
            internal NfcMap(byte nfc, NumberFormatValues format) {
                Nfc = nfc;
                Format = format;
            }

            internal byte Nfc { get; }

            internal NumberFormatValues Format { get; }
        }
    }
}
