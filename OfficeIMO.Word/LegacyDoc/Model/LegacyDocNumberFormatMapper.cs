
namespace OfficeIMO.Word.LegacyDoc.Model {
    internal static class LegacyDocNumberFormatMapper {
        private static readonly NfcMap[] Mappings = {
            new NfcMap(0, LegacyDocNumberFormat.Decimal),
            new NfcMap(1, LegacyDocNumberFormat.UpperRoman),
            new NfcMap(2, LegacyDocNumberFormat.LowerRoman),
            new NfcMap(3, LegacyDocNumberFormat.UpperLetter),
            new NfcMap(4, LegacyDocNumberFormat.LowerLetter),
            new NfcMap(5, LegacyDocNumberFormat.Ordinal),
            new NfcMap(6, LegacyDocNumberFormat.CardinalText),
            new NfcMap(7, LegacyDocNumberFormat.OrdinalText),
            new NfcMap(8, LegacyDocNumberFormat.Hex),
            new NfcMap(9, LegacyDocNumberFormat.Chicago),
            new NfcMap(10, LegacyDocNumberFormat.IdeographDigital),
            new NfcMap(11, LegacyDocNumberFormat.JapaneseCounting),
            new NfcMap(12, LegacyDocNumberFormat.Aiueo),
            new NfcMap(13, LegacyDocNumberFormat.Iroha),
            new NfcMap(14, LegacyDocNumberFormat.DecimalFullWidth),
            new NfcMap(15, LegacyDocNumberFormat.DecimalHalfWidth),
            new NfcMap(16, LegacyDocNumberFormat.JapaneseLegal),
            new NfcMap(17, LegacyDocNumberFormat.JapaneseDigitalTenThousand),
            new NfcMap(18, LegacyDocNumberFormat.DecimalEnclosedCircle),
            new NfcMap(19, LegacyDocNumberFormat.DecimalFullWidth2),
            new NfcMap(20, LegacyDocNumberFormat.AiueoFullWidth),
            new NfcMap(21, LegacyDocNumberFormat.IrohaFullWidth),
            new NfcMap(22, LegacyDocNumberFormat.DecimalZero),
            new NfcMap(23, LegacyDocNumberFormat.Bullet),
            new NfcMap(24, LegacyDocNumberFormat.Ganada),
            new NfcMap(25, LegacyDocNumberFormat.Chosung),
            new NfcMap(26, LegacyDocNumberFormat.DecimalEnclosedFullstop),
            new NfcMap(27, LegacyDocNumberFormat.DecimalEnclosedParen),
            new NfcMap(28, LegacyDocNumberFormat.DecimalEnclosedCircleChinese),
            new NfcMap(29, LegacyDocNumberFormat.IdeographEnclosedCircle),
            new NfcMap(30, LegacyDocNumberFormat.IdeographTraditional),
            new NfcMap(31, LegacyDocNumberFormat.IdeographZodiac),
            new NfcMap(32, LegacyDocNumberFormat.IdeographZodiacTraditional),
            new NfcMap(33, LegacyDocNumberFormat.TaiwaneseCounting),
            new NfcMap(34, LegacyDocNumberFormat.IdeographLegalTraditional),
            new NfcMap(35, LegacyDocNumberFormat.TaiwaneseCountingThousand),
            new NfcMap(36, LegacyDocNumberFormat.TaiwaneseDigital),
            new NfcMap(37, LegacyDocNumberFormat.ChineseCounting),
            new NfcMap(38, LegacyDocNumberFormat.ChineseLegalSimplified),
            new NfcMap(39, LegacyDocNumberFormat.ChineseCountingThousand),
            new NfcMap(41, LegacyDocNumberFormat.KoreanDigital),
            new NfcMap(42, LegacyDocNumberFormat.KoreanCounting),
            new NfcMap(43, LegacyDocNumberFormat.KoreanLegal),
            new NfcMap(44, LegacyDocNumberFormat.KoreanDigital2),
            new NfcMap(45, LegacyDocNumberFormat.Hebrew1),
            new NfcMap(46, LegacyDocNumberFormat.ArabicAlpha),
            new NfcMap(47, LegacyDocNumberFormat.Hebrew2),
            new NfcMap(48, LegacyDocNumberFormat.ArabicAbjad),
            new NfcMap(49, LegacyDocNumberFormat.HindiVowels),
            new NfcMap(50, LegacyDocNumberFormat.HindiConsonants),
            new NfcMap(51, LegacyDocNumberFormat.HindiNumbers),
            new NfcMap(52, LegacyDocNumberFormat.HindiCounting),
            new NfcMap(53, LegacyDocNumberFormat.ThaiLetters),
            new NfcMap(54, LegacyDocNumberFormat.ThaiNumbers),
            new NfcMap(55, LegacyDocNumberFormat.ThaiCounting),
            new NfcMap(56, LegacyDocNumberFormat.VietnameseCounting),
            new NfcMap(57, LegacyDocNumberFormat.NumberInDash),
            new NfcMap(58, LegacyDocNumberFormat.RussianLower),
            new NfcMap(59, LegacyDocNumberFormat.RussianUpper)
        };

        internal static LegacyDocNumberFormat? FromNfc(byte nfc) {
            foreach (NfcMap mapping in Mappings) {
                if (mapping.Nfc == nfc) {
                    return mapping.Format;
                }
            }

            return null;
        }

        internal static byte? ToNfc(LegacyDocNumberFormat format) {
            if (format == LegacyDocNumberFormat.Bullet) {
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
            internal NfcMap(byte nfc, LegacyDocNumberFormat format) {
                Nfc = nfc;
                Format = format;
            }

            internal byte Nfc { get; }

            internal LegacyDocNumberFormat Format { get; }
        }
    }
}
