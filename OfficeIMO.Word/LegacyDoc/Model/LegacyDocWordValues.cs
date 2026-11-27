namespace OfficeIMO.Word.LegacyDoc.Model {
    // WordprocessingML simple-type values the legacy DOC model carries. Each member name is the schema value with an
    // upper-case first letter, so ToXml and TryParseXml convert by name.

    /// <summary>ST_SectionMark.</summary>
    internal enum LegacyDocSectionMark { NextPage, NextColumn, Continuous, EvenPage, OddPage }

    /// <summary>ST_PageOrientation.</summary>
    internal enum LegacyDocPageOrientation { Portrait, Landscape }

    /// <summary>ST_VerticalJc.</summary>
    internal enum LegacyDocVerticalJustification { Top, Center, Both, Bottom }

    /// <summary>ST_LineNumberRestart.</summary>
    internal enum LegacyDocLineNumberRestart { NewPage, NewSection, Continuous }

    /// <summary>ST_FtnPos.</summary>
    internal enum LegacyDocFootnotePosition { PageBottom, BeneathText, SectEnd, DocEnd }

    /// <summary>ST_EdnPos.</summary>
    internal enum LegacyDocEndnotePosition { SectEnd, DocEnd }

    /// <summary>ST_RestartNumber.</summary>
    internal enum LegacyDocRestartNumber { Continuous, EachSect, EachPage }

    /// <summary>ST_HdrFtr.</summary>
    internal enum LegacyDocHeaderFooterType { Even, Default, First }

    /// <summary>ST_NumberFormat, the values a DOC number format code maps to.</summary>
    internal enum LegacyDocNumberFormat {
        Decimal, UpperRoman, LowerRoman, UpperLetter, LowerLetter, Ordinal, CardinalText, OrdinalText, Hex, Chicago,
        IdeographDigital, JapaneseCounting, Aiueo, Iroha, DecimalFullWidth, DecimalHalfWidth, JapaneseLegal,
        JapaneseDigitalTenThousand, DecimalEnclosedCircle, DecimalFullWidth2, AiueoFullWidth, IrohaFullWidth, DecimalZero,
        Bullet, Ganada, Chosung, DecimalEnclosedFullstop, DecimalEnclosedParen, DecimalEnclosedCircleChinese,
        IdeographEnclosedCircle, IdeographTraditional, IdeographZodiac, IdeographZodiacTraditional, TaiwaneseCounting,
        IdeographLegalTraditional, TaiwaneseCountingThousand, TaiwaneseDigital, ChineseCounting, ChineseLegalSimplified,
        ChineseCountingThousand, KoreanDigital, KoreanCounting, KoreanLegal, KoreanDigital2, Hebrew1, ArabicAlpha, Hebrew2,
        ArabicAbjad, HindiVowels, HindiConsonants, HindiNumbers, HindiCounting, ThaiLetters, ThaiNumbers, ThaiCounting,
        VietnameseCounting, NumberInDash, RussianLower, RussianUpper, None
    }

    internal static class LegacyDocWordValues {
        /// <summary>The schema value of <paramref name="value"/>.</summary>
        internal static string ToXml<TEnum>(this TEnum value) where TEnum : struct, Enum {
            string name = value.ToString();
            return char.ToLowerInvariant(name[0]) + name.Substring(1);
        }

        /// <summary>Reads a schema value, or returns null when it is not one of <typeparamref name="TEnum"/>.</summary>
        internal static TEnum? TryParseXml<TEnum>(string? xml) where TEnum : struct, Enum {
            if (string.IsNullOrEmpty(xml) || !char.IsLower(xml![0])) return null;
            return Enum.TryParse(char.ToUpperInvariant(xml[0]) + xml.Substring(1), ignoreCase: false, out TEnum value)
                && Enum.IsDefined(typeof(TEnum), value)
                ? value
                : null;
        }
    }
}
