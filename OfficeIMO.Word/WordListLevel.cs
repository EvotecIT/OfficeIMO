using System;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public enum WordListLevelKind {
        None,
        Bullet,
        BulletSquareSymbol,
        BulletBlackCircle,
        BulletDiamondSymbol,
        BulletArrowSymbol,
        BulletSolidRound,
        BulletOpenCircle,
        BulletSquare,
        /// <summary>
        /// The BulletSquare2, not the same as BulletSquare, and not even used by Word
        /// </summary>
        BulletSquare2,
        BulletCheckmark,
        BulletClubs,
        BulletDiamond,
        BulletArrow,

        Decimal,
        DecimalBracket,
        DecimalDot,
        LowerLetter,
        LowerLetterBracket,
        LowerLetterDot,
        UpperLetter,
        UpperLetterBracket,
        UpperLetterDot,
        LowerRoman,
        LowerRomanBracket,
        LowerRomanDot,
        UpperRoman,
        UpperRomanBracket,
        UpperRomanDot
    }
    public class WordListLevel {
        public WordListLevel(Level level) {
            _level = level;
        }

        public Level _level { get; set; }

        /// <summary>
        /// Gets or sets the start numbering value.
        /// </summary>
        public int StartNumberingValue {
            get {
                return _level.Descendants<StartNumberingValue>().First().Val;
            }
            set {
                _level.Descendants<StartNumberingValue>().First().Val = value;
            }
        }

        /// <summary>
        /// Sets the starting number for this level.
        /// </summary>
        /// <param name="value">The starting number.</param>
        /// <returns>The current <see cref="WordListLevel"/> instance.</returns>
        public WordListLevel SetStartNumberingValue(int value) {
            StartNumberingValue = value;
            return this;
        }

        /// <summary>
        /// Gets or sets the text indentation (left indentation) in twentieths of a point.
        /// </summary>
        public int IndentationLeft {
            get {
                return int.Parse(_level.Descendants<Indentation>().First().Left);
            }
            set {
                _level.Descendants<Indentation>().First().Left = value.ToString();
            }
        }

        /// <summary>
        /// Gets or sets the text indentation (left indentation) in centimeters.
        /// </summary>
        public double IndentationLeftCentimeters {
            get => Helpers.ConvertTwentiethsOfPointToCentimeters(IndentationLeft);
            set => IndentationLeft = (int)Math.Round(Helpers.ConvertCentimetersToTwentiethsOfPoint(value));
        }

        /// <summary>
        /// Gets or sets the number position (hanging indentation) in twentieths of a point.
        /// </summary>
        public int IndentationHanging {
            get {
                return int.Parse(_level.Descendants<Indentation>().First().Hanging);
            }
            set {
                _level.Descendants<Indentation>().First().Hanging = value.ToString();
            }
        }

        /// <summary>
        /// Gets or sets the number position (hanging indentation) in centimeters.
        /// </summary>
        public double IndentationHangingCentimeters {
            get => Helpers.ConvertTwentiethsOfPointToCentimeters(IndentationHanging);
            set => IndentationHanging = (int)Math.Round(Helpers.ConvertCentimetersToTwentiethsOfPoint(value));
        }

        /// <summary>
        /// Gets or sets the level text.
        /// </summary>
        public string LevelText {
            get {
                return _level.Descendants<LevelText>().First().Val;
            }
            set {
                _level.Descendants<LevelText>().First().Val = value;
            }
        }

        /// <summary>
        /// Gets or sets the level justification.
        /// </summary>
        public LevelJustificationValues LevelJustification {
            get {
                return _level.Descendants<LevelJustification>().First().Val;
            }
            set {
                _level.Descendants<LevelJustification>().First().Val = value;
            }
        }

        /// <summary>
        /// Removes this level from the list.
        /// </summary>
        public void Remove() {
            _level.Remove();
        }

        /// <summary>
        /// Adds the level using custom simplified list number.
        /// </summary>
        /// <param name="simplifiedListNumbers"></param>
        /// <exception cref="ArgumentOutOfRangeException"></exception>
        public WordListLevel(WordListLevelKind simplifiedListNumbers) {
            switch (simplifiedListNumbers) {
                case WordListLevelKind.Bullet:
                    _level = new Level() {
                        LevelIndex = 0,
                        TemplateCode = "",
                        StartNumberingValue = new StartNumberingValue() { Val = 1 },
                        NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.Bullet },
                        LevelText = new LevelText() { Val = "\u2022" },
                        LevelJustification = new LevelJustification() { Val = LevelJustificationValues.Left },
                        PreviousParagraphProperties = new PreviousParagraphProperties() {
                            Indentation = new Indentation() { Left = "720", Hanging = "360" }
                        }
                    };
                    break;
                case WordListLevelKind.BulletSquareSymbol:
                    _level = new Level() {
                        LevelIndex = 0,
                        TemplateCode = "",
                        StartNumberingValue = new StartNumberingValue() { Val = 1 },
                        NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.Bullet },
                        LevelText = new LevelText() { Val = "\u25A0" },
                        LevelJustification = new LevelJustification() { Val = LevelJustificationValues.Left },
                        PreviousParagraphProperties = new PreviousParagraphProperties() {
                            Indentation = new Indentation() { Left = "720", Hanging = "360" }
                        }
                    };
                    break;
                case WordListLevelKind.BulletBlackCircle:
                    _level = new Level() {
                        LevelIndex = 0,
                        TemplateCode = "",
                        StartNumberingValue = new StartNumberingValue() { Val = 1 },
                        NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.Bullet },
                        LevelText = new LevelText() { Val = "\u25CF" },
                        LevelJustification = new LevelJustification() { Val = LevelJustificationValues.Left },
                        PreviousParagraphProperties = new PreviousParagraphProperties() {
                            Indentation = new Indentation() { Left = "720", Hanging = "360" }
                        }
                    };
                    break;
                case WordListLevelKind.BulletDiamondSymbol:
                    _level = new Level() {
                        LevelIndex = 0,
                        TemplateCode = "",
                        StartNumberingValue = new StartNumberingValue() { Val = 1 },
                        NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.Bullet },
                        LevelText = new LevelText() { Val = "\u25C6" },
                        LevelJustification = new LevelJustification() { Val = LevelJustificationValues.Left },
                        PreviousParagraphProperties = new PreviousParagraphProperties() {
                            Indentation = new Indentation() { Left = "720", Hanging = "360" }
                        }
                    };
                    break;
                case WordListLevelKind.BulletArrowSymbol:
                    _level = new Level() {
                        LevelIndex = 0,
                        TemplateCode = "",
                        StartNumberingValue = new StartNumberingValue() { Val = 1 },
                        NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.Bullet },
                        LevelText = new LevelText() { Val = "\u25BA" },
                        LevelJustification = new LevelJustification() { Val = LevelJustificationValues.Left },
                        PreviousParagraphProperties = new PreviousParagraphProperties() {
                            Indentation = new Indentation() { Left = "720", Hanging = "360" }
                        }
                    };
                    break;
                case WordListLevelKind.BulletSolidRound:
                    _level = new Level() {
                        LevelIndex = 0,
                        TemplateCode = "",
                        StartNumberingValue = new StartNumberingValue() { Val = 1 },
                        NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.Bullet },
                        LevelText = new LevelText() { Val = "·" },
                        LevelJustification = new LevelJustification() { Val = LevelJustificationValues.Left },
                        PreviousParagraphProperties = new PreviousParagraphProperties() {
                            Indentation = new Indentation() { Left = "720", Hanging = "360" }
                        },
                        NumberingSymbolRunProperties = new NumberingSymbolRunProperties() {
                            RunFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol" }
                        }
                    };
                    break;
                case WordListLevelKind.BulletOpenCircle:
                    _level = new Level() {
                        LevelIndex = 0,
                        TemplateCode = "",
                        StartNumberingValue = new StartNumberingValue() { Val = 1 },
                        NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.Bullet },
                        LevelText = new LevelText() { Val = "o" },
                        LevelJustification = new LevelJustification() { Val = LevelJustificationValues.Left },
                        PreviousParagraphProperties = new PreviousParagraphProperties() {
                            Indentation = new Indentation() { Left = "720", Hanging = "360" }
                        },
                        NumberingSymbolRunProperties = new NumberingSymbolRunProperties() {
                            RunFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol" }
                        }
                    };
                    break;
                case WordListLevelKind.BulletSquare2:
                    _level = new Level() {
                        LevelIndex = 0,
                        TemplateCode = "",
                        StartNumberingValue = new StartNumberingValue() { Val = 1 },
                        NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.Bullet },
                        LevelText = new LevelText() { Val = "■" },
                        LevelJustification = new LevelJustification() { Val = LevelJustificationValues.Left },
                        PreviousParagraphProperties = new PreviousParagraphProperties() {
                            Indentation = new Indentation() { Left = "720", Hanging = "360" }
                        }
                    };
                    break;
                case WordListLevelKind.BulletSquare:
                    _level = new Level() {
                        LevelIndex = 0,
                        TemplateCode = "",
                        StartNumberingValue = new StartNumberingValue() { Val = 1 },
                        NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.Bullet },
                        LevelText = new LevelText() { Val = "§" },
                        LevelJustification = new LevelJustification() { Val = LevelJustificationValues.Left },
                        PreviousParagraphProperties = new PreviousParagraphProperties() {
                            Indentation = new Indentation() { Left = "720", Hanging = "360" }
                        },
                        NumberingSymbolRunProperties = new NumberingSymbolRunProperties() {
                            RunFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" }
                        }
                    };
                    break;
                case WordListLevelKind.BulletClubs:
                    _level = new Level() {
                        LevelIndex = 0,
                        TemplateCode = "",
                        StartNumberingValue = new StartNumberingValue() { Val = 1 },
                        NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.Bullet },
                        LevelText = new LevelText() { Val = "v" },
                        LevelJustification = new LevelJustification() { Val = LevelJustificationValues.Left },
                        PreviousParagraphProperties = new PreviousParagraphProperties() {
                            Indentation = new Indentation() { Left = "720", Hanging = "360" }
                        },
                        NumberingSymbolRunProperties = new NumberingSymbolRunProperties() {
                            RunFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" }
                        }
                    };
                    break;
                case WordListLevelKind.BulletArrow:
                    _level = new Level() {
                        LevelIndex = 0,
                        TemplateCode = "",
                        StartNumberingValue = new StartNumberingValue() { Val = 1 },
                        NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.Bullet },
                        LevelText = new LevelText() { Val = "Ø" },
                        LevelJustification = new LevelJustification() { Val = LevelJustificationValues.Left },
                        PreviousParagraphProperties = new PreviousParagraphProperties() {
                            Indentation = new Indentation() { Left = "720", Hanging = "360" }
                        },
                        NumberingSymbolRunProperties = new NumberingSymbolRunProperties() {
                            RunFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" }
                        }
                    };
                    break;
                case WordListLevelKind.BulletDiamond:
                    _level = new Level() {
                        LevelIndex = 0,
                        TemplateCode = "",
                        StartNumberingValue = new StartNumberingValue() { Val = 1 },
                        NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.Bullet },
                        LevelText = new LevelText() { Val = "¨" },
                        LevelJustification = new LevelJustification() { Val = LevelJustificationValues.Left },
                        PreviousParagraphProperties = new PreviousParagraphProperties() {
                            Indentation = new Indentation() { Left = "720", Hanging = "360" }
                        },
                        NumberingSymbolRunProperties = new NumberingSymbolRunProperties() {
                            RunFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol" }
                        }
                    };
                    break;
                case WordListLevelKind.BulletCheckmark:
                    _level = new Level() {
                        LevelIndex = 0,
                        TemplateCode = "",
                        StartNumberingValue = new StartNumberingValue() { Val = 1 },
                        NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.Bullet },
                        LevelText = new LevelText() { Val = "ü" },
                        LevelJustification = new LevelJustification() { Val = LevelJustificationValues.Left },
                        PreviousParagraphProperties = new PreviousParagraphProperties() {
                            Indentation = new Indentation() { Left = "720", Hanging = "360" }
                        },
                        NumberingSymbolRunProperties = new NumberingSymbolRunProperties() {
                            RunFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" }
                        }
                    };
                    break;
                case WordListLevelKind.Decimal:
                    _level = new Level() {
                        LevelIndex = 0,
                        TemplateCode = "",
                        StartNumberingValue = new StartNumberingValue() { Val = 1 },
                        NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.Decimal },
                        LevelText = new LevelText() { Val = "%CurrentLevel" },
                        LevelJustification = new LevelJustification() { Val = LevelJustificationValues.Left },
                        PreviousParagraphProperties = new PreviousParagraphProperties() {
                            Indentation = new Indentation() { Left = "720", Hanging = "360" }
                        }
                    };
                    break;
                case WordListLevelKind.DecimalBracket:
                    _level = new Level() {
                        LevelIndex = 0,
                        TemplateCode = "",
                        StartNumberingValue = new StartNumberingValue() { Val = 1 },
                        NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.Decimal },
                        LevelText = new LevelText() { Val = "%CurrentLevel)" },
                        LevelJustification = new LevelJustification() { Val = LevelJustificationValues.Left },
                        PreviousParagraphProperties = new PreviousParagraphProperties() {
                            Indentation = new Indentation() { Left = "720", Hanging = "360" }
                        }
                    };
                    break;
                case WordListLevelKind.DecimalDot:
                    _level = new Level() {
                        LevelIndex = 0,
                        TemplateCode = "",
                        StartNumberingValue = new StartNumberingValue() { Val = 1 },
                        NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.Decimal },
                        LevelText = new LevelText() { Val = "%CurrentLevel." },
                        LevelJustification = new LevelJustification() { Val = LevelJustificationValues.Left },
                        PreviousParagraphProperties = new PreviousParagraphProperties() {
                            Indentation = new Indentation() { Left = "720", Hanging = "360" }
                        }
                    };
                    break;
                case WordListLevelKind.LowerLetter:
                    _level = new Level() {
                        LevelIndex = 0,
                        TemplateCode = "",
                        StartNumberingValue = new StartNumberingValue() { Val = 1 },
                        NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.LowerLetter },
                        LevelText = new LevelText() { Val = "%CurrentLevel" },
                        LevelJustification = new LevelJustification() { Val = LevelJustificationValues.Left },
                        PreviousParagraphProperties = new PreviousParagraphProperties() {
                            Indentation = new Indentation() { Left = "720", Hanging = "360" }
                        }
                    };
                    break;
                case WordListLevelKind.LowerLetterBracket:
                    _level = new Level() {
                        LevelIndex = 0,
                        TemplateCode = "",
                        StartNumberingValue = new StartNumberingValue() { Val = 1 },
                        NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.LowerLetter },
                        LevelText = new LevelText() { Val = "%CurrentLevel)" },
                        LevelJustification = new LevelJustification() { Val = LevelJustificationValues.Left },
                        PreviousParagraphProperties = new PreviousParagraphProperties() {
                            Indentation = new Indentation() { Left = "720", Hanging = "360" }
                        }
                    };
                    break;
                case WordListLevelKind.LowerLetterDot:
                    _level = new Level() {
                        LevelIndex = 0,
                        TemplateCode = "",
                        StartNumberingValue = new StartNumberingValue() { Val = 1 },
                        NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.LowerLetter },
                        LevelText = new LevelText() { Val = "%CurrentLevel." },
                        LevelJustification = new LevelJustification() { Val = LevelJustificationValues.Left },
                        PreviousParagraphProperties = new PreviousParagraphProperties() {
                            Indentation = new Indentation() { Left = "720", Hanging = "360" }
                        }
                    };
                    break;
                case WordListLevelKind.UpperLetter:
                    _level = new Level() {
                        LevelIndex = 0,
                        TemplateCode = "",
                        StartNumberingValue = new StartNumberingValue() { Val = 1 },
                        NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.UpperLetter },
                        LevelText = new LevelText() { Val = "%CurrentLevel" },
                        LevelJustification = new LevelJustification() { Val = LevelJustificationValues.Left },
                        PreviousParagraphProperties = new PreviousParagraphProperties() {
                            Indentation = new Indentation() { Left = "720", Hanging = "360" }
                        }
                    };
                    break;
                case WordListLevelKind.UpperLetterBracket:
                    _level = new Level() {
                        LevelIndex = 0,
                        TemplateCode = "",
                        StartNumberingValue = new StartNumberingValue() { Val = 1 },
                        NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.UpperLetter },
                        LevelText = new LevelText() { Val = "%CurrentLevel)" },
                        LevelJustification = new LevelJustification() { Val = LevelJustificationValues.Left },
                        PreviousParagraphProperties = new PreviousParagraphProperties() {
                            Indentation = new Indentation() { Left = "720", Hanging = "360" }
                        }
                    };
                    break;
                case WordListLevelKind.UpperLetterDot:
                    _level = new Level() {
                        LevelIndex = 0,
                        TemplateCode = "",
                        StartNumberingValue = new StartNumberingValue() { Val = 1 },
                        NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.UpperLetter },
                        LevelText = new LevelText() { Val = "%CurrentLevel." },
                        LevelJustification = new LevelJustification() { Val = LevelJustificationValues.Left },
                        PreviousParagraphProperties = new PreviousParagraphProperties() {
                            Indentation = new Indentation() { Left = "720", Hanging = "360" }
                        }
                    };
                    break;
                case WordListLevelKind.LowerRoman:
                    _level = new Level() {
                        LevelIndex = 0,
                        TemplateCode = "",
                        StartNumberingValue = new StartNumberingValue() { Val = 1 },
                        NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.LowerRoman },
                        LevelText = new LevelText() { Val = "%CurrentLevel" },
                        LevelJustification = new LevelJustification() { Val = LevelJustificationValues.Left },
                        PreviousParagraphProperties = new PreviousParagraphProperties() {
                            Indentation = new Indentation() { Left = "720", Hanging = "360" }
                        },
                    };
                    break;
                case WordListLevelKind.LowerRomanBracket:
                    _level = new Level() {
                        LevelIndex = 0,
                        TemplateCode = "",
                        StartNumberingValue = new StartNumberingValue() { Val = 1 },
                        NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.LowerRoman },
                        LevelText = new LevelText() { Val = "%CurrentLevel)" },
                        LevelJustification = new LevelJustification() { Val = LevelJustificationValues.Left },
                        PreviousParagraphProperties = new PreviousParagraphProperties() {
                            Indentation = new Indentation() { Left = "720", Hanging = "360" }
                        }
                    };
                    break;
                case WordListLevelKind.LowerRomanDot:
                    _level = new Level() {
                        LevelIndex = 0,
                        TemplateCode = "",
                        StartNumberingValue = new StartNumberingValue() { Val = 1 },
                        NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.LowerRoman },
                        LevelText = new LevelText() { Val = "%CurrentLevel." },
                        LevelJustification = new LevelJustification() { Val = LevelJustificationValues.Left },
                        PreviousParagraphProperties = new PreviousParagraphProperties() {
                            Indentation = new Indentation() { Left = "720", Hanging = "360" }
                        }
                    };
                    break;
                case WordListLevelKind.UpperRoman:
                    _level = new Level() {
                        LevelIndex = 0,
                        TemplateCode = "",
                        StartNumberingValue = new StartNumberingValue() { Val = 1 },
                        NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.UpperRoman },
                        LevelText = new LevelText() { Val = "%CurrentLevel" },
                        LevelJustification = new LevelJustification() { Val = LevelJustificationValues.Left },
                        PreviousParagraphProperties = new PreviousParagraphProperties() {
                            Indentation = new Indentation() { Left = "720", Hanging = "360" }
                        }
                    };
                    break;
                case WordListLevelKind.UpperRomanBracket:
                    _level = new Level() {
                        LevelIndex = 0,
                        TemplateCode = "",
                        StartNumberingValue = new StartNumberingValue() { Val = 1 },
                        NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.UpperRoman },
                        LevelText = new LevelText() { Val = "%CurrentLevel)" },
                        LevelJustification = new LevelJustification() { Val = LevelJustificationValues.Left },
                        PreviousParagraphProperties = new PreviousParagraphProperties() {
                            Indentation = new Indentation() { Left = "720", Hanging = "360" }
                        }
                    };
                    break;
                case WordListLevelKind.UpperRomanDot:
                    _level = new Level() {
                        LevelIndex = 0,
                        TemplateCode = "",
                        StartNumberingValue = new StartNumberingValue() { Val = 1 },
                        NumberingFormat = new NumberingFormat() { Val = NumberFormatValues.UpperRoman },
                        LevelText = new LevelText() { Val = "%CurrentLevel." },
                        LevelJustification = new LevelJustification() { Val = LevelJustificationValues.Left },
                        PreviousParagraphProperties = new PreviousParagraphProperties() {
                            Indentation = new Indentation() { Left = "720", Hanging = "360" }
                        }
                    };
                    break;
                case WordListLevelKind.None:
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(simplifiedListNumbers), simplifiedListNumbers, null);
                    break;
            }
        }
    }
}
