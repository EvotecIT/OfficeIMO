using System;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public enum SimplifiedListNumbers {
        None,
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

        /// <summary>
        /// Gets or sets the _level.
        /// </summary>
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
        public WordListLevel(SimplifiedListNumbers simplifiedListNumbers) {
            switch (simplifiedListNumbers) {
                case SimplifiedListNumbers.BulletSolidRound:
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
                case SimplifiedListNumbers.BulletOpenCircle:
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
                case SimplifiedListNumbers.BulletSquare2:
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
                case SimplifiedListNumbers.BulletSquare:
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
                case SimplifiedListNumbers.BulletClubs:
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
                case SimplifiedListNumbers.BulletArrow:
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
                case SimplifiedListNumbers.BulletDiamond:
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
                case SimplifiedListNumbers.BulletCheckmark:
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
                case SimplifiedListNumbers.Decimal:
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
                case SimplifiedListNumbers.DecimalBracket:
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
                case SimplifiedListNumbers.DecimalDot:
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
                case SimplifiedListNumbers.LowerLetter:
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
                case SimplifiedListNumbers.LowerLetterBracket:
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
                case SimplifiedListNumbers.LowerLetterDot:
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
                case SimplifiedListNumbers.UpperLetter:
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
                case SimplifiedListNumbers.UpperLetterBracket:
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
                case SimplifiedListNumbers.UpperLetterDot:
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
                case SimplifiedListNumbers.LowerRoman:
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
                case SimplifiedListNumbers.LowerRomanBracket:
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
                case SimplifiedListNumbers.LowerRomanDot:
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
                case SimplifiedListNumbers.UpperRoman:
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
                case SimplifiedListNumbers.UpperRomanBracket:
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
                case SimplifiedListNumbers.UpperRomanDot:
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
                case SimplifiedListNumbers.None:
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(simplifiedListNumbers), simplifiedListNumbers, null);
                    break;
            }
        }
    }
}
