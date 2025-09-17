using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Defines the numbering or bullet style used for a particular list level.
    /// </summary>
public enum WordListLevelKind {
        /// <summary>
        /// No bullet or numbering.
        /// </summary>
        None,
        /// <summary>
        /// Standard bullet symbol.
        /// </summary>
        Bullet,
        /// <summary>
        /// Square bullet symbol.
        /// </summary>
        BulletSquareSymbol,
        /// <summary>
        /// Black circle bullet symbol.
        /// </summary>
        BulletBlackCircle,
        /// <summary>
        /// Diamond bullet symbol.
        /// </summary>
        BulletDiamondSymbol,
        /// <summary>
        /// Arrow bullet symbol.
        /// </summary>
        BulletArrowSymbol,
        /// <summary>
        /// Solid round bullet symbol.
        /// </summary>
        BulletSolidRound,
        /// <summary>
        /// Open circle bullet symbol.
        /// </summary>
        BulletOpenCircle,
        /// <summary>
        /// Square bullet style.
        /// </summary>
        BulletSquare,
        /// <summary>
        /// The BulletSquare2, not the same as BulletSquare, and not even used by Word
        /// </summary>
        BulletSquare2,
        /// <summary>
        /// Checkmark bullet symbol.
        /// </summary>
        BulletCheckmark,
        /// <summary>
        /// Clubs bullet symbol.
        /// </summary>
        BulletClubs,
        /// <summary>
        /// Diamond shape bullet.
        /// </summary>
        BulletDiamond,
        /// <summary>
        /// Arrow shape bullet.
        /// </summary>
        BulletArrow,

        /// <summary>
        /// Decimal numbers.
        /// </summary>
        Decimal,
        /// <summary>
        /// Decimal numbers followed by a bracket.
        /// </summary>
        DecimalBracket,
        /// <summary>
        /// Decimal numbers followed by a dot.
        /// </summary>
        DecimalDot,
        /// <summary>
        /// Lowercase letters.
        /// </summary>
        LowerLetter,
        /// <summary>
        /// Lowercase letters followed by a bracket.
        /// </summary>
        LowerLetterBracket,
        /// <summary>
        /// Lowercase letters followed by a dot.
        /// </summary>
        LowerLetterDot,
        /// <summary>
        /// Uppercase letters.
        /// </summary>
        UpperLetter,
        /// <summary>
        /// Uppercase letters followed by a bracket.
        /// </summary>
        UpperLetterBracket,
        /// <summary>
        /// Uppercase letters followed by a dot.
        /// </summary>
        UpperLetterDot,
        /// <summary>
        /// Lowercase Roman numerals.
        /// </summary>
        LowerRoman,
        /// <summary>
        /// Lowercase Roman numerals followed by a bracket.
        /// </summary>
        LowerRomanBracket,
        /// <summary>
        /// Lowercase Roman numerals followed by a dot.
        /// </summary>
        LowerRomanDot,
        /// <summary>
        /// Uppercase Roman numerals.
        /// </summary>
        UpperRoman,
        /// <summary>
        /// Uppercase Roman numerals followed by a bracket.
        /// </summary>
        UpperRomanBracket,
        /// <summary>
        /// Uppercase Roman numerals followed by a dot.
        /// </summary>
        UpperRomanDot
    }
    /// <summary>
    /// Represents a single level within a list and provides access to numbering and indentation settings.
    /// </summary>
    public class WordListLevel {
        /// <summary>
        /// Initializes a new instance of the <see cref="WordListLevel"/> class
        /// based on an existing Open XML <see cref="Level"/> element.
        /// </summary>
        /// <param name="level">The underlying Open XML list level element.</param>
        public WordListLevel(Level level) {
            _level = level ?? throw new ArgumentNullException(nameof(level));
        }

        /// <summary>
        /// Gets or sets the underlying Open XML list level element that backs
        /// this instance.
        /// </summary>
        public Level _level { get; set; } = null!;

        private StartNumberingValue GetStartNumberingValueElement() {
            var element = _level.Descendants<StartNumberingValue>().FirstOrDefault();
            if (element == null) {
                element = new StartNumberingValue { Val = 1 };
                _level.Append(element);
            } else if (element.Val == null || !element.Val.HasValue) {
                element.Val = 1;
            }

            return element;
        }

        private Indentation GetIndentationElement() {
            var indentation = _level.Descendants<Indentation>().FirstOrDefault();
            if (indentation != null) {
                return indentation;
            }

            var paragraphProperties = _level.GetFirstChild<PreviousParagraphProperties>();
            if (paragraphProperties == null) {
                paragraphProperties = new PreviousParagraphProperties();
                _level.Append(paragraphProperties);
            }

            indentation = new Indentation { Left = "0", Hanging = "0" };
            paragraphProperties.Append(indentation);
            return indentation;
        }

        private LevelText GetLevelTextElement() {
            var levelText = _level.GetFirstChild<LevelText>();
            if (levelText == null) {
                levelText = new LevelText { Val = string.Empty };
                _level.Append(levelText);
            } else if (levelText.Val == null) {
                levelText.Val = string.Empty;
            }

            return levelText;
        }

        private LevelJustification GetLevelJustificationElement() {
            var justification = _level.GetFirstChild<LevelJustification>();
            if (justification == null) {
                justification = new LevelJustification { Val = LevelJustificationValues.Left };
                _level.Append(justification);
            } else if (justification.Val == null || !justification.Val.HasValue) {
                justification.Val = LevelJustificationValues.Left;
            }

            return justification;
        }

        /// <summary>
        /// Gets or sets the start numbering value.
        /// </summary>
        public int StartNumberingValue {
            get {
                var element = GetStartNumberingValueElement();
                return element.Val?.Value ?? 0;
            }
            set {
                var element = GetStartNumberingValueElement();
                element.Val = value;
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
                var indentation = GetIndentationElement();
                return int.TryParse(indentation.Left, out int value) ? value : 0;
            }
            set {
                var indentation = GetIndentationElement();
                indentation.Left = value.ToString();
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
                var indentation = GetIndentationElement();
                return int.TryParse(indentation.Hanging, out int value) ? value : 0;
            }
            set {
                var indentation = GetIndentationElement();
                indentation.Hanging = value.ToString();
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
                var element = GetLevelTextElement();
                return element.Val?.Value ?? string.Empty;
            }
            set {
                var element = GetLevelTextElement();
                element.Val = value ?? string.Empty;
            }
        }

        /// <summary>
        /// Gets or sets the level justification.
        /// </summary>
        public LevelJustificationValues LevelJustification {
            get {
                var element = GetLevelJustificationElement();
                return element.Val?.Value ?? LevelJustificationValues.Left;
            }
            set {
                var element = GetLevelJustificationElement();
                element.Val = value;
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
            }
        }
    }
}
