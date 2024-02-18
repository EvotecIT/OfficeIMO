using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;
using SixLabors.ImageSharp;

namespace OfficeIMO.Word {
    public class WordListLevel {
        public WordListLevel(Level level) {
            _level = level;
        }

        public Level _level { get; set; }

        public int StartNumberingValue {
            get {
                return _level.Descendants<StartNumberingValue>().First().Val;
            }
            set {
                _level.Descendants<StartNumberingValue>().First().Val = value;
            }
        }
        public int IndentationLeft {
            get {
                return int.Parse(_level.Descendants<Indentation>().First().Left);
            }
            set {
                _level.Descendants<Indentation>().First().Left = value.ToString();
            }
        }

        public WordListLevel(NumberFormatValues numberFormat) {
            StartNumberingValue startNumberingValue1 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat1 = new NumberingFormat() { Val = numberFormat };
            LevelText levelText1 = new LevelText() { Val = "·" };
            LevelJustification levelJustification1 = new LevelJustification() { Val = LevelJustificationValues.Left };
            PreviousParagraphProperties previousParagraphProperties1 = new PreviousParagraphProperties() {
                //Indentation = new Indentation() { Left = "720", Hanging = "360" }
            };

            // level index doesn't matter here, it will be set when adding to the numbering based on current levels
            _level = new Level() { LevelIndex = 0, TemplateCode = "" };
            _level.Append(startNumberingValue1);
            _level.Append(numberingFormat1);
            _level.Append(levelText1);
            _level.Append(levelJustification1);
            _level.Append(previousParagraphProperties1);

            if (numberFormat == NumberFormatValues.Bullet) {
                NumberingSymbolRunProperties numberingSymbolRunProperties1 = new NumberingSymbolRunProperties() {
                    RunFonts = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol" }
                };
                _level.Append(numberingSymbolRunProperties1);
            }
        }

        public void Remove() {
            _level.Remove();
        }
    }

    public enum SimplifiedNumbers {
        None,
        BulletLevel0,
        BulletLevel1,
        BulletLevel2,
        Decimal,
        LowerLetter,
        UpperLetter,
        LowerRoman,
        UpperRoman
    }
}
