using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Provides public operations for modifying Word lists.
    /// </summary>
    public partial class WordList {
        public WordParagraph AddItem(WordParagraph wordParagraph, int level = 0) {
            return AddItem(null, level, wordParagraph);
        }

        public WordParagraph AddItem(string text, int level = 0, WordParagraph wordParagraph = null) {
            var paragraph = new Paragraph();
            var run = new Run();
            run.Append(new RunProperties());
            run.Append(new Text { Space = SpaceProcessingModeValues.Preserve });

            var paragraphProperties = new ParagraphProperties();
            paragraphProperties.Append(new ParagraphStyleId { Val = "ListParagraph" });
            paragraphProperties.Append(
                new NumberingProperties(
                    new NumberingLevelReference { Val = level },
                    new NumberingId { Val = _numberId }
                ));
            paragraph.Append(paragraphProperties);
            paragraph.Append(run);

            if (wordParagraph != null) {
                wordParagraph._paragraph.InsertAfterSelf(paragraph);
            } else if (this.ListItems.Count == 0 && _wordParagraph != null) {
                _wordParagraph._paragraph.InsertAfterSelf(paragraph);
            } else if (_isToc || IsToc) {
                _wordprocessingDocument.MainDocumentPart!.Document.Body!.AppendChild(paragraph);
            } else if (_headerFooter != null) {
                if (_headerFooter._header != null) {
                    _headerFooter._header.Append(paragraph);
                } else if (_headerFooter._footer != null) {
                    _headerFooter._footer.Append(paragraph);
                }
            } else if (_wordParagraph != null && _wordParagraph._paragraph.Parent is TableCell) {
                var parent = _wordParagraph._paragraph.Parent;
                if (this.ListItems.Count > 0) {
                    var lastItem = this.ListItems.Last();
                    lastItem._paragraph.InsertAfterSelf(paragraph);
                } else {
                    parent.Append(paragraph);
                }
            } else {
                _wordprocessingDocument.MainDocumentPart!.Document.Body!.AppendChild(paragraph);
            }

            var newParagraph = new WordParagraph(_document, paragraph, run);
            if (text != null) {
                newParagraph.Text = text;
            }

            if (_isToc || IsToc) {
                newParagraph.Style = WordParagraphStyle.GetStyle(level);
            }

            return newParagraph;
        }

        public void Remove() {
            var numberingPart = _document._wordprocessingDocument.MainDocumentPart.NumberingDefinitionsPart;
            var numbering = numberingPart?.Numbering;
            if (numbering != null) {
                var abstractNum = numbering.Elements<AbstractNum>().FirstOrDefault(a => a.AbstractNumberId.Value == _abstractId);
                abstractNum?.Remove();

                var numberingInstance = numbering.Elements<NumberingInstance>().FirstOrDefault(n => n.NumberID.Value == _numberId);
                numberingInstance?.Remove();

                if (!numbering.ChildElements.OfType<AbstractNum>().Any() &&
                    !numbering.ChildElements.OfType<NumberingInstance>().Any() &&
                    !numbering.ChildElements.OfType<NumberingPictureBullet>().Any()) {
                    _document._wordprocessingDocument.MainDocumentPart.DeletePart(numberingPart);
                }
            }

            foreach (var listItem in ListItems.ToList()) {
                listItem.Remove();
            }
        }

        public WordList Clone() {
            var reference = ListItems.LastOrDefault()?._paragraph;
            if (reference == null) {
                throw new InvalidOperationException("Cannot clone an empty list.");
            }
            return Clone(reference, true);
        }

        public WordList Clone(WordParagraph paragraph, bool after = true) {
            return Clone(paragraph._paragraph, after);
        }

        public void Merge(WordList documentList) {
            foreach (var item in documentList.ListItems) {
                var numberingProperties = item._paragraphProperties.NumberingProperties;
                if (numberingProperties != null && numberingProperties.NumberingId != null) {
                    numberingProperties.NumberingId.Val = this._numberId;
                }
            }
            documentList.Remove();
        }

        public static WordList AddCustomBulletList(WordDocument document, char symbol, string fontName, string colorHex, int? fontSize = null) {
            if (document == null) throw new ArgumentNullException(nameof(document));

            var list = document.AddList(WordListStyle.Custom);

            var level = new Level();
            level.Append(new StartNumberingValue() { Val = 1 });
            level.Append(new NumberingFormat() { Val = NumberFormatValues.Bullet });
            level.Append(new LevelText() { Val = symbol.ToString() });
            level.Append(new LevelJustification() { Val = LevelJustificationValues.Left });

            var prevProps = new PreviousParagraphProperties();
            prevProps.Append(new Indentation() { Left = "720", Hanging = "360" });
            level.Append(prevProps);

            var symbolProps = new NumberingSymbolRunProperties();
            if (!string.IsNullOrEmpty(fontName)) {
                symbolProps.Append(new RunFonts { Ascii = fontName, HighAnsi = fontName });
            }
            if (!string.IsNullOrEmpty(colorHex)) {
                symbolProps.Append(new DocumentFormat.OpenXml.Wordprocessing.Color { Val = colorHex.Replace("#", "").ToLowerInvariant() });
            }
            if (fontSize.HasValue) {
                var size = (fontSize.Value * 2).ToString();
                symbolProps.Append(new FontSize { Val = size });
                symbolProps.Append(new FontSizeComplexScript { Val = size });
            }
            level.Append(symbolProps);

            list.Numbering.AddLevel(level);

            return list;
        }

        public static WordList AddCustomBulletList(WordDocument document, WordBulletSymbol symbol, string fontName, SixLabors.ImageSharp.Color? color = null, string colorHex = null, int? fontSize = null) {
            string finalColor = color?.ToHexColor() ?? colorHex;
            return AddCustomBulletList(document, (char)symbol, fontName, finalColor, fontSize);
        }

        public static WordList AddCustomList(WordDocument document) {
            if (document == null) throw new ArgumentNullException(nameof(document));

            var list = document.AddList(WordListStyle.Custom);
            list.Numbering.RemoveAllLevels();
            return list;
        }

        public WordList AddListLevel(int levelIndex, char symbol, string fontName, string colorHex, int? fontSize = null) {
            if (levelIndex < 1) throw new ArgumentOutOfRangeException(nameof(levelIndex));

            var currentCount = this.Numbering.Levels.Count;
            if (currentCount > 0 && levelIndex > currentCount + 1) {
                var last = this.Numbering.Levels.Last()._level;
                while (this.Numbering.Levels.Count < levelIndex - 1) {
                    var clone = (Level)last.CloneNode(true);
                    this.Numbering.AddLevel(clone);
                }
            }

            if (levelIndex > 1 && this.Numbering.Levels.Count == 0) {
                levelIndex = 1;
            }

            var newLevel = CreateBulletLevel(symbol, fontName, colorHex, fontSize);
            this.Numbering.AddLevel(newLevel);
            return this;
        }

        public WordList AddListLevel(int levelIndex, WordBulletSymbol symbol, string fontName, SixLabors.ImageSharp.Color? color = null, string colorHex = null, int? fontSize = null) {
            string finalColor = color?.ToHexColor() ?? colorHex;
            return AddListLevel(levelIndex, (char)symbol, fontName, finalColor, fontSize);
        }
    }
}
