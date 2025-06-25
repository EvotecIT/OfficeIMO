using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using V = DocumentFormat.OpenXml.Vml;

namespace OfficeIMO.Word {
    public partial class WordList {
        /// <summary>
        /// Adds a new empty item to the list.
        /// </summary>
        /// <param name="wordParagraph">The paragraph after which the item should be inserted. If <c>null</c> the item is appended at the default position.</param>
        /// <param name="level">The zero-based list level for the new item.</param>
        /// <returns>The newly created <see cref="WordParagraph"/> representing the list item.</returns>
        public WordParagraph AddItem(WordParagraph wordParagraph, int level = 0) {
            return AddItem(null, level, wordParagraph);
        }

        /// <summary>
        /// Adds a new item to the list and optionally sets its text.
        /// </summary>
        /// <param name="text">Text to assign to the new list item. Pass <c>null</c> to insert an empty item.</param>
        /// <param name="level">The zero-based list level for the item.</param>
        /// <param name="wordParagraph">The paragraph after which the item is inserted. When <c>null</c> the item is appended in the default position.</param>
        /// <returns>The <see cref="WordParagraph"/> representing the created list item.</returns>
        public WordParagraph AddItem(string text, int level = 0, WordParagraph wordParagraph = null) {
            var paragraph = new Paragraph();
            var run = new Run();
            run.Append(new RunProperties());
            run.Append(new Text { Space = SpaceProcessingModeValues.Preserve });

            var levelIndex = level;
            if (_isToc || IsToc) {
                if (levelIndex < 0) levelIndex = 0;
                else if (levelIndex > 8) levelIndex = 8;
            }

            var paragraphProperties = new ParagraphProperties();
            paragraphProperties.Append(new ParagraphStyleId { Val = "ListParagraph" });
            paragraphProperties.Append(
                new NumberingProperties(
                    new NumberingLevelReference { Val = levelIndex },
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
            newParagraph._list = this;
            _listItems.Add(newParagraph);
            if (text != null) {
                newParagraph.Text = text;
            }

            if (_isToc || IsToc) {
                newParagraph.Style = WordParagraphStyle.GetStyle(levelIndex);
            }

            return newParagraph;
        }

        internal void RemoveItem(WordParagraph paragraph) {
            _listItems.Remove(paragraph);
        }

        /// <summary>
        /// Removes the list from the document including all list items.
        /// </summary>
        /// <remarks>If no other list references the underlying numbering definition, it is deleted from the document.</remarks>
        public void Remove() {
            foreach (var listItem in _listItems.ToList()) {
                listItem.Remove();
            }
            _listItems.Clear();

            var numberingPart = _document._wordprocessingDocument.MainDocumentPart.NumberingDefinitionsPart;
            var numbering = numberingPart?.Numbering;
            if (numbering != null) {
                bool stillReferenced = _document.EnumerateAllParagraphs()
                    .Any(p => p.IsListItem && p._listNumberId == _numberId);

                if (!stillReferenced) {
                    var abstractNum = numbering.Elements<AbstractNum>().FirstOrDefault(a => a.AbstractNumberId.Value == _abstractId);
                    abstractNum?.Remove();

                    var numberingInstance = numbering.Elements<NumberingInstance>().FirstOrDefault(n => n.NumberID.Value == _numberId);
                    numberingInstance?.Remove();

                    if (!numbering.Elements<AbstractNum>().Any() &&
                        !numbering.Elements<NumberingInstance>().Any() &&
                        !numbering.Elements<NumberingPictureBullet>().Any()) {
                        _document._wordprocessingDocument.MainDocumentPart.DeletePart(numberingPart);
                    }
                }
            }
        }

        /// <summary>
        /// Creates a copy of this list after the last item.
        /// </summary>
        /// <returns>The newly created <see cref="WordList"/>.</returns>
        /// <exception cref="InvalidOperationException">Thrown when the list contains no items.</exception>
        public WordList Clone() {
            var reference = ListItems.LastOrDefault()?._paragraph;
            if (reference == null) {
                throw new InvalidOperationException("Cannot clone an empty list.");
            }
            return Clone(reference, true);
        }

        /// <summary>
        /// Clones this list and inserts it relative to the provided paragraph.
        /// </summary>
        /// <param name="paragraph">The reference paragraph used for insertion.</param>
        /// <param name="after">If set to <c>true</c> the clone is inserted after the reference paragraph; otherwise it is inserted before.</param>
        /// <returns>The cloned <see cref="WordList"/>.</returns>
        public WordList Clone(WordParagraph paragraph, bool after = true) {
            return Clone(paragraph._paragraph, after);
        }

        /// <summary>
        /// Merges another list into this list.
        /// </summary>
        /// <param name="documentList">The list whose items should be moved into this instance.</param>
        /// <remarks>The source list is removed after its items are transferred.</remarks>
        public void Merge(WordList documentList) {
            foreach (var item in documentList.ListItems.ToList()) {
                var numberingProperties = item._paragraphProperties.NumberingProperties;
                if (numberingProperties != null && numberingProperties.NumberingId != null) {
                    numberingProperties.NumberingId.Val = this._numberId;
                }
                documentList._listItems.Remove(item);
                item._list = this;
                _listItems.Add(item);
            }
            documentList.Remove();
        }

        /// <summary>
        /// Creates a bulleted list using a custom symbol and formatting.
        /// </summary>
        /// <param name="document">The parent document.</param>
        /// <param name="symbol">The character to use as the bullet symbol.</param>
        /// <param name="fontName">Font name used for the bullet symbol.</param>
        /// <param name="colorHex">Optional hexadecimal color for the symbol.</param>
        /// <param name="fontSize">Optional font size for the symbol in points.</param>
        /// <returns>The created <see cref="WordList"/>.</returns>
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

        /// <summary>
        /// Creates a custom bulleted list using a predefined <see cref="WordBulletSymbol"/>.
        /// </summary>
        /// <param name="document">The parent document.</param>
        /// <param name="symbol">Predefined bullet symbol to use.</param>
        /// <param name="fontName">Font name used for the symbol.</param>
        /// <param name="color">Optional <see cref="SixLabors.ImageSharp.Color"/> for the symbol.</param>
        /// <param name="colorHex">Optional color specified as hex string. Ignored when <paramref name="color"/> is provided.</param>
        /// <param name="fontSize">Optional font size in points.</param>
        /// <returns>The created <see cref="WordList"/>.</returns>
        public static WordList AddCustomBulletList(WordDocument document, WordBulletSymbol symbol, string fontName, SixLabors.ImageSharp.Color? color = null, string colorHex = null, int? fontSize = null) {
            string finalColor = color?.ToHexColor() ?? colorHex;
            return AddCustomBulletList(document, (char)symbol, fontName, finalColor, fontSize);
        }

        /// <summary>
        /// Creates a bulleted list using one of the predefined <see cref="WordListLevelKind"/> values.
        /// </summary>
        /// <param name="document">The parent document.</param>
        /// <param name="kind">Determines which bullet symbol to use.</param>
        /// <param name="fontName">Font name used for the symbol.</param>
        /// <param name="color">Optional symbol color.</param>
        /// <param name="colorHex">Optional color specified as hex string. Ignored when <paramref name="color"/> is provided.</param>
        /// <param name="fontSize">Optional font size in points.</param>
        /// <returns>The created <see cref="WordList"/>.</returns>
        public static WordList AddCustomBulletList(WordDocument document, WordListLevelKind kind, string fontName, SixLabors.ImageSharp.Color? color = null, string colorHex = null, int? fontSize = null) {
            char symbol = GetBulletSymbol(kind);
            string finalColor = color?.ToHexColor() ?? colorHex;
            return AddCustomBulletList(document, symbol, fontName, finalColor, fontSize);
        }

        /// <summary>
        /// Creates an empty custom list without any predefined levels.
        /// </summary>
        /// <param name="document">The parent document.</param>
        /// <returns>The created <see cref="WordList"/>.</returns>
        public static WordList AddCustomList(WordDocument document) {
            if (document == null) throw new ArgumentNullException(nameof(document));

            var list = document.AddList(WordListStyle.Custom);
            list.Numbering.RemoveAllLevels();
            return list;
        }

        /// <summary>
        /// Creates a bulleted list using the specified image as the bullet symbol.
        /// </summary>
        /// <param name="document">The parent document.</param>
        /// <param name="imageStream">Stream containing the bullet image.</param>
        /// <param name="fileName">Name of the image file.</param>
        /// <returns>The created <see cref="WordList"/>.</returns>
        public static WordList AddPictureBulletList(WordDocument document, Stream imageStream, string fileName) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (imageStream == null) throw new ArgumentNullException(nameof(imageStream));

            var numberingPart = document._wordprocessingDocument.MainDocumentPart.NumberingDefinitionsPart;
            if (numberingPart == null) {
                numberingPart = document._wordprocessingDocument.MainDocumentPart.AddNewPart<NumberingDefinitionsPart>();
                numberingPart.Numbering = new Numbering();
            }
            var numbering = numberingPart.Numbering;
            EnsureW15Namespace(numbering);

            var characteristics = Helpers.GetImageCharacteristics(imageStream, fileName);
            var imagePart = numberingPart.AddImagePart(characteristics.Type.ToOpenXmlImagePartType());
            imageStream.Position = 0;
            imagePart.FeedData(imageStream);
            var relId = numberingPart.GetIdOfPart(imagePart);

            int picId = numbering.Elements<NumberingPictureBullet>()
                .Select(p => (int)p.NumberingPictureBulletId.Value)
                .DefaultIfEmpty(0)
                .Max() + 1;

            var numPicBullet = new NumberingPictureBullet { NumberingPictureBulletId = picId };
            var pict = new PictureBulletBase();
            var shape = new V.Shape { Style = "width:12pt;height:12pt" };
            var imageData = new V.ImageData { RelationshipId = relId };
            shape.Append(imageData);
            pict.Append(shape);
            numPicBullet.Append(pict);
            numbering.Append(numPicBullet);

            var list = document.AddList(WordListStyle.Custom);
            list.Numbering.RemoveAllLevels();
            list.ReplaceAbstractNum(WordListStyles.CreatePictureBulletStyle(picId));
            return list;
        }

        /// <summary>
        /// Creates a bulleted list using an image file as the bullet symbol.
        /// </summary>
        /// <param name="document">The parent document.</param>
        /// <param name="imagePath">Path to the image file.</param>
        /// <returns>The created <see cref="WordList"/>.</returns>
        public static WordList AddPictureBulletList(WordDocument document, string imagePath) {
            using var stream = new FileStream(imagePath, FileMode.Open, FileAccess.Read, FileShare.Read);
            return AddPictureBulletList(document, stream, System.IO.Path.GetFileName(imagePath));
        }

        /// <summary>
        /// Adds a bullet level to the list using a custom symbol.
        /// </summary>
        /// <param name="levelIndex">The one-based level index to add.</param>
        /// <param name="symbol">Bullet character to use.</param>
        /// <param name="fontName">Font name for the symbol.</param>
        /// <param name="colorHex">Color of the symbol in hex format.</param>
        /// <param name="fontSize">Optional symbol size in points.</param>
        /// <returns>The current <see cref="WordList"/> instance.</returns>
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

            if (currentCount == 0 && levelIndex > 1) {
                var placeholder = CreateBulletLevel(symbol, fontName, colorHex, fontSize);
                while (this.Numbering.Levels.Count < levelIndex - 1) {
                    var clone = (Level)placeholder.CloneNode(true);
                    this.Numbering.AddLevel(clone);
                }
            }

            var newLevel = CreateBulletLevel(symbol, fontName, colorHex, fontSize);
            this.Numbering.AddLevel(newLevel);
            return this;
        }

        /// <summary>
        /// Adds a bullet level using a predefined <see cref="WordBulletSymbol"/>.
        /// </summary>
        /// <param name="levelIndex">The one-based level index to add.</param>
        /// <param name="symbol">Predefined symbol.</param>
        /// <param name="fontName">Font name for the symbol.</param>
        /// <param name="color">Optional color for the symbol.</param>
        /// <param name="colorHex">Optional color as hex when <paramref name="color"/> is not provided.</param>
        /// <param name="fontSize">Optional symbol size in points.</param>
        /// <returns>The current <see cref="WordList"/> instance.</returns>
        public WordList AddListLevel(int levelIndex, WordBulletSymbol symbol, string fontName, SixLabors.ImageSharp.Color? color = null, string colorHex = null, int? fontSize = null) {
            string finalColor = color?.ToHexColor() ?? colorHex;
            return AddListLevel(levelIndex, (char)symbol, fontName, finalColor, fontSize);
        }

        /// <summary>
        /// Adds a bullet level using one of the predefined <see cref="WordListLevelKind"/> values.
        /// </summary>
        /// <param name="levelIndex">The one-based level index to add.</param>
        /// <param name="kind">Specifies the symbol kind.</param>
        /// <param name="fontName">Font name for the symbol.</param>
        /// <param name="color">Optional color for the symbol.</param>
        /// <param name="colorHex">Optional color as hex when <paramref name="color"/> is not provided.</param>
        /// <param name="fontSize">Optional symbol size in points.</param>
        /// <returns>The current <see cref="WordList"/> instance.</returns>
        public WordList AddListLevel(int levelIndex, WordListLevelKind kind, string fontName, SixLabors.ImageSharp.Color? color = null, string colorHex = null, int? fontSize = null) {
            char symbol = GetBulletSymbol(kind);
            string finalColor = color?.ToHexColor() ?? colorHex;
            return AddListLevel(levelIndex, symbol, fontName, finalColor, fontSize);
        }

        /// <summary>
        /// Converts this list to a numbered style while preserving existing list items.
        /// </summary>
        public void ConvertToNumbered() {
            ReplaceAbstractNum(WordListStyles.GetStyle(WordListStyle.Headings111));
        }

        /// <summary>
        /// Converts this list to a bulleted style while preserving existing list items.
        /// </summary>
        public void ConvertToBulleted() {
            ReplaceAbstractNum(WordListStyles.GetStyle(WordListStyle.Bulleted));
        }
    }
}
