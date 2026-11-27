using System.Globalization;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.Word.OpenXml;


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
        private const int W = OfficeXmlNamespaces.WordprocessingML;

        /// <summary>
        /// Initializes a new instance of the <see cref="WordListLevel"/> class over an existing <c>w:lvl</c> element.
        /// </summary>
        /// <param name="level">The list level element.</param>
        internal WordListLevel(OfficeXmlNode level) {
            if (level.IsNone) throw new ArgumentNullException(nameof(level));
            Node = level;
        }

        /// <summary>
        /// The <c>w:lvl</c> element: in a small tree of its own until the level is added to a list, then the element
        /// the numbering part holds.
        /// </summary>
        internal OfficeXmlNode Node { get; set; }

        /// <summary>Gets the zero-based list level index.</summary>
        public int LevelIndex => OfficeXmlViewValues.ParseInt32(Node.Attribute(W, "ilvl")) ?? 0;

        /// <summary>Sets the zero-based list level index.</summary>
        internal void SetLevelIndex(int levelIndex) =>
            Node.SetAttribute(W, "ilvl", levelIndex.ToString(CultureInfo.InvariantCulture));

        /// <summary>The level's w:start, added with a value of 1 when missing.</summary>
        private OfficeXmlNode GetStartNumberingValueElement() {
            // Only the level's own w:start counts, not one inside extension markup.
            OfficeXmlNode existing = Node.Child(W, "start");
            return existing.IsNone ? AddValueElement("start", "1") : existing;
        }

        private OfficeXmlNode GetIndentationElement() {
            // The level's paragraph properties hold its indentation; one inside extension markup isn't the level's.
            OfficeXmlNode existing = Node.Child(W, "pPr").Child(W, "ind");
            if (!existing.IsNone) return existing;

            OfficeXmlNode paragraphProperties = Node.GetOrAddChild(W, "pPr", WLvlView.Order);
            OfficeXmlNode indentation = paragraphProperties.AddOrdered(W, "ind", WPPrView.Order);
            indentation.SetAttribute(W, "left", "0");
            indentation.SetAttribute(W, "hanging", "0");
            return indentation;
        }

        /// <summary>The child element with this name, added with <paramref name="defaultValue"/> when missing.</summary>
        private OfficeXmlNode GetValueElement(string localName, string defaultValue) {
            OfficeXmlNode element = Node.Child(W, localName);
            return element.IsNone ? AddValueElement(localName, defaultValue) : element;
        }

        private OfficeXmlNode AddValueElement(string localName, string value) {
            OfficeXmlNode element = Node.AddOrdered(W, localName, WLvlView.Order);
            element.SetAttribute(W, "val", value);
            return element;
        }

        /// <summary>
        /// Reads the value of a level setting, adding the element with <paramref name="defaultValue"/> when it is missing
        /// and writing the default over a value <paramref name="isValid"/> rejects, so the level always holds what it reports.
        /// </summary>
        private string ReadValue(OfficeXmlNode element, string defaultValue, Func<string, bool> isValid) {
            string? value = element.Attribute(W, "val");
            if (value != null && isValid(value)) return value;
            element.SetAttribute(W, "val", defaultValue);
            return defaultValue;
        }

        /// <summary>
        /// Reads a list level justification, accepting every ST_Jc token Word and other
        /// producers write. The logical <c>start</c>/<c>end</c> values (written by
        /// other office suites and strict producers) map to left/right for left-to-right lists;
        /// paragraph-only values such as <c>both</c> fall back to left, as Word does.
        /// </summary>
        internal static WordListLevelAlignment? ReadLevelJustification(string? token) {
            if (string.IsNullOrWhiteSpace(token)) {
                return null;
            }

            switch (token!.Trim()) {
                case "left":
                case "start":
                    return WordListLevelAlignment.Left;
                case "right":
                case "end":
                    return WordListLevelAlignment.Right;
                case "center":
                    return WordListLevelAlignment.Center;
                case "both":
                case "distribute":
                case "mediumKashida":
                case "highKashida":
                case "lowKashida":
                case "thaiDistribute":
                case "numTab":
                    return WordListLevelAlignment.Left;
                default:
                    return null;
            }
        }

        /// <summary>
        /// Gets or sets the start numbering value.
        /// </summary>
        public int StartNumberingValue {
            get => int.Parse(ReadValue(GetStartNumberingValueElement(), "1", static value => OfficeXmlViewValues.ParseInt32(value) != null).Trim(),
                NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture);
            set => GetStartNumberingValueElement().SetAttribute(W, "val", value.ToString(CultureInfo.InvariantCulture));
        }

        /// <summary>Gets or sets whether this level shows the numbers of the levels above it as Arabic numerals ("1.1" under "I"), Word's legal numbering (<c>w:isLgl</c>).</summary>
        public bool LegalNumbering {
            get {
                OfficeXmlNode legal = Node.Child(W, "isLgl");
                if (legal.IsNone) return false;
                string? value = legal.Attribute(W, "val")?.Trim();
                return value == null || value == "1" || value == "true" || value == "on";
            }
            set {
                OfficeXmlNode legal = Node.Child(W, "isLgl");
                if (!legal.IsNone) legal.Remove();
                if (value) Node.AddOrdered(W, "isLgl", WLvlView.Order);
            }
        }

        /// <summary>Gets or sets the numbering format of this list level.</summary>
        public WordNumberFormat NumberFormat {
            get {
                string? token = Node.Child(W, "numFmt").Attribute(W, "val");
                return token != null && WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out WordNumberFormat format) ? format : WordNumberFormat.Decimal;
            }
            set {
                // A new element, as the format replaces any other attributes the old one had. The token is resolved
                // first, so an undefined format leaves the level as it was.
                string token = value.ToOpenXmlToken();
                OfficeXmlNode existing = Node.Child(W, "numFmt");
                if (!existing.IsNone) existing.Remove();
                Node.AddOrdered(W, "numFmt", WLvlView.Order).SetAttribute(W, "val", token);
            }
        }

        /// <summary>
        /// Sets the starting number for this level.
        /// </summary>
        /// <param name="value">The starting number.</param>
        /// <returns>The current <see cref="WordListLevel"/> instance.</returns>
        /// <example>
        /// <code><![CDATA[
        /// WordList list = document.AddCustomList();
        /// WordListLevel level = new WordListLevel(WordListLevelKind.Decimal)
        ///     .SetStartNumberingValue(3);
        ///
        /// list.Numbering.AddLevel(level);
        /// list.AddItem("Starts at three");
        /// ]]></code>
        /// </example>
        public WordListLevel SetStartNumberingValue(int value) {
            StartNumberingValue = value;
            return this;
        }

        /// <summary>Gets or sets the physical text indentation for this level.</summary>
        public OfficeLength IndentationLeft {
            get => ReadTwips(GetIndentationElement().Attribute(W, "left"));
            set {
                string twips = checked((int)value.ToTwips()).ToString(CultureInfo.InvariantCulture);
                GetIndentationElement().SetAttribute(W, "left", twips);
            }
        }

        /// <summary>Gets or sets the physical number position for this level.</summary>
        public OfficeLength IndentationHanging {
            get => ReadTwips(GetIndentationElement().Attribute(W, "hanging"));
            set {
                string twips = checked((int)value.ToTwips()).ToString(CultureInfo.InvariantCulture);
                GetIndentationElement().SetAttribute(W, "hanging", twips);
            }
        }

        private static OfficeLength ReadTwips(string? value) {
            int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int nativeValue);
            return OfficeLength.Points(nativeValue / 20D);
        }

        /// <summary>
        /// Gets or sets the level text.
        /// </summary>
        public string LevelText {
            get => ReadValue(GetValueElement("lvlText", string.Empty), string.Empty, static _ => true);
            set => GetValueElement("lvlText", string.Empty).SetAttribute(W, "val", value ?? string.Empty);
        }

        /// <summary>
        /// Gets or sets the level justification.
        /// </summary>
        public WordListLevelAlignment LevelJustification {
            get => ReadLevelJustification(ReadValue(GetValueElement("lvlJc", "left"), "left", static token => ReadLevelJustification(token) != null))!.Value;
            set {
                string token = value.ToOpenXmlToken();
                GetValueElement("lvlJc", "left").SetAttribute(W, "val", token);
            }
        }

        /// <summary>
        /// Gets or sets the suffix placed after the numbering marker.
        /// </summary>
        public WordListLevelSuffix LevelSuffix {
            get {
                string token = ReadValue(GetValueElement("suff", "tab"), "tab",
                    static value => WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue(value, out WordListLevelSuffix _));
                WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out WordListLevelSuffix suffix);
                return suffix;
            }
            set {
                string token = value.ToOpenXmlToken();
                GetValueElement("suff", "tab").SetAttribute(W, "val", token);
            }
        }

        /// <summary>
        /// Removes this level from the list.
        /// </summary>
        public void Remove() {
            object? numbering = Node.AncestorOrSelf(W, "numbering").SdkElement;
            Node.Remove();
            WordListNumberingResolver.InvalidateNumbering(numbering);
        }

        /// <summary>Puts <paramref name="replacement"/> where this level is, and removes this level.</summary>
        internal void ReplaceWith(WordListLevel replacement) {
            replacement.Node = Node.InsertAfterSelf(replacement.Node);
            Remove();
        }

        /// <summary>
        /// Adds the level using custom simplified list number.
        /// </summary>
        /// <param name="simplifiedListNumbers"></param>
        /// <exception cref="ArgumentOutOfRangeException"></exception>
        public WordListLevel(WordListLevelKind simplifiedListNumbers) {
            Node = CreateLevel(simplifiedListNumbers);
        }

        /// <summary>
        /// A detached level of the given kind: the number format and level text, then, for every kind but
        /// <see cref="WordListLevelKind.None"/>, left justification, a half inch of indentation with a quarter inch
        /// hanging, and for some bullets the symbol font.
        /// </summary>
        private static OfficeXmlNode CreateLevel(WordListLevelKind kind) {
            (string format, string text, string? font) = kind switch {
                WordListLevelKind.None => ("none", string.Empty, null),
                WordListLevelKind.Bullet => ("bullet", "•", null),
                WordListLevelKind.BulletSquareSymbol => ("bullet", "■", null),
                WordListLevelKind.BulletBlackCircle => ("bullet", "●", null),
                WordListLevelKind.BulletDiamondSymbol => ("bullet", "◆", null),
                WordListLevelKind.BulletArrowSymbol => ("bullet", "►", null),
                WordListLevelKind.BulletSolidRound => ("bullet", "·", "Symbol"),
                WordListLevelKind.BulletOpenCircle => ("bullet", "o", "Symbol"),
                WordListLevelKind.BulletSquare2 => ("bullet", "■", null),
                WordListLevelKind.BulletSquare => ("bullet", "§", "Wingdings"),
                WordListLevelKind.BulletClubs => ("bullet", "v", "Wingdings"),
                WordListLevelKind.BulletArrow => ("bullet", "Ø", "Wingdings"),
                WordListLevelKind.BulletDiamond => ("bullet", "¨", "Symbol"),
                WordListLevelKind.BulletCheckmark => ("bullet", "ü", "Wingdings"),
                WordListLevelKind.Decimal => ("decimal", "%CurrentLevel", null),
                WordListLevelKind.DecimalBracket => ("decimal", "%CurrentLevel)", null),
                WordListLevelKind.DecimalDot => ("decimal", "%CurrentLevel.", null),
                WordListLevelKind.LowerLetter => ("lowerLetter", "%CurrentLevel", null),
                WordListLevelKind.LowerLetterBracket => ("lowerLetter", "%CurrentLevel)", null),
                WordListLevelKind.LowerLetterDot => ("lowerLetter", "%CurrentLevel.", null),
                WordListLevelKind.UpperLetter => ("upperLetter", "%CurrentLevel", null),
                WordListLevelKind.UpperLetterBracket => ("upperLetter", "%CurrentLevel)", null),
                WordListLevelKind.UpperLetterDot => ("upperLetter", "%CurrentLevel.", null),
                WordListLevelKind.LowerRoman => ("lowerRoman", "%CurrentLevel", null),
                WordListLevelKind.LowerRomanBracket => ("lowerRoman", "%CurrentLevel)", null),
                WordListLevelKind.LowerRomanDot => ("lowerRoman", "%CurrentLevel.", null),
                WordListLevelKind.UpperRoman => ("upperRoman", "%CurrentLevel", null),
                WordListLevelKind.UpperRomanBracket => ("upperRoman", "%CurrentLevel)", null),
                WordListLevelKind.UpperRomanDot => ("upperRoman", "%CurrentLevel.", null),
                _ => throw new ArgumentOutOfRangeException(nameof(kind), kind, null)
            };

            OfficeXmlTree tree = OfficeXmlTree.Create(new OfficeXmlNamespaces.Table(), W, "lvl", 16);
            var level = new OfficeXmlNode(tree, tree.Root);
            level.SetAttribute(W, "ilvl", "0");
            level.Append(level.CreateElement(W, "start")).SetAttribute(W, "val", "1");
            level.Append(level.CreateElement(W, "numFmt")).SetAttribute(W, "val", format);
            level.Append(level.CreateElement(W, "lvlText")).SetAttribute(W, "val", text);
            if (kind == WordListLevelKind.None) return level;

            level.Append(level.CreateElement(W, "lvlJc")).SetAttribute(W, "val", "left");
            OfficeXmlNode indentation = level.Append(level.CreateElement(W, "pPr")).Append(level.CreateElement(W, "ind"));
            indentation.SetAttribute(W, "left", "720");
            indentation.SetAttribute(W, "hanging", "360");
            if (font != null) {
                OfficeXmlNode fonts = level.Append(level.CreateElement(W, "rPr")).Append(level.CreateElement(W, "rFonts"));
                fonts.SetAttribute(W, "hint", "default");
                fonts.SetAttribute(W, "ascii", font);
                fonts.SetAttribute(W, "hAnsi", font);
            }
            return level;
        }
    }
}
