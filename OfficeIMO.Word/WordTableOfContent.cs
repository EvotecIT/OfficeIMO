using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.RegularExpressions;

namespace OfficeIMO.Word {
    /// <summary>
    /// Defines template styles that can be used when generating a table of contents.
    /// </summary>
    public enum WordTableOfContentsStyle {
        /// <summary>
        /// Built-in layout with a heading followed by entries.
        /// </summary>
        Template1,

        /// <summary>
        /// Alternative layout with the same structure but different identifiers.
        /// </summary>
        Template2
    }

    /// <summary>
    /// Represents a table of contents within a Word document.
    /// </summary>
    public partial class WordTableOfContent : WordElement {
        private static readonly string TocLevelSwitchPattern = @"\\o\s+(?:""\d+-\d+""|&quot;\d+-\d+&quot;)";
        private readonly WordDocument _document;
        private readonly SdtBlock _sdtBlock;

        /// <summary>
        /// Exposes the underlying structured document tag for this table of contents.
        /// </summary>
        internal SdtBlock SdtBlock => _sdtBlock;

        internal WordDocument Document => _document;

        /// <summary>
        /// True when Word shows the stored entries of this table: their page numbers are <c>PAGEREF</c> fields, as Word
        /// and <see cref="WordFieldUpdateOptions.UpdateTablesOfContents"/> write them, and nothing asks Word to rebuild the
        /// table when the document opens (<see cref="WordSettings.UpdateFieldsOnOpen"/> or a dirty field). Renderers draw
        /// such a table as stored and generate any other one from the headings, as Word shows it after the rebuild.
        /// </summary>
        internal bool ShowsStoredEntries {
            get {
                SdtContentBlock? content = _sdtBlock.SdtContentBlock;
                if (content == null || _document.Settings.UpdateFieldsOnOpen) return false;
                if (content.Descendants().Any(element => element is FieldChar { Dirty.Value: true } || element is SimpleField { Dirty.Value: true })) {
                    return false;
                }
                return EnumerateFieldInstructions(content).Any(instruction => GetFieldType(instruction.Instruction) == WordFieldType.PageRef);
            }
        }

        /// <summary>
        /// The complete instruction of every field under <paramref name="root"/>, in document order, with the element that
        /// starts the field: a simple field, or the begin character of a complex field. A complex field's instruction can be
        /// split across runs, so it is returned once its separator or end is reached.
        /// </summary>
        private static IEnumerable<(OpenXmlElement Start, string Instruction)> EnumerateFieldInstructions(OpenXmlElement root) {
            var open = new Stack<(FieldChar Begin, System.Text.StringBuilder? Instruction)>();
            foreach (OpenXmlElement element in root.Descendants()) {
                switch (element) {
                    case SimpleField simple:
                        yield return (simple, simple.Instruction?.Value ?? string.Empty);
                        break;
                    case FieldChar { FieldCharType.Value: var type } fieldChar:
                        if (type == FieldCharValues.Begin) {
                            open.Push((fieldChar, new System.Text.StringBuilder()));
                        } else if (open.Count > 0 && (type == FieldCharValues.Separate || type == FieldCharValues.End)) {
                            (FieldChar begin, System.Text.StringBuilder? instruction) = open.Pop();
                            if (instruction != null) yield return (begin, instruction.ToString());
                            // After the separator the field result follows; its instruction is already returned.
                            if (type == FieldCharValues.Separate) open.Push((begin, null));
                        }
                        break;
                    case FieldCode code when open.Count > 0:
                        open.Peek().Instruction?.Append(code.Text);
                        break;
                }
            }
        }

        private static WordFieldType? GetFieldType(string instruction) =>
            string.IsNullOrWhiteSpace(instruction) ? null : WordFieldInventory.ParseInstruction(instruction).FieldType;

        /// <summary>
        /// Gets the template style used to create this table of contents.
        /// </summary>
        public WordTableOfContentsStyle Style { get; }

        /// <summary>
        /// Gets the minimum heading level included by the TOC field switch.
        /// </summary>
        public int MinLevel {
            get {
                var levels = GetConfiguredLevels();
                return levels.MinLevel;
            }
        }

        /// <summary>
        /// Gets the maximum heading level included by the TOC field switch.
        /// </summary>
        public int MaxLevel {
            get {
                var levels = GetConfiguredLevels();
                return levels.MaxLevel;
            }
        }

        /// <summary>
        /// Gets or sets the heading text displayed for the table of contents.
        /// </summary>
        public string Text {
            get {
                return GetTitleTextElement()?.Text ?? string.Empty;
            }
            set {
                var text = GetTitleTextElement();
                if (text != null) {
                    text.Text = value;
                }
            }
        }

        /// <summary>Gets the existing title paragraph, or null when this table of contents has no title.</summary>
        /// <remarks>Formatting and page-break settings applied here are retained by <see cref="RefreshEntries"/>.</remarks>
        public WordParagraph? TitleParagraph {
            get {
                Paragraph? paragraph = GetTitleParagraph();
                if (paragraph == null) return null;
                Run? run = paragraph.Descendants<Run>().FirstOrDefault();
                return run == null ? new WordParagraph(_document, paragraph) : new WordParagraph(_document, paragraph, run);
            }
        }

        /// <summary>
        /// Gets or sets the text shown when the document has no entries for the table of contents.
        /// </summary>
        public string TextNoContent {
            get => GetPlaceholderText();
            set => SetPlaceholderText(value);
        }

        private Text? GetTitleTextElement() {
            var contentBlock = _sdtBlock?.SdtContentBlock;
            if (contentBlock == null) {
                return null;
            }

            foreach (var paragraph in contentBlock.ChildElements.OfType<Paragraph>()) {
                if (paragraph.Descendants<SimpleField>().Any()) {
                    continue;
                }

                var text = paragraph.Descendants<Text>().FirstOrDefault();
                if (text != null) {
                    return text;
                }
            }

            return null;
        }



        /// <summary>
        /// Initializes a new instance of the <see cref="WordTableOfContent"/> class and appends it to the document body.
        /// </summary>
        /// <param name="wordDocument">Parent document where the table of contents will be created.</param>
        /// <param name="tableOfContentStyle">Template style used to generate the table of contents.</param>
        /// <param name="minLevel">Minimum heading level to include (1..9).</param>
        /// <param name="maxLevel">Maximum heading level to include (1..9).</param>
        public WordTableOfContent(WordDocument wordDocument, WordTableOfContentsStyle tableOfContentStyle, int minLevel = 1, int maxLevel = 3)
            : this(
                wordDocument,
                PrepareTemplate(wordDocument, tableOfContentStyle),
                tableOfContentStyle,
                appendToBody: true,
                queueUpdateOnOpen: true,
                minLevel: minLevel,
                maxLevel: maxLevel) {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="WordTableOfContent"/> class using an existing structured document tag.
        /// </summary>
        /// <param name="wordDocument">Parent document that owns the table of contents.</param>
        /// <param name="sdtBlock">Structured document tag representing the table of contents.</param>
        internal WordTableOfContent(WordDocument wordDocument, SdtBlock sdtBlock)
            : this(wordDocument, sdtBlock, WordTableOfContentsStyle.Template1, appendToBody: false, queueUpdateOnOpen: true) {
        }

        internal WordTableOfContent(WordDocument wordDocument, SdtBlock sdtBlock, bool queueUpdateOnOpen)
            : this(wordDocument, sdtBlock, WordTableOfContentsStyle.Template1, appendToBody: false, queueUpdateOnOpen: queueUpdateOnOpen) {
        }

        private WordTableOfContent(
            WordDocument wordDocument,
            SdtBlock sdtBlock,
            WordTableOfContentsStyle style,
            bool appendToBody,
            bool queueUpdateOnOpen,
            int? minLevel = null,
            int? maxLevel = null) {
            this._document = wordDocument ?? throw new ArgumentNullException(nameof(wordDocument));
            this._sdtBlock = sdtBlock ?? throw new ArgumentNullException(nameof(sdtBlock));
            this.Style = style;
            GetPlaceholderPart();

            if (appendToBody) {
                this._document.AppendBlockToBody(_sdtBlock);
                // A TOC added by this library owns its title and entry styles and
                // its visible entries; both must exist for viewers that never update fields.
                this._document.EnsureStyleDefinitionsInitialized();
                WordTableOfContentStyles.EnsureDefined(this._document._wordprocessingDocument.MainDocumentPart?.StyleDefinitionsPart);
                this._document.InvalidateBuiltInStyleIndex();
                this._document.UseBuiltInStyle("TOCHeading");
                this._document.LocalizeBuiltInStyleReferences(_sdtBlock);
                this._document._generatedTableOfContent = this;
            }

            if (minLevel.HasValue && maxLevel.HasValue) {
                var (normalizedMin, normalizedMax) = NormalizeLevels(minLevel.Value, maxLevel.Value);
                // Avoid mutating the template when the caller requests the default 1-3 range.
                if (normalizedMin != 1 || normalizedMax != 3) {
                    ConfigureLevels(normalizedMin, normalizedMax);
                }
            }

            if (queueUpdateOnOpen) {
                QueueUpdateOnOpen();
            }
        }

        /// <summary>
        /// Flags the document to update this table of contents when the file is opened.
        /// </summary>
        public void Update() {
            QueueUpdateOnOpen(force: true);
        }

        internal void QueueUpdateOnOpen(bool force = false) {
            if (force) {
                MarkFieldsAsDirty();
            } else if (!MarkFieldsAsDirty()) {
                return;
            }

            // Write only when turning the setting on: the setter marks the value as the caller's.
            bool enabled = !this._document.Settings.UpdateFieldsOnOpen;
            if (enabled) {
                this._document.Settings.UpdateFieldsOnOpen = true;
            }

            this._document.NotifyTableOfContentUpdateQueued(enabled);
        }

        /// <summary>
        /// Updates the heading levels included by the TOC field switch (for example: \o "1-5").
        /// </summary>
        /// <param name="minLevel">Minimum heading level to include (1..9).</param>
        /// <param name="maxLevel">Maximum heading level to include (1..9).</param>
        public void SetLevels(int minLevel, int maxLevel) {
            ConfigureLevels(minLevel, maxLevel);
            QueueUpdateOnOpen(force: true);
        }

        /// <summary>
        /// Marks all fields that participate in the table of contents as dirty so Word refreshes them on open.
        /// </summary>
        /// <returns>
        /// <c>true</c> when an update should be queued. This includes corrupted tables of contents that no longer expose
        /// any fields so the caller can still request a refresh without throwing.
        /// </returns>
        private bool MarkFieldsAsDirty() {
            if (_sdtBlock?.SdtContentBlock == null) {
                return false;
            }

            var fieldsFound = false;
            var marked = false;

            foreach (var simpleField in _sdtBlock.Descendants<SimpleField>()) {
                fieldsFound = true;
                if (simpleField.Dirty?.Value != true) {
                    simpleField.Dirty = true;
                    marked = true;
                }
            }

            foreach (var fieldChar in _sdtBlock.Descendants<FieldChar>()) {
                fieldsFound = true;
                if (fieldChar.Dirty?.Value != true) {
                    fieldChar.Dirty = true;
                    marked = true;
                }
            }

            return fieldsFound ? marked : true;
        }

        private void ConfigureLevels(int minLevel, int maxLevel) {
            var (normalizedMin, normalizedMax) = NormalizeLevels(minLevel, maxLevel);
            var levelSwitch = $"\\o \"{normalizedMin}-{normalizedMax}\"";

            foreach (var simpleField in _sdtBlock.Descendants<SimpleField>()) {
                var instruction = simpleField.Instruction?.Value ?? simpleField.Instruction;
                if (string.IsNullOrWhiteSpace(instruction)) {
                    continue;
                }

                if (!IsTocInstruction(instruction)) {
                    continue;
                }

                var instructionText = instruction!;
                var hadLevelSwitch = Regex.IsMatch(instructionText, TocLevelSwitchPattern);
                string updated;
                if (hadLevelSwitch) {
                    updated = Regex.Replace(instructionText, TocLevelSwitchPattern, levelSwitch);
                } else {
                    updated = instructionText.TrimEnd() + " " + levelSwitch + " ";
                }

                simpleField.Instruction = updated;
                simpleField.Dirty = true;
            }

            // Word often converts TOC fields into complex fields (FieldChar/FieldCode).
            // Update those instructions as well so SetLevels works on existing documents.
            foreach (var fieldCode in _sdtBlock.Descendants<FieldCode>()) {
                var instructionText = fieldCode.Text;
                if (string.IsNullOrWhiteSpace(instructionText)) {
                    continue;
                }

                if (!IsTocInstruction(instructionText)) {
                    continue;
                }

                var hadLevelSwitch = Regex.IsMatch(instructionText, TocLevelSwitchPattern);
                string updated;
                if (hadLevelSwitch) {
                    updated = Regex.Replace(instructionText, TocLevelSwitchPattern, levelSwitch);
                } else {
                    updated = instructionText.TrimEnd() + " " + levelSwitch + " ";
                }

                fieldCode.Text = updated;
            }

            foreach (var fieldChar in _sdtBlock.Descendants<FieldChar>()) {
                if (fieldChar.Dirty?.Value != true) {
                    fieldChar.Dirty = true;
                }
            }
        }

        private (int MinLevel, int MaxLevel) GetConfiguredLevels() {
            if (_sdtBlock == null) {
                return (1, 3);
            }

            foreach (var simpleField in _sdtBlock.Descendants<SimpleField>()) {
                string? instruction = simpleField.Instruction?.Value ?? simpleField.Instruction;
                if (TryParseLevelSwitch(instruction, out int minLevel, out int maxLevel)) {
                    return (minLevel, maxLevel);
                }
            }

            foreach (var fieldCode in _sdtBlock.Descendants<FieldCode>()) {
                if (TryParseLevelSwitch(fieldCode.Text, out int minLevel, out int maxLevel)) {
                    return (minLevel, maxLevel);
                }
            }

            return (1, 3);
        }

        private static bool TryParseLevelSwitch(string? instruction, out int minLevel, out int maxLevel) {
            minLevel = 1;
            maxLevel = 3;
            if (string.IsNullOrWhiteSpace(instruction) ||
                instruction!.IndexOf("TOC", StringComparison.OrdinalIgnoreCase) < 0) {
                return false;
            }

            Match match = Regex.Match(instruction!, @"\\o\s+(?:""(?<min>\d+)-(?<max>\d+)""|&quot;(?<min>\d+)-(?<max>\d+)&quot;)");
            if (!match.Success ||
                !int.TryParse(match.Groups["min"].Value, out int parsedMin) ||
                !int.TryParse(match.Groups["max"].Value, out int parsedMax)) {
                return false;
            }

            (minLevel, maxLevel) = NormalizeLevels(parsedMin, parsedMax);
            return true;
        }

        private static (int MinLevel, int MaxLevel) NormalizeLevels(int minLevel, int maxLevel) {
            var min = ClampLevel(minLevel);
            var max = ClampLevel(maxLevel);
            if (min > max) {
                (min, max) = (max, min);
            }
            return (min, max);
        }

        private static int ClampLevel(int level) {
            if (level < 1) {
                return 1;
            }
            if (level > 9) {
                return 9;
            }
            return level;
        }

        /// <summary>
        /// Deletes this table of contents from the parent document.
        /// </summary>
        public void Remove() {
            _document.RemoveTableOfContent();
        }

        /// <summary>
        /// Removes this table of contents and creates a new one in the same location.
        /// </summary>
        /// <returns>The newly created <see cref="WordTableOfContent"/> instance.</returns>
        public WordTableOfContent Regenerate() {
            return _document.RegenerateTableOfContent();
        }

        /// <summary>
        /// Returns a predefined structured document tag matching the chosen style.
        /// </summary>
        /// <param name="document">Document that will own the structured document tag.</param>
        /// <param name="style">Template identifier to retrieve.</param>
        private static SdtBlock PrepareTemplate(WordDocument document, WordTableOfContentsStyle style) {
            var block = GetStyle(style);
            document.AssignNewSdtIds(block);
            return block;
        }

        private static SdtBlock GetStyle(WordTableOfContentsStyle style) =>
            WordEmbeddedMarkup.CreateBuildingBlock(WordEmbeddedMarkup.TablesOfContents, style.ToString(), nameof(style));
    }
}
