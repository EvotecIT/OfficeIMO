using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using System.Text;
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
        private const int TocNamespace = OfficeXmlNamespaces.WordprocessingML;
        private static readonly string TocLevelSwitchPattern = @"\\o\s+(?:""\d+-\d+""|&quot;\d+-\d+&quot;)";
        private readonly WordDocument _document;
        private SdtBlock? _sdkSdtBlock;
        private SdtBlock _sdtBlock {
            get {
                if (_sdkSdtBlock != null) return _sdkSdtBlock;
                if (_document.IsDisposed) _document.OpenSdkPackageAfterDispose();
                else _document.OpenSdkPackage(null);
                return _sdkSdtBlock = (SdtBlock?)ResolveControl().SdkElement
                    ?? throw new InvalidOperationException("The table of contents is no longer attached to the document.");
            }
        }

        /// <summary>
        /// Exposes the underlying structured document tag for this table of contents.
        /// </summary>
        internal SdtBlock SdtBlock => _sdtBlock;

        internal OfficeXmlNode SdtNode => _sdkSdtBlock != null ? OfficeXmlNode.FromSdk(_sdkSdtBlock) : ResolveControl();
        private OfficeXmlNode ContentNode => SdtNode.Child(TocNamespace, "sdtContent");

        private OfficeXmlNode GetGlossaryRoot(bool create) {
            if (_document._ownedModel is { } owned) return GetOwnedGlossaryRoot(owned, create);
            if (_document.IsDisposed && _document.DisposedReadModel is { } copy)
                return GetOwnedGlossaryRoot(copy, create: false);
            MainDocumentPart? main = _document._wordprocessingDocument.MainDocumentPart;
            GlossaryDocumentPart? glossary = main?.GlossaryDocumentPart;
            if (create) {
                glossary ??= main!.AddNewPart<GlossaryDocumentPart>();
                glossary.GlossaryDocument ??= new GlossaryDocument();
            }
            return OfficeXmlNode.FromSdk(glossary?.GlossaryDocument);
        }

        private void LocalizeGeneratedStyles(OfficeXmlNode content) {
            if (_document._ownedModel is { } model) LocalizeOwnedTocStyles(content, model);
            else if (content.SdkElement is SdtContentBlock sdk) _document.LocalizeBuiltInStyleReferences(sdk);
        }

        internal WordDocument Document => _document;

        /// <summary>
        /// True when Word shows the stored entries of this table: their page numbers are <c>PAGEREF</c> fields, as Word
        /// and <see cref="WordFieldUpdateOptions.UpdateTablesOfContents"/> write them, and nothing asks Word to rebuild the
        /// table when the document opens (<see cref="WordSettings.UpdateFieldsOnOpen"/> or a dirty field). Renderers draw
        /// such a table as stored and generate any other one from the headings, as Word shows it after the rebuild.
        /// </summary>
        internal bool ShowsStoredEntries {
            get {
                OfficeXmlNode content = ContentNode;
                if (content.IsNone || ReadTocUpdateOnOpen()) return false;
                foreach (OfficeXmlNode marker in WordFieldInventory.EnumerateFieldContent(content)) {
                    if (IsFieldMarker(marker) && WordFieldInventory.ReadFlag(marker, "dirty")) return false;
                }
                return EnumerateFieldInstructions(content).Any(instruction => GetFieldType(instruction.Instruction) == WordFieldType.PageRef);
            }
        }

        /// <summary>
        /// Reads instructions from their simple field or complex field begin marker.
        /// A complex instruction can span several runs and is complete at its separator or end.
        /// </summary>
        private static IEnumerable<(OfficeXmlNode Start, string Instruction)> EnumerateFieldInstructions(OfficeXmlNode root) {
            const int W = TocNamespace;
            var open = new Stack<(OfficeXmlNode Begin, StringBuilder? Instruction)>();
            foreach (OfficeXmlNode element in WordFieldInventory.EnumerateFieldContent(root)) {
                if (element.Is(W, "fldSimple")) {
                    yield return (element, element.Attribute(W, "instr") ?? string.Empty);
                } else if (element.Is(W, "fldChar")) {
                    string? type = element.Attribute(W, "fldCharType");
                    if (type == "begin") {
                        open.Push((element, new StringBuilder()));
                    } else if (open.Count > 0 && (type is "separate" or "end")) {
                        (OfficeXmlNode begin, StringBuilder? instruction) = open.Pop();
                        if (instruction != null) yield return (begin, instruction.ToString());
                        if (type == "separate") open.Push((begin, null));
                    }
                } else if (open.Count > 0 && element.Is(W, "instrText")) {
                    open.Peek().Instruction?.Append(element.InnerText);
                }
            }
        }

        private static bool IsFieldMarker(OfficeXmlNode element) =>
            element.Is(TocNamespace, "fldSimple") || element.Is(TocNamespace, "fldChar");

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
            get => GetTitleTextElement().InnerText;
            set {
                OfficeXmlNode text = GetTitleTextElement();
                if (text.Exists) {
                    text.SetInnerText(value ?? string.Empty);
                    RecordOwnedTocEdit();
                }
            }
        }

        /// <summary>Gets the existing title paragraph, or null when this table of contents has no title.</summary>
        /// <remarks>Formatting and page-break settings applied here are retained by <see cref="RefreshEntries"/>.</remarks>
        public WordParagraph? TitleParagraph {
            get {
                OfficeXmlNode title = GetTitleParagraphNode();
                if (title.IsNone) return null;
                if (title.Tree is { } tree) {
                    OfficeXmlNode ownedRun = WordFieldInventory.EnumerateFieldContent(title)
                        .FirstOrDefault(node => node.Is(TocNamespace, "r"));
                    return WordParagraph.CreateOwned(_document, tree, title.Node, ownedRun.Node);
                }
                var paragraph = (Paragraph)title.SdkElement!;

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

        private OfficeXmlNode GetTitleTextElement() {
            const int W = TocNamespace;
            foreach (OfficeXmlNode paragraph in ContentNode.Children(W, "p")) {
                if (WordFieldInventory.EnumerateFieldContent(paragraph).Any(element => element.Is(W, "fldSimple"))) continue;
                OfficeXmlNode text = WordFieldInventory.EnumerateFieldContent(paragraph)
                    .FirstOrDefault(element => element.Is(W, "t"));
                if (text.Exists) return text;
            }
            return default;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="WordTableOfContent"/> class and appends it to the document body.
        /// </summary>
        /// <param name="wordDocument">Parent document where the table of contents will be created.</param>
        /// <param name="tableOfContentStyle">Template style used to generate the table of contents.</param>
        /// <param name="minLevel">Minimum heading level to include (1..9).</param>
        /// <param name="maxLevel">Maximum heading level to include (1..9).</param>
        public WordTableOfContent(WordDocument wordDocument, WordTableOfContentsStyle tableOfContentStyle, int minLevel = 1, int maxLevel = 3) {
            _document = wordDocument ?? throw new ArgumentNullException(nameof(wordDocument));
            Style = tableOfContentStyle;
            if (!TryAppendOwnedTemplate(tableOfContentStyle)) {
                _sdkSdtBlock = PrepareTemplate(wordDocument, tableOfContentStyle);
                GetPlaceholderPart();
                _document.AppendBlockToBody(_sdkSdtBlock);
                _document.EnsureStyleDefinitionsInitialized();
                WordTableOfContentStyles.EnsureDefined(_document._wordprocessingDocument.MainDocumentPart?.StyleDefinitionsPart);
                _document.InvalidateBuiltInStyleIndex();
                _document.UseBuiltInStyle("TOCHeading");
                _document.LocalizeBuiltInStyleReferences(_sdkSdtBlock);
            }
            _document._generatedTableOfContent = this;

            var (normalizedMin, normalizedMax) = NormalizeLevels(minLevel, maxLevel);
            if (normalizedMin != 1 || normalizedMax != 3) ConfigureLevels(normalizedMin, normalizedMax);
            QueueUpdateOnOpen();
        }

        /// <summary>Wraps an imported structured document tag representing a table of contents.</summary>
        internal WordTableOfContent(WordDocument wordDocument, SdtBlock sdtBlock, bool queueUpdateOnOpen = true) {
            _document = wordDocument ?? throw new ArgumentNullException(nameof(wordDocument));
            _sdkSdtBlock = sdtBlock ?? throw new ArgumentNullException(nameof(sdtBlock));
            Style = WordTableOfContentsStyle.Template1;
            GetPlaceholderPart();
            if (queueUpdateOnOpen) QueueUpdateOnOpen();
        }

        private static WordTableOfContent? WrapSdkControl(WordDocument document, OfficeXmlNode control) =>
            control.SdkElement is SdtBlock block ? new WordTableOfContent(document, block, queueUpdateOnOpen: false) : null;

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
            bool enabled = !ReadTocUpdateOnOpen();
            if (enabled) EnableTocUpdateOnOpen();

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
            if (ContentNode.IsNone) return false;
            bool fieldsFound = false;
            bool marked = false;
            foreach (OfficeXmlNode field in WordFieldInventory.EnumerateFieldContent(SdtNode)) {
                if (!IsFieldMarker(field)) continue;
                fieldsFound = true;
                if (WordFieldInventory.ReadFlag(field, "dirty")) continue;
                field.SetAttribute(TocNamespace, "dirty", "true");
                marked = true;
            }
            if (marked) RecordOwnedTocEdit();
            return !fieldsFound || marked;
        }

        private void ConfigureLevels(int minLevel, int maxLevel) {
            const int W = TocNamespace;
            var (normalizedMin, normalizedMax) = NormalizeLevels(minLevel, maxLevel);
            var levelSwitch = $"\\o \"{normalizedMin}-{normalizedMax}\"";
            foreach (OfficeXmlNode field in WordFieldInventory.EnumerateFieldContent(SdtNode)) {
                if (field.Is(W, "fldSimple")) {
                    string? instruction = field.Attribute(W, "instr");
                    if (!string.IsNullOrWhiteSpace(instruction) && IsTocInstruction(instruction!)) {
                        field.SetAttribute(W, "instr", UpdateLevelSwitch(instruction!));
                        field.SetAttribute(W, "dirty", "true");
                    }
                } else if (field.Is(W, "instrText")) {
                    string instruction = field.InnerText;
                    if (!string.IsNullOrWhiteSpace(instruction) && IsTocInstruction(instruction))
                        field.SetInnerText(UpdateLevelSwitch(instruction));
                } else if (field.Is(W, "fldChar") && !WordFieldInventory.ReadFlag(field, "dirty")) {
                    field.SetAttribute(W, "dirty", "true");
                }
            }

            RecordOwnedTocEdit();

            string UpdateLevelSwitch(string instruction) =>
                Regex.IsMatch(instruction, TocLevelSwitchPattern)
                    ? Regex.Replace(instruction, TocLevelSwitchPattern, levelSwitch)
                    : instruction.TrimEnd() + " " + levelSwitch + " ";
        }

        private (int MinLevel, int MaxLevel) GetConfiguredLevels() {
            const int W = TocNamespace;
            (int MinLevel, int MaxLevel)? complex = null;
            foreach (OfficeXmlNode element in WordFieldInventory.EnumerateFieldContent(SdtNode)) {
                if (element.Is(W, "fldSimple") && TryParseLevelSwitch(element.Attribute(W, "instr"), out int min, out int max))
                    return (min, max);
                if (complex == null && element.Is(W, "instrText") && TryParseLevelSwitch(element.InnerText, out min, out max))
                    complex = (min, max);
            }
            return complex ?? (1, 3);
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
