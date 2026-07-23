using OfficeIMO.Rtf.Diagnostics;
using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

internal static partial class RtfSemanticReader {
    public static RtfReadResult Read(RtfSyntaxTree tree, RtfReadOptions options, CancellationToken cancellationToken = default) {
        if (tree == null) throw new ArgumentNullException(nameof(tree));
        options ??= new RtfReadOptions();

        var diagnostics = new List<RtfDiagnostic>(tree.Diagnostics);
        var binder = new Binder(options, diagnostics, cancellationToken);
        RtfDocument document = binder.Bind(tree.Root);
        return new RtfReadResult(document, tree, diagnostics.AsReadOnly());
    }

    private sealed partial class Binder {
        private readonly RtfReadOptions _options;
        private readonly List<RtfDiagnostic> _diagnostics;
        private readonly RtfReadLimitGuard _limits;
        private RtfDocument _document = null!;
        private RtfParagraph _currentParagraph = null!;
        private RtfTable? _currentTable;
        private RtfTableRow? _currentRow;
        private RtfHeaderFooter? _currentHeaderFooter;
        private RtfNote? _currentNote;
        private RtfShape? _currentShape;
        private RtfSection? _currentSection;
        private int _currentCellIndex;
        private int _currentCellDefinitionIndex;
        private bool _currentParagraphIsInTable;
        private PendingTableCellProperties _pendingCellProperties = new PendingTableCellProperties();
        private RtfTableRowBorderSide? _currentRowBorderSide;
        private RowBoxMeasurements _currentRowPadding = new RowBoxMeasurements();
        private RowBoxMeasurements _currentRowSpacing = new RowBoxMeasurements();
        private int _inlineCaptureDepth;
        private bool _hasSemanticSections;
        private int? _currentSectionColumnNumber;
        private Dictionary<int, RtfListOverride> _listOverridesById = null!;
        private Dictionary<int, RtfListDefinition> _listDefinitionsById = null!;
        private Dictionary<int, RtfFont> _fontsById = null!;
        private readonly List<NestedTableContext> _nestedTableContexts = new List<NestedTableContext>();
        private readonly HashSet<int> _nestedTableBoundaryLevels = new HashSet<int>();

        public Binder(RtfReadOptions options, List<RtfDiagnostic> diagnostics, CancellationToken cancellationToken) {
            _options = options;
            _diagnostics = diagnostics;
            _limits = new RtfReadLimitGuard(options, cancellationToken);
        }

        public RtfDocument Bind(RtfGroup root) {
            _document = RtfDocument.Create();
            RtfControlWord? ansiCodePageControl = FindAnsiCodePageControl(root);
            RtfDocumentCharacterSet? characterSet = FindDocumentCharacterSet(root);
            int ansiCodePage = ansiCodePageControl?.Parameter ?? RtfAnsiCodePage.GetDefaultCodePage(characterSet);
            RtfControlWord? unicodeSkipCountControl = FindUnicodeSkipCountControl(root);
            int unicodeSkipCount = unicodeSkipCountControl?.Parameter ?? 1;
            if (_options.WarnOnUnsupportedCodePages && !RtfAnsiCodePage.IsSupported(ansiCodePage)) {
                _diagnostics.Add(new RtfDiagnostic(RtfDiagnosticSeverity.Warning, "RTF103", $"ANSI code page '{ansiCodePage}' is not supported yet; Windows-1252 fallback decoding was used for semantic text.", ansiCodePageControl?.Position ?? root.Position));
            }

            _document.ReplaceFonts(ReadFontTable(root, ansiCodePage, unicodeSkipCount));
            _fontsById = new Dictionary<int, RtfFont>();
            foreach (RtfFont font in _document.Fonts) {
                if (!_fontsById.ContainsKey(font.Id)) _fontsById.Add(font.Id, font);
            }
            _document.ReplaceColors(ReadColorTable(root));
            _document.ReplaceStyles(ReadStylesheet(root, ansiCodePage, unicodeSkipCount));
            _document.ReplaceListDefinitions(ReadListDefinitions(root, ansiCodePage, unicodeSkipCount));
            _document.ReplaceListOverrides(ReadListOverrides(root));
            _document.ReplaceRevisionAuthors(ReadRevisionAuthors(root, ansiCodePage, unicodeSkipCount));
            _document.RevisionRootSaveId = ReadRevisionRootSaveId(root);
            _document.ReplaceRevisionSaveIds(ReadRevisionSaveIds(root));
            if (_options.ReadFileReferences) {
                _document.ReplaceFileReferences(ReadFileReferences(root, ansiCodePage, unicodeSkipCount));
            } else if (root.Children.OfType<RtfGroup>().Any(group => group.Destination == "filetbl")) {
                _diagnostics.Add(new RtfDiagnostic(RtfDiagnosticSeverity.Warning, "RTF106", "File-table references were blocked by the configured read policy.", root.Position));
            }
            _document.ReplaceXmlNamespaces(ReadXmlNamespaces(root, ansiCodePage, unicodeSkipCount));
            _listDefinitionsById = CreateListDefinitionLookup(_document.ListDefinitions);
            _listOverridesById = CreateListOverrideLookup(_document.ListOverrides);
            ReadPageSetup(root, _document.PageSetup);
            ReadNoteSettings(root, _document.NoteSettings);
            ReadDocumentSettings(root, _document.Settings);
            ReadInfo(root, _document.Info, ansiCodePage, unicodeSkipCount);
            _document.ReplaceUserProperties(ReadUserProperties(root, ansiCodePage, unicodeSkipCount));
            _document.ReplaceDocumentVariables(ReadDocumentVariables(root, ansiCodePage, unicodeSkipCount));
            _document.HtmlEncapsulation = ReadHtmlEncapsulation(root, ansiCodePage, unicodeSkipCount);
            _currentParagraph = new RtfParagraph();
            _currentSection = new RtfSection();

            bool hasExplicitAnsiCodePage = ansiCodePageControl != null;
            WalkGroup(root, CreateInitialState(ansiCodePage, hasExplicitAnsiCodePage, unicodeSkipCount), depth: 0, allowDestinationSkip: true);
            FlushParagraphIfNeeded(force: true, CreateInitialState(ansiCodePage, hasExplicitAnsiCodePage, unicodeSkipCount));
            CompleteOpenSection();
            return _document;
        }

        private CharacterState CreateInitialState(int ansiCodePage, bool hasExplicitAnsiCodePage, int unicodeSkipCount) {
            int effectiveCodePage = ResolveFontCodePage(_document.Settings.DefaultFontId, ansiCodePage);
            return new CharacterState {
                AnsiCodePage = effectiveCodePage,
                DocumentAnsiCodePage = ansiCodePage,
                HasExplicitAnsiCodePage = hasExplicitAnsiCodePage,
                UnicodeSkipCount = unicodeSkipCount,
                DefaultLanguageId = _document.Settings.DefaultLanguageId,
                LanguageId = _document.Settings.DefaultLanguageId
            };
        }

        private int ResolveFontCodePage(int? fontId, int fallbackCodePage) {
            if (!fontId.HasValue) return fallbackCodePage;
            if (!_fontsById.TryGetValue(fontId.Value, out RtfFont? font)) return fallbackCodePage;
            if (font.CodePage.HasValue) return font.CodePage.Value;
            return RtfAnsiCodePage.GetCodePageForCharset(font.Charset) ?? fallbackCodePage;
        }

        private void WalkGroup(RtfGroup group, CharacterState state, int depth, bool allowDestinationSkip) {
            _limits.CheckCancellation();
            if (depth > _options.MaxDepth) {
                _diagnostics.Add(new RtfDiagnostic(RtfDiagnosticSeverity.Error, "RTF100", "Maximum RTF group depth was exceeded.", group.Position));
                return;
            }

            string? destination = group.Destination;
            RtfBookmarkMarkerKind? bookmarkMarkerKind = TryGetBookmarkMarkerKind(destination);
            if (bookmarkMarkerKind.HasValue) {
                ReadBookmarkMarker(group, bookmarkMarkerKind.Value, state);
                return;
            }

            RtfNoteKind? noteKind = TryGetNoteKind(destination);
            if (noteKind.HasValue) {
                ReadNote(group, noteKind.Value, state, depth);
                return;
            }

            RtfHeaderFooterKind? headerFooterKind = TryGetHeaderFooterKind(destination);
            if (headerFooterKind.HasValue) {
                ReadHeaderFooter(group, headerFooterKind.Value, state, depth);
                return;
            }

            if (destination == "pict") {
                RtfImage? image = ReadPicture(group);
                if (image != null) {
                    AddPicture(image);
                }

                return;
            }

            if (destination == "object") {
                if (!_options.ReadEmbeddedObjects) {
                    _diagnostics.Add(new RtfDiagnostic(RtfDiagnosticSeverity.Warning, "RTF105", "Embedded object was blocked by the configured read policy.", group.Position));
                    return;
                }

                RtfObject? rtfObject = ReadObject(group, state, depth);
                if (rtfObject != null) {
                    _currentParagraph.AddObject(rtfObject);
                }

                return;
            }

            if (destination == "shp") {
                RtfShape? shape = ReadShape(group, state, depth);
                if (shape != null) {
                    AddShape(shape);
                }

                return;
            }

            if (destination == "field") {
                if (TryAppendField(group, state, depth)) {
                    return;
                }
            }

            if (destination == "nesttableprops") {
                ReadNestedTableProperties(group, state);
                return;
            }

            if (destination == "officeimonestedtableboundary") {
                ReadNestedTableBoundary(group);
                return;
            }

            if (allowDestinationSkip && destination == "listtext") {
                ReadListText(group, state, depth);
                return;
            }

            if (destination == "upr") {
                RtfGroup? unicodeAlternative = FindUnicodeAlternative(group);
                if (unicodeAlternative != null) {
                    WalkGroup(unicodeAlternative, state.Clone(), depth + 1, allowDestinationSkip: false);
                    return;
                }
            }

            bool isIgnorableDestination = RtfDestinationRegistry.IsIgnorableDestinationGroup(group);
            if (allowDestinationSkip && (RtfDestinationRegistry.ShouldSkipSemanticBinding(destination) || isIgnorableDestination)) {
                if (_options.WarnOnUnsupportedDestinations &&
                    (RtfDestinationRegistry.IsUnsupportedSemanticDestination(destination) ||
                     (isIgnorableDestination && !RtfDestinationRegistry.IsKnown(destination)))) {
                    _diagnostics.Add(new RtfDiagnostic(RtfDiagnosticSeverity.Warning, "RTF101", $"Destination '{destination}' is preserved in syntax but not bound semantically yet.", group.Position));
                }

                return;
            }

            var childState = state.Clone();
            foreach (RtfNode node in group.Children) {
                _limits.CheckCancellation();
                switch (node) {
                    case RtfGroup childGroup:
                        if (childGroup.Destination == "pn") {
                            ReadLegacyNumbering(childGroup, childState);
                        } else if (childGroup.Destination == "listtext") {
                            ReadListText(childGroup, childState, depth + 1);
                        } else {
                            WalkGroup(childGroup, childState.Clone(), depth + 1, allowDestinationSkip: true);
                        }
                        break;
                    case RtfText text:
                        AppendAnsiText(text.Text, childState);
                        break;
                    case RtfControlWord control:
                        ApplyControlWord(control, childState);
                        break;
                    case RtfControlSymbol symbol:
                        ApplyControlSymbol(symbol, childState);
                        break;
                }
            }
        }

        private static RtfGroup? FindUnicodeAlternative(RtfGroup group) => group.Children.OfType<RtfGroup>().FirstOrDefault(child => child.Destination == "ud");

        private void ApplyControlSymbol(RtfControlSymbol symbol, CharacterState state) {
            switch (symbol.Symbol) {
                case '\\':
                case '{':
                case '}':
                    AppendText(symbol.Symbol.ToString(), state);
                    return;
                case '~':
                    AppendText("\u00A0", state);
                    return;
                case '_':
                    AppendText("\u2011", state);
                    return;
                case '-':
                    AppendText("\u00AD", state);
                    return;
                case '\'':
                    if (symbol.Parameter.HasValue) {
                        AppendAnsiByte(symbol.Parameter.Value, state);
                    }
                    return;
            }
        }

        private static string GetSpecialCharacterText(string controlName) {
            switch (controlName) {
                case "emdash":
                    return "\u2014";
                case "endash":
                    return "\u2013";
                case "emspace":
                    return "\u2003";
                case "enspace":
                    return "\u2002";
                case "qmspace":
                    return "\u2005";
                case "bullet":
                    return "\u2022";
                case "lquote":
                    return "\u2018";
                case "rquote":
                    return "\u2019";
                case "ldblquote":
                    return "\u201C";
                case "rdblquote":
                    return "\u201D";
                case "ltrmark":
                    return "\u200E";
                case "rtlmark":
                    return "\u200F";
                case "zwj":
                    return "\u200D";
                case "zwnj":
                    return "\u200C";
                default:
                    return string.Empty;
            }
        }

        private void AppendText(string text, CharacterState state) {
            if (state.PendingAnsiLeadByte.HasValue) {
                text = "\uFFFD" + text;
                state.PendingAnsiLeadByte = null;
            }

            if (string.IsNullOrEmpty(text)) return;
            if (IsFormattingTrivia(text)) return;
            if (state.PendingHighSurrogate.HasValue) {
                text = "\uFFFD" + text;
                state.PendingHighSurrogate = null;
            }

            ApplyParagraphState(_currentParagraph, state);
            var run = new RtfRun(text) {
                Bold = state.Bold,
                Italic = state.Italic,
                UnderlineStyle = state.UnderlineStyle,
                Strike = state.Strike,
                DoubleStrike = state.DoubleStrike,
                Hidden = state.Hidden,
                Outline = state.Outline,
                Shadow = state.Shadow,
                Emboss = state.Emboss,
                Imprint = state.Imprint,
                CapsStyle = state.CapsStyle,
                VerticalPosition = state.VerticalPosition,
                FontSize = state.FontSize,
                FontId = state.FontId,
                ForegroundColorIndex = state.ForegroundColorIndex,
                HighlightColorIndex = state.HighlightColorIndex,
                CharacterBackgroundColorIndex = state.CharacterBackgroundColorIndex,
                CharacterShadingForegroundColorIndex = state.CharacterShadingForegroundColorIndex,
                CharacterShadingPatternPercent = state.CharacterShadingPatternPercent,
                CharacterShadingPattern = state.CharacterShadingPattern,
                UnderlineColorIndex = state.UnderlineColorIndex,
                CharacterSpacingTwips = state.CharacterSpacingTwips,
                CharacterScalePercent = state.CharacterScalePercent,
                KerningHalfPoints = state.KerningHalfPoints,
                CharacterOffsetHalfPoints = state.CharacterOffsetHalfPoints,
                StyleId = state.CharacterStyleId,
                Direction = state.Direction,
                RevisionKind = state.RevisionKind,
                RevisionAuthorIndex = state.RevisionAuthorIndex,
                RevisionTimestampValue = state.RevisionTimestampValue,
                CharacterRevisionSaveId = state.CharacterRevisionSaveId,
                InsertionRevisionSaveId = state.InsertionRevisionSaveId,
                DeletionRevisionSaveId = state.DeletionRevisionSaveId,
                LanguageId = state.LanguageId
            };
            run.CharacterBorder.CopyFrom(state.CharacterBorder);
            _currentParagraph.AddRun(run);
        }

        private void AppendAnsiText(string text, CharacterState state) {
            if (string.IsNullOrEmpty(text)) return;
            int start = ConsumeAnsiFallbackBytes(state, text.Length);
            if (start >= text.Length) return;
            if (state.PendingAnsiLeadByte.HasValue) {
                if (text[start] <= byte.MaxValue) {
                    byte lead = state.PendingAnsiLeadByte.Value;
                    state.PendingAnsiLeadByte = null;
                    AppendText(RtfAnsiCodePage.DecodeBytes(state.AnsiCodePage, new[] { lead, (byte)text[start] }), state);
                    start++;
                } else {
                    state.PendingAnsiLeadByte = null;
                    AppendText("\uFFFD", state);
                }
            }

            if (start < text.Length) {
                AppendText(RtfAnsiCodePage.DecodeText(state.AnsiCodePage, text.Substring(start)), state);
            }
        }

        private void AppendAnsiByte(int value, CharacterState state) {
            if (ConsumeAnsiFallbackBytes(state, 1) == 1) return;
            byte current = (byte)(value & 0xFF);
            if (state.PendingAnsiLeadByte.HasValue) {
                byte lead = state.PendingAnsiLeadByte.Value;
                state.PendingAnsiLeadByte = null;
                AppendText(RtfAnsiCodePage.DecodeBytes(state.AnsiCodePage, new[] { lead, current }), state);
                return;
            }

            if (RtfAnsiCodePage.IsLeadByte(state.AnsiCodePage, current)) {
                state.PendingAnsiLeadByte = current;
                return;
            }

            AppendText(RtfAnsiCodePage.DecodeByte(state.AnsiCodePage, current), state);
        }

        private void AppendGeneratedText(RtfGeneratedTextKind kind, CharacterState state) {
            ApplyParagraphState(_currentParagraph, state);
            _currentParagraph.AddGeneratedText(kind);
        }

        private static void ApplyParagraphState(RtfParagraph paragraph, CharacterState state) {
            paragraph.Alignment = state.Alignment;
            paragraph.Direction = state.ParagraphDirection;
            paragraph.StyleId = state.ParagraphStyleId;
            paragraph.ListId = state.ListId;
            paragraph.ListDefinitionId = state.ListDefinitionId;
            paragraph.ListLevel = state.ListLevel;
            paragraph.ListKind = state.ListKind;
            paragraph.LegacyNumbering.CopyFrom(state.LegacyNumbering);
            paragraph.SetParsedListText(state.ListText);
            paragraph.LeftIndentTwips = state.LeftIndentTwips;
            paragraph.RightIndentTwips = state.RightIndentTwips;
            paragraph.FirstLineIndentTwips = state.FirstLineIndentTwips;
            paragraph.SpaceBeforeTwips = state.SpaceBeforeTwips;
            paragraph.SpaceAfterTwips = state.SpaceAfterTwips;
            paragraph.SpaceBeforeAuto = state.SpaceBeforeAuto;
            paragraph.SpaceAfterAuto = state.SpaceAfterAuto;
            paragraph.LineSpacingTwips = state.LineSpacingTwips;
            paragraph.LineSpacingMultiple = state.LineSpacingMultiple;
            paragraph.BackgroundColorIndex = state.BackgroundColorIndex;
            paragraph.ShadingForegroundColorIndex = state.ShadingForegroundColorIndex;
            paragraph.ShadingPatternPercent = state.ShadingPatternPercent;
            paragraph.ShadingPattern = state.ShadingPattern;
            CopyParagraphBorder(state.TopBorder, paragraph.TopBorder);
            CopyParagraphBorder(state.LeftBorder, paragraph.LeftBorder);
            CopyParagraphBorder(state.BottomBorder, paragraph.BottomBorder);
            CopyParagraphBorder(state.RightBorder, paragraph.RightBorder);
            paragraph.PageBreakBefore = state.PageBreakBefore;
            paragraph.KeepWithNext = state.KeepWithNext;
            paragraph.KeepLinesTogether = state.KeepLinesTogether;
            paragraph.SuppressLineNumbers = state.SuppressLineNumbers;
            paragraph.AutoHyphenation = state.AutoHyphenation;
            paragraph.ContextualSpacing = state.ContextualSpacing;
            paragraph.AdjustRightIndent = state.AdjustRightIndent;
            paragraph.SnapToLineGrid = state.SnapToLineGrid;
            paragraph.WidowControl = state.WidowControl;
            paragraph.OutlineLevel = state.OutlineLevel;
            paragraph.RevisionSaveId = state.ParagraphRevisionSaveId;
            paragraph.Frame.CopyFrom(state.Frame);
            paragraph.ReplaceTabStops(state.TabStops);
        }

        private static void CopyParagraphBorder(RtfParagraphBorder source, RtfParagraphBorder destination) {
            destination.Style = source.Style;
            destination.Width = source.Width;
            destination.ColorIndex = source.ColorIndex;
        }

        private static bool IsFormattingTrivia(string text) {
            for (int i = 0; i < text.Length; i++) {
                if (text[i] != '\r' && text[i] != '\n') {
                    return false;
                }
            }

            return true;
        }

        private void FlushParagraphIfNeeded(bool force, CharacterState state) {
            if (_inlineCaptureDepth > 0) return;
            if (!force && _currentParagraph.Inlines.Count == 0) return;
            ApplyParagraphState(_currentParagraph, state);
            state.PendingLegacyNumberingAfterReset.Clear();
            state.HasPendingLegacyNumberingAfterReset = false;
            state.ListText = null;
            state.PendingListTextAfterReset = null;
            if (_currentNote != null) {
                if (_currentParagraph.Inlines.Count > 0) {
                    CountSemanticBlock();
                    _currentNote.AddParsedParagraph(_currentParagraph);
                }

                _currentParagraph = new RtfParagraph();
                _currentParagraphIsInTable = false;
                return;
            }

            if (_currentHeaderFooter != null) {
                if (_currentParagraph.Inlines.Count > 0) {
                    CountSemanticBlock();
                    _currentHeaderFooter.AddParsedParagraph(_currentParagraph);
                }

                _currentParagraph = new RtfParagraph();
                _currentParagraphIsInTable = false;
                return;
            }

            if (_currentShape != null) {
                if (_currentParagraph.Inlines.Count > 0) {
                    CountSemanticBlock();
                    _currentShape.AddParsedTextBoxParagraph(_currentParagraph);
                }

                _currentParagraph = new RtfParagraph();
                _currentParagraphIsInTable = false;
                return;
            }

            if (_nestedTableContexts.Count > 0) {
                if (_currentParagraph.Inlines.Count > 0) {
                    CountSemanticBlock();
                    _nestedTableContexts[_nestedTableContexts.Count - 1].CurrentCellBlocks.Add(_currentParagraph);
                }

                _currentParagraph = new RtfParagraph();
                _currentParagraphIsInTable = true;
                return;
            }

            if (_currentParagraphIsInTable || _currentRow != null) {
                AddParagraphToCurrentCell(_currentParagraph);
                _currentParagraph = new RtfParagraph();
                _currentParagraphIsInTable = false;
                return;
            }

            if (_currentParagraph.Inlines.Count > 0 || (_document.Blocks.Count == 0 && _document.Paragraphs.Count == 0)) {
                AddDocumentBlock(_currentParagraph);
            }

            _currentParagraph = new RtfParagraph();
        }

        private static int ConsumeAnsiFallbackBytes(CharacterState state, int availableBytes) {
            int skip = Math.Min(state.SkipCharacters, availableBytes);
            state.SkipCharacters -= skip;
            return skip;
        }

        private void AppendUnicodeValue(int value, CharacterState state) {
            int unsigned = value < 0 ? value + 65536 : value;
            char codeUnit = (char)unsigned;

            if (char.IsHighSurrogate(codeUnit)) {
                if (state.PendingHighSurrogate.HasValue) {
                    state.PendingHighSurrogate = null;
                    AppendText("\uFFFD", state);
                }

                state.PendingHighSurrogate = codeUnit;
                return;
            }

            if (char.IsLowSurrogate(codeUnit)) {
                if (state.PendingHighSurrogate.HasValue) {
                    char highSurrogate = state.PendingHighSurrogate.Value;
                    state.PendingHighSurrogate = null;
                    AppendText(new string(new[] { highSurrogate, codeUnit }), state);
                } else {
                    AppendText("\uFFFD", state);
                }

                return;
            }

            if (state.PendingHighSurrogate.HasValue) {
                state.PendingHighSurrogate = null;
                AppendText("\uFFFD", state);
            }

            AppendText(codeUnit.ToString(), state);
        }

        private RtfImage? ReadPicture(RtfGroup group) {
            _limits.BeginImage(group.Position);
            RtfImageFormat format = RtfImageFormat.Unknown;
            int? sourceWidth = null;
            int? sourceHeight = null;
            int? desiredWidth = null;
            int? desiredHeight = null;
            var data = new List<byte>();
            long imageBytes = 0;

            foreach (RtfNode node in group.Children) {
                _limits.CheckCancellation();
                if (node is RtfControlWord control) {
                    switch (control.Name) {
                        case "pngblip":
                            format = RtfImageFormat.Png;
                            break;
                        case "jpegblip":
                            format = RtfImageFormat.Jpeg;
                            break;
                        case "dibitmap":
                            format = RtfImageFormat.Dib;
                            break;
                        case "wmetafile":
                            format = RtfImageFormat.Wmf;
                            break;
                        case "emfblip":
                            format = RtfImageFormat.Emf;
                            break;
                        case "picw":
                            sourceWidth = control.Parameter;
                            break;
                        case "pich":
                            sourceHeight = control.Parameter;
                            break;
                        case "picwgoal":
                            desiredWidth = control.Parameter;
                            break;
                        case "pichgoal":
                            desiredHeight = control.Parameter;
                            break;
                    }
                } else if (node is RtfBinary binary) {
                    _limits.AddImageBytes(ref imageBytes, binary.Data.Length, binary.Position);
                    data.AddRange(binary.Data);
                } else if (node is RtfText text) {
                    AppendHexBytes(text.Text, data, count => _limits.AddImageBytes(ref imageBytes, count, text.Position));
                }
            }

            if (data.Count == 0) {
                return null;
            }

            if (format == RtfImageFormat.Unknown && _options.WarnOnUnsupportedDestinations) {
                _diagnostics.Add(new RtfDiagnostic(RtfDiagnosticSeverity.Warning, "RTF102", "Picture payload was read, but the picture format is unknown.", group.Position));
            }

            return new RtfImage(format, data.ToArray()) {
                SourceWidth = sourceWidth,
                SourceHeight = sourceHeight,
                DesiredWidthTwips = desiredWidth,
                DesiredHeightTwips = desiredHeight
            };
        }

        private static void AppendHexBytes(string text, List<byte> data, Action<int>? beforeAppend = null) {
            int? highNibble = null;
            foreach (char ch in text) {
                int value = HexValue(ch);
                if (value < 0) continue;
                if (highNibble.HasValue) {
                    beforeAppend?.Invoke(1);
                    data.Add((byte)((highNibble.Value << 4) | value));
                    highNibble = null;
                } else {
                    highNibble = value;
                }
            }
        }

        private static int HexValue(char ch) {
            if (ch >= '0' && ch <= '9') return ch - '0';
            if (ch >= 'a' && ch <= 'f') return ch - 'a' + 10;
            if (ch >= 'A' && ch <= 'F') return ch - 'A' + 10;
            return -1;
        }

    }
}
