using AngleSharp.Html.Dom;

namespace OfficeIMO.Html;

internal static partial class RtfHtmlReader {
    internal static RtfDocument Read(string html, HtmlToRtfOptions options) {
        RtfDocument document = RtfDocument.Create();
        ReadDom(html, options, document);
        return document;
    }

    internal static RtfDocument Read(IHtmlDocument htmlDocument, HtmlToRtfOptions options) {
        RtfDocument document = RtfDocument.Create();
        ReadDom(htmlDocument, options, document);
        return document;
    }

    private sealed partial class ReadContext {
        private readonly RtfDocument _document;
        private readonly HtmlToRtfOptions _options;
        private readonly Uri? _baseUri;
        private readonly Stack<HtmlListState> _lists = new Stack<HtmlListState>();
        private readonly Stack<HtmlStyleScope> _styles = new Stack<HtmlStyleScope>();
        private readonly Stack<RtfRevisionScope> _revisions = new Stack<RtfRevisionScope>();
        private readonly Stack<TableReadState> _tableStates = new Stack<TableReadState>();
        private readonly List<RowSpanState> _rowSpans = new List<RowSpanState>();
        private RtfParagraph? _paragraph;
        private RtfTable? _table;
        private RtfTableRow? _row;
        private RtfTableCell? _cell;
        private Uri? _hyperlink;
        private RtfRun? _lastRun;
        private RtfGeneratedText? _lastGeneratedText;
        private int _bold;
        private int _italic;
        private int _underline;
        private int _strike;
        private int _superscript;
        private int _subscript;
        private int _preformatted;
        private int _tableHead;
        private int _tableColumnIndex;
        private int _nextListId = 1;
        private RtfTextAlignment? _cellTextAlignment;
        private bool _pageBreakAfterParagraph;
        private int _headDepth;
        private int _titleDepth;
        private StringBuilder? _titleText;
        private RtfSection? _currentSection;
        private int _sectionElementDepth;

        internal ReadContext(RtfDocument document, HtmlToRtfOptions options, Uri? baseUri) {
            _document = document;
            _options = options;
            _baseUri = baseUri;
        }

        internal void Start(IElement token) {
            string name = token.LocalName;
            if (TryReadDocumentMetadata(token)) {
                return;
            }

            if (TryStartSection(token)) {
                return;
            }

            EnterSectionElement();

            if (TryReadNote(token)) {
                return;
            }

            if (TryReadObject(token) || TryReadShape(token)) {
                return;
            }

            if (TryReadGeneratedText(token)) {
                return;
            }

            bool startedField = TryStartField(token);
            if (!startedField) {
                EnterFieldElement();
            }

            HtmlStyleDeclaration style = HtmlStyleDeclarationParser.Parse(GetAttribute(token, "style"));
            style = ApplyLanguageDirectionAttributes(style, token);
            ApplyDocumentLanguageDirection(name, style);
            PushRevisionScope(token);
            switch (name) {
                case "p":
                case "div":
                case "section":
                case "article":
                case "blockquote":
                    StartParagraph();
                    ApplyParagraphStyleAttributes(token);
                    ApplyParagraphRevisionAttributes(token);
                    ApplyParagraphControlAttributes(token);
                    ApplyParagraphFrameAttributes(token);
                    ApplyParagraphStyle(style);
                    if (name == "blockquote" && style.LeftIndentTwips == null) {
                        EnsureParagraph().LeftIndentTwips = 720;
                    }

                    break;
                case "h1":
                case "h2":
                case "h3":
                case "h4":
                case "h5":
                case "h6":
                    StartParagraph();
                    ApplyParagraphStyleAttributes(token);
                    ApplyParagraphRevisionAttributes(token);
                    ApplyParagraphControlAttributes(token);
                    ApplyParagraphFrameAttributes(token);
                    ApplyParagraphStyle(style);
                    EnsureParagraph().OutlineLevel = GetHeadingOutlineLevel(name);
                    _bold++;
                    break;
                case "br":
                    AddBreak(token);
                    break;
                case "strong":
                case "b":
                    _bold++;
                    break;
                case "em":
                case "i":
                    _italic++;
                    break;
                case "u":
                    _underline++;
                    break;
                case "s":
                case "strike":
                    _strike++;
                    break;
                case "del":
                case "ins":
                    break;
                case "sup":
                    _superscript++;
                    break;
                case "sub":
                    _subscript++;
                    break;
                case "pre":
                case "code":
                    _preformatted++;
                    break;
                case "a":
                    if (!startedField) {
                        StartAnchor(token);
                    }

                    break;
                case "span":
                    break;
                case "ul":
                    _lists.Push(CreateListState(RtfListKind.Bullet));
                    break;
                case "ol":
                    _lists.Push(CreateListState(RtfListKind.Decimal));
                    break;
                case "thead":
                    _tableHead++;
                    break;
                case "li":
                    StartParagraph();
                    ApplyListAttributes(token);
                    ApplyParagraphStyleAttributes(token);
                    ApplyParagraphRevisionAttributes(token);
                    ApplyParagraphControlAttributes(token);
                    ApplyParagraphFrameAttributes(token);
                    ApplyParagraphStyle(style);
                    break;
                case "table":
                    StartTable();
                    break;
                case "tr":
                    StartRow(token, style);
                    break;
                case "td":
                    StartCell(token, style, isHeader: false);
                    break;
                case "th":
                    StartCell(token, style, isHeader: true);
                    break;
                case "img":
                    AddImage(token);
                    break;
                case "html":
                case "body":
                    break;
                default:
                    if (_options.PreserveUnknownTagsAsText) {
                        AppendText("<" + name + ">");
                    }

                    break;
            }

            int? styleId = IsInlineStyleScope(name) ? ReadStyleIdAttribute(token) : null;
            if (style.HasInlineFormatting || styleId.HasValue) {
                _styles.Push(new HtmlStyleScope(name, style, styleId));
            }
        }

        internal void End(string name) {
            if (EndDocumentMetadata(name)) {
                return;
            }

            if (TryEndSection(name)) {
                return;
            }

            switch (name) {
                case "p":
                case "div":
                case "section":
                case "article":
                case "blockquote":
                    EndParagraph();
                    break;
                case "h1":
                case "h2":
                case "h3":
                case "h4":
                case "h5":
                case "h6":
                    Decrement(ref _bold);
                    EndParagraph();
                    break;
                case "strong":
                case "b":
                    Decrement(ref _bold);
                    break;
                case "em":
                case "i":
                    Decrement(ref _italic);
                    break;
                case "u":
                    Decrement(ref _underline);
                    break;
                case "s":
                case "strike":
                case "del":
                    Decrement(ref _strike);
                    break;
                case "sup":
                    Decrement(ref _superscript);
                    break;
                case "sub":
                    Decrement(ref _subscript);
                    break;
                case "pre":
                case "code":
                    Decrement(ref _preformatted);
                    break;
                case "a":
                    _hyperlink = null;
                    break;
                case "span":
                    break;
                case "ul":
                case "ol":
                    if (_lists.Count > 0) {
                        _lists.Pop();
                    }

                    break;
                case "thead":
                    Decrement(ref _tableHead);
                    break;
                case "li":
                    EndParagraph();
                    break;
                case "td":
                    EndParagraph();
                    _cell = null;
                    _cellTextAlignment = null;
                    break;
                case "th":
                    Decrement(ref _bold);
                    EndParagraph();
                    _cell = null;
                    _cellTextAlignment = null;
                    break;
                case "tr":
                    EndRow();
                    _row = null;
                    break;
                case "table":
                    EndTable();
                    break;
                case "html":
                case "body":
                    break;
                default:
                    if (_options.PreserveUnknownTagsAsText) {
                        AppendText("</" + name + ">");
                    }

                    break;
            }

            PopStyleScope(name);
            PopRevisionScope(name);
            ExitFieldElement();
            ExitSectionElement();
        }

        internal void AppendText(string text) {
            if (AppendDocumentMetadataText(text)) {
                return;
            }

            string value = _preformatted > 0 ? text : NormalizeWhitespace(text);
            if (value.Length == 0) {
                return;
            }

            if (_options.IgnoreInsignificantWhitespace && _preformatted == 0 && string.IsNullOrWhiteSpace(value)) {
                return;
            }

            RtfRun run = EnsureInlineParagraph().AddText(value);
            _lastRun = run;
            _lastGeneratedText = null;
            run.Bold = ResolveStyleValue(style => style.Bold, _bold > 0);
            run.Italic = ResolveStyleValue(style => style.Italic, _italic > 0);
            bool underline = ResolveStyleValue(style => style.Underline, _underline > 0);
            run.Underline = underline;
            ApplyRichUnderline(run, underline);
            bool strike = ResolveStyleValue(style => style.Strike, _strike > 0);
            run.Strike = strike;
            ApplyRichStrike(run, strike);
            ApplyCharacterEffects(run);
            run.CapsStyle = ResolveCapsStyle();
            run.VerticalPosition = ResolveVerticalPosition();
            run.Direction = ResolveTextDirection();
            run.StyleId = ResolveRunStyleId();
            int? languageId = ResolveLanguageId();
            if (languageId.HasValue) {
                run.LanguageId = languageId.Value;
            }

            run.Hyperlink = _hyperlink;
            RtfColor? foreground = ResolveStyleColor(style => style.ForegroundColor);
            RtfColor? background = ResolveStyleColor(style => style.BackgroundColor);
            if (foreground != null) {
                run.ForegroundColorIndex = GetOrAddColorIndex(foreground);
            }

            if (background != null) {
                run.CharacterBackgroundColorIndex = GetOrAddColorIndex(background);
            }

            ApplyCharacterShading(run);

            HtmlBorderDeclaration? characterBorder = ResolveCharacterBorder();
            if (characterBorder != null) {
                ApplyCharacterBorder(run.CharacterBorder, characterBorder);
            }

            string? fontFamily = ResolveStyleString(style => style.FontFamily);
            if (!string.IsNullOrWhiteSpace(fontFamily)) {
                run.FontId = _document.AddFont(fontFamily!);
            }

            double? fontSize = ResolveStyleNumber(style => style.FontSizePoints);
            if (fontSize.HasValue) {
                run.FontSize = fontSize.Value;
            }

            ApplyCharacterMetrics(run);
            ApplyRevision(run);
        }

        internal void TrimEmptyTrailingParagraph() {
            _paragraph = null;
        }

        private void StartParagraph() {
            if (_paragraph != null && HasContent(_paragraph)) {
                EndParagraph();
            }

            _paragraph = _cell == null ? _document.AddParagraph() : _cell.AddParagraph();
            _lastRun = null;
            _lastGeneratedText = null;
            if (_cell == null) {
                AddSectionBlock(_paragraph);
            }

            if (_cellTextAlignment.HasValue) {
                _paragraph.Alignment = _cellTextAlignment.Value;
            }

            _pageBreakAfterParagraph = false;
        }

        private void EndParagraph() {
            if (_paragraph != null && _pageBreakAfterParagraph && !EndsWithPageBreak(_paragraph)) {
                _paragraph.AddPageBreak();
            }

            _paragraph = null;
            _lastRun = null;
            _lastGeneratedText = null;
            _pageBreakAfterParagraph = false;
        }

        private void ApplyParagraphStyle(HtmlStyleDeclaration style) {
            if (_paragraph != null && style.TextAlignment.HasValue) {
                _paragraph.Alignment = style.TextAlignment.Value;
            }

            if (_paragraph != null && style.Direction.HasValue) {
                _paragraph.Direction = style.Direction.Value;
            }

            if (_paragraph != null && style.LeftIndentTwips.HasValue) {
                _paragraph.LeftIndentTwips = style.LeftIndentTwips.Value;
            }

            if (_paragraph != null && style.RightIndentTwips.HasValue) {
                _paragraph.RightIndentTwips = style.RightIndentTwips.Value;
            }

            if (_paragraph != null && style.FirstLineIndentTwips.HasValue) {
                _paragraph.FirstLineIndentTwips = style.FirstLineIndentTwips.Value;
            }

            if (_paragraph != null && style.SpaceBeforeTwips.HasValue) {
                _paragraph.SpaceBeforeTwips = style.SpaceBeforeTwips.Value;
            }

            if (_paragraph != null && style.SpaceAfterTwips.HasValue) {
                _paragraph.SpaceAfterTwips = style.SpaceAfterTwips.Value;
            }

            if (_paragraph != null && style.LineSpacingTwips.HasValue) {
                _paragraph.LineSpacingTwips = style.LineSpacingTwips.Value;
                _paragraph.LineSpacingMultiple = style.LineSpacingMultiple;
            }

            if (_paragraph != null && style.BackgroundColor != null) {
                _paragraph.BackgroundColorIndex = GetOrAddColorIndex(style.BackgroundColor);
            }

            ApplyParagraphShading(style);

            if (_paragraph != null) {
                ApplyParagraphBorder(_paragraph.TopBorder, style.TopBorder);
                ApplyParagraphBorder(_paragraph.LeftBorder, style.LeftBorder);
                ApplyParagraphBorder(_paragraph.BottomBorder, style.BottomBorder);
                ApplyParagraphBorder(_paragraph.RightBorder, style.RightBorder);
            }

            if (_paragraph != null && style.PageBreakBefore) {
                _paragraph.PageBreakBefore = true;
            }

            if (style.PageBreakAfter) {
                _pageBreakAfterParagraph = true;
            }
        }

        private void ApplyParagraphBorder(RtfParagraphBorder target, HtmlBorderDeclaration? source) {
            if (source == null) {
                return;
            }

            target.Style = MapParagraphBorderStyle(source.Style);
            target.Width = source.Width;
            target.ColorIndex = source.Color == null ? null : GetOrAddColorIndex(source.Color);
        }

        private static RtfParagraphBorderStyle MapParagraphBorderStyle(RtfTableCellBorderStyle style) {
            switch (style) {
                case RtfTableCellBorderStyle.Double:
                    return RtfParagraphBorderStyle.Double;
                case RtfTableCellBorderStyle.Dotted:
                    return RtfParagraphBorderStyle.Dotted;
                case RtfTableCellBorderStyle.Dashed:
                    return RtfParagraphBorderStyle.Dashed;
                case RtfTableCellBorderStyle.None:
                    return RtfParagraphBorderStyle.None;
                default:
                    return RtfParagraphBorderStyle.Single;
            }
        }

        private RtfParagraph EnsureParagraph() {
            if (_paragraph == null) {
                StartParagraph();
            }

            return _paragraph!;
        }

        private Uri? ReadUri(IElement token, string name) {
            return ReadUriValue(GetAttribute(token, name));
        }

        private Uri? ReadUriValue(string? value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return null;
            }

            string resolved = HtmlUrlPolicyEvaluator.ResolveUrl(value, _baseUri, _options.UrlPolicy);
            if (string.IsNullOrWhiteSpace(resolved)) {
                return null;
            }

            return Uri.TryCreate(resolved, UriKind.RelativeOrAbsolute, out Uri? uri) ? uri : null;
        }

        private static string? GetAttribute(IElement token, string name) {
            return token.GetAttribute(name);
        }

        private static void Decrement(ref int value) {
            if (value > 0) {
                value--;
            }
        }

        private static int GetHeadingOutlineLevel(string name) {
            return name.Length == 2 && name[0] == 'h' && name[1] >= '1' && name[1] <= '6'
                ? name[1] - '1'
                : 0;
        }

        private void PopStyleScope(string name) {
            if (_styles.Count == 0) {
                return;
            }

            var deferred = new List<HtmlStyleScope>();
            while (_styles.Count > 0) {
                HtmlStyleScope scope = _styles.Pop();
                if (string.Equals(scope.Name, name, StringComparison.OrdinalIgnoreCase)) {
                    break;
                }

                deferred.Add(scope);
            }

            for (int index = deferred.Count - 1; index >= 0; index--) {
                _styles.Push(deferred[index]);
            }
        }

        private bool ResolveStyleValue(Func<HtmlStyleDeclaration, bool?> selector, bool fallback) {
            foreach (HtmlStyleScope scope in _styles) {
                bool? value = selector(scope.Style);
                if (value.HasValue) {
                    return value.Value;
                }
            }

            return fallback;
        }

        private RtfVerticalPosition ResolveVerticalPosition() {
            foreach (HtmlStyleScope scope in _styles) {
                if (scope.Style.VerticalPosition.HasValue) {
                    return scope.Style.VerticalPosition.Value;
                }
            }

            if (_superscript > 0) {
                return RtfVerticalPosition.Superscript;
            }

            return _subscript > 0 ? RtfVerticalPosition.Subscript : RtfVerticalPosition.Baseline;
        }

        private RtfColor? ResolveStyleColor(Func<HtmlStyleDeclaration, RtfColor?> selector) {
            foreach (HtmlStyleScope scope in _styles) {
                RtfColor? value = selector(scope.Style);
                if (value != null) {
                    return value;
                }
            }

            return null;
        }

        private string? ResolveStyleString(Func<HtmlStyleDeclaration, string?> selector) {
            foreach (HtmlStyleScope scope in _styles) {
                string? value = selector(scope.Style);
                if (!string.IsNullOrWhiteSpace(value)) {
                    return value;
                }
            }

            return null;
        }

        private double? ResolveStyleNumber(Func<HtmlStyleDeclaration, double?> selector) {
            foreach (HtmlStyleScope scope in _styles) {
                double? value = selector(scope.Style);
                if (value.HasValue) {
                    return value.Value;
                }
            }

            return null;
        }

        private int? ResolveRunStyleId() {
            foreach (HtmlStyleScope scope in _styles) {
                if (scope.StyleId.HasValue) {
                    return scope.StyleId.Value;
                }
            }

            return null;
        }

        private int GetOrAddColorIndex(RtfColor color) {
            for (int index = 0; index < _document.Colors.Count; index++) {
                RtfColor existing = _document.Colors[index];
                if (existing.Red == color.Red &&
                    existing.Green == color.Green &&
                    existing.Blue == color.Blue &&
                    existing.ThemeColor == color.ThemeColor &&
                    existing.Tint == color.Tint &&
                    existing.Shade == color.Shade) {
                    return index + 1;
                }
            }

            return _document.AddColor(color.Red, color.Green, color.Blue);
        }

        private static bool HasContent(RtfParagraph paragraph) {
            return paragraph.Inlines.Count > 0 || paragraph.Runs.Count > 0;
        }

        private static bool EndsWithPageBreak(RtfParagraph paragraph) {
            return paragraph.Inlines.Count > 0 &&
                   paragraph.Inlines[paragraph.Inlines.Count - 1] is RtfBreak rtfBreak &&
                   (rtfBreak.Kind == RtfBreakKind.Page || rtfBreak.Kind == RtfBreakKind.SoftPage);
        }

        private static string NormalizeWhitespace(string text) {
            if (text.Length == 0) {
                return string.Empty;
            }

            var builder = new StringBuilder(text.Length);
            bool lastWasWhitespace = false;
            foreach (char character in text) {
                if (char.IsWhiteSpace(character)) {
                    if (!lastWasWhitespace) {
                        builder.Append(' ');
                        lastWasWhitespace = true;
                    }
                } else {
                    builder.Append(character);
                    lastWasWhitespace = false;
                }
            }

            return builder.ToString();
        }

    }

    private sealed class HtmlStyleScope {
        internal HtmlStyleScope(string name, HtmlStyleDeclaration style, int? styleId) {
            Name = name;
            Style = style;
            StyleId = styleId;
        }

        internal string Name { get; }

        internal HtmlStyleDeclaration Style { get; }

        internal int? StyleId { get; }
    }
}
