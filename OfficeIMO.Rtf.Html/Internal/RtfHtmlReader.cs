namespace OfficeIMO.Rtf.Html;

internal static class RtfHtmlReader {
    internal static RtfDocument Read(string html, RtfHtmlReadOptions options) {
        RtfDocument document = RtfDocument.Create();
        var context = new ReadContext(document, options);
        foreach (HtmlToken token in HtmlTokenizer.Tokenize(html)) {
            switch (token.Kind) {
                case HtmlTokenKind.Text:
                    context.AppendText(token.Value);
                    break;
                case HtmlTokenKind.StartTag:
                    context.Start(token);
                    if (token.SelfClosing) {
                        context.End(token.Value);
                    }

                    break;
                case HtmlTokenKind.EndTag:
                    context.End(token.Value);
                    break;
            }
        }

        context.TrimEmptyTrailingParagraph();
        return document;
    }

    private sealed class ReadContext {
        private readonly RtfDocument _document;
        private readonly RtfHtmlReadOptions _options;
        private readonly Stack<RtfListKind> _lists = new Stack<RtfListKind>();
        private readonly Stack<HtmlStyleScope> _styles = new Stack<HtmlStyleScope>();
        private RtfParagraph? _paragraph;
        private RtfTable? _table;
        private RtfTableRow? _row;
        private RtfTableCell? _cell;
        private Uri? _hyperlink;
        private int _bold;
        private int _italic;
        private int _underline;
        private int _strike;
        private int _superscript;
        private int _subscript;
        private int _preformatted;
        private bool _pageBreakAfterParagraph;

        internal ReadContext(RtfDocument document, RtfHtmlReadOptions options) {
            _document = document;
            _options = options;
        }

        internal void Start(HtmlToken token) {
            HtmlStyleDeclaration style = HtmlStyleDeclarationParser.Parse(GetAttribute(token, "style"));
            switch (token.Value) {
                case "p":
                case "div":
                case "section":
                case "article":
                case "blockquote":
                    StartParagraph();
                    ApplyParagraphStyle(style);
                    if (token.Value == "blockquote" && style.LeftIndentTwips == null) {
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
                    ApplyParagraphStyle(style);
                    EnsureParagraph().OutlineLevel = GetHeadingOutlineLevel(token.Value);
                    _bold++;
                    break;
                case "br":
                    EnsureParagraph().AddLineBreak();
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
                case "del":
                    _strike++;
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
                    _hyperlink = ReadUri(token, "href");
                    break;
                case "ul":
                    _lists.Push(RtfListKind.Bullet);
                    break;
                case "ol":
                    _lists.Push(RtfListKind.Decimal);
                    break;
                case "li":
                    StartParagraph();
                    EnsureParagraph().ListKind = _lists.Count == 0 ? RtfListKind.Bullet : _lists.Peek();
                    ApplyParagraphStyle(style);
                    break;
                case "table":
                    StartTable();
                    break;
                case "tr":
                    StartRow();
                    break;
                case "td":
                case "th":
                    StartCell();
                    break;
                case "img":
                    AddImage(token);
                    break;
                default:
                    if (_options.PreserveUnknownTagsAsText) {
                        AppendText("<" + token.Value + ">");
                    }

                    break;
            }

            if (style.HasInlineFormatting) {
                _styles.Push(new HtmlStyleScope(token.Value, style));
            }
        }

        internal void End(string name) {
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
                case "ul":
                case "ol":
                    if (_lists.Count > 0) {
                        _lists.Pop();
                    }

                    break;
                case "li":
                    EndParagraph();
                    break;
                case "td":
                case "th":
                    _paragraph = null;
                    _cell = null;
                    break;
                case "tr":
                    _row = null;
                    break;
                case "table":
                    _paragraph = null;
                    _cell = null;
                    _row = null;
                    _table = null;
                    break;
                default:
                    if (_options.PreserveUnknownTagsAsText) {
                        AppendText("</" + name + ">");
                    }

                    break;
            }

            PopStyleScope(name);
        }

        internal void AppendText(string text) {
            string value = _preformatted > 0 ? text : NormalizeWhitespace(text);
            if (value.Length == 0) {
                return;
            }

            if (_options.IgnoreInsignificantWhitespace && _preformatted == 0 && string.IsNullOrWhiteSpace(value)) {
                return;
            }

            RtfRun run = EnsureParagraph().AddText(value);
            run.Bold = ResolveStyleValue(style => style.Bold, _bold > 0);
            run.Italic = ResolveStyleValue(style => style.Italic, _italic > 0);
            run.Underline = ResolveStyleValue(style => style.Underline, _underline > 0);
            run.Strike = ResolveStyleValue(style => style.Strike, _strike > 0);
            run.VerticalPosition = ResolveVerticalPosition();
            run.Hyperlink = _hyperlink;
            RtfColor? foreground = ResolveStyleColor(style => style.ForegroundColor);
            RtfColor? background = ResolveStyleColor(style => style.BackgroundColor);
            if (foreground != null) {
                run.ForegroundColorIndex = GetOrAddColorIndex(foreground);
            }

            if (background != null) {
                run.CharacterBackgroundColorIndex = GetOrAddColorIndex(background);
            }

            string? fontFamily = ResolveStyleString(style => style.FontFamily);
            if (!string.IsNullOrWhiteSpace(fontFamily)) {
                run.FontId = _document.AddFont(fontFamily!);
            }

            double? fontSize = ResolveStyleNumber(style => style.FontSizePoints);
            if (fontSize.HasValue) {
                run.FontSize = fontSize.Value;
            }
        }

        internal void TrimEmptyTrailingParagraph() {
            _paragraph = null;
        }

        private void StartParagraph() {
            if (_paragraph != null && HasContent(_paragraph)) {
                EndParagraph();
            }

            _paragraph = _cell == null ? _document.AddParagraph() : _cell.AddParagraph();
            _pageBreakAfterParagraph = false;
        }

        private void EndParagraph() {
            if (_paragraph != null && _pageBreakAfterParagraph && !EndsWithPageBreak(_paragraph)) {
                _paragraph.AddPageBreak();
            }

            _paragraph = null;
            _pageBreakAfterParagraph = false;
        }

        private void ApplyParagraphStyle(HtmlStyleDeclaration style) {
            if (_paragraph != null && style.TextAlignment.HasValue) {
                _paragraph.Alignment = style.TextAlignment.Value;
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

            if (_paragraph != null && style.PageBreakBefore) {
                _paragraph.PageBreakBefore = true;
            }

            if (style.PageBreakAfter) {
                _pageBreakAfterParagraph = true;
            }
        }

        private RtfParagraph EnsureParagraph() {
            if (_paragraph == null) {
                StartParagraph();
            }

            return _paragraph!;
        }

        private void StartTable() {
            _table = _document.AddTable(0, 1);
            _row = null;
            _cell = null;
            _paragraph = null;
        }

        private void StartRow() {
            if (_table == null) {
                StartTable();
            }

            _row = _table!.AddRow();
            _cell = null;
            _paragraph = null;
        }

        private void StartCell() {
            if (_row == null) {
                StartRow();
            }

            int cellIndex = _row!.Cells.Count + 1;
            _cell = _row.AddCell(cellIndex * 2400);
            _paragraph = null;
        }

        private void AddImage(HtmlToken token) {
            string? source = GetAttribute(token, "src");
            if (string.IsNullOrWhiteSpace(source) || !TryReadDataImage(source!, out RtfImageFormat format, out byte[]? data)) {
                string? alt = GetAttribute(token, "alt");
                if (!string.IsNullOrWhiteSpace(alt)) {
                    AppendText(alt!);
                }

                return;
            }

            RtfImage image = _cell == null
                ? EnsureParagraph().AddImage(format, data!)
                : EnsureParagraph().AddImage(format, data!);
            image.Description = GetAttribute(token, "alt");
        }

        private Uri? ReadUri(HtmlToken token, string name) {
            string? value = GetAttribute(token, name);
            if (string.IsNullOrWhiteSpace(value)) {
                return null;
            }

            if (_options.BaseUri != null && Uri.TryCreate(_options.BaseUri, value, out Uri? resolved)) {
                return resolved;
            }

            return Uri.TryCreate(value, UriKind.RelativeOrAbsolute, out Uri? uri) ? uri : null;
        }

        private static string? GetAttribute(HtmlToken token, string name) {
            return token.Attributes.TryGetValue(name, out string? value) ? value : null;
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
                   paragraph.Inlines[paragraph.Inlines.Count - 1] is RtfBreak { Kind: RtfBreakKind.Page };
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

        private static bool TryReadDataImage(string source, out RtfImageFormat format, out byte[]? data) {
            format = RtfImageFormat.Unknown;
            data = null;
            const string prefix = "data:";
            int comma = source.IndexOf(',');
            if (!source.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) || comma < 0) {
                return false;
            }

            string media = source.Substring(prefix.Length, comma - prefix.Length).ToLowerInvariant();
            if (media.IndexOf(";base64", StringComparison.OrdinalIgnoreCase) < 0) {
                return false;
            }

            if (media.StartsWith("image/png", StringComparison.OrdinalIgnoreCase)) {
                format = RtfImageFormat.Png;
            } else if (media.StartsWith("image/jpeg", StringComparison.OrdinalIgnoreCase) || media.StartsWith("image/jpg", StringComparison.OrdinalIgnoreCase)) {
                format = RtfImageFormat.Jpeg;
            } else {
                return false;
            }

            try {
                data = Convert.FromBase64String(source.Substring(comma + 1));
                return true;
            } catch (FormatException) {
                format = RtfImageFormat.Unknown;
                data = null;
                return false;
            }
        }
    }

    private sealed class HtmlStyleScope {
        internal HtmlStyleScope(string name, HtmlStyleDeclaration style) {
            Name = name;
            Style = style;
        }

        internal string Name { get; }

        internal HtmlStyleDeclaration Style { get; }
    }
}
