using System.Globalization;
using System.Text;
using System.Text.Json;

namespace OfficeIMO.Notebook.Import;

internal static partial class NotebookScriptImporter {
    internal static void ReadJupytext(ReadOnlyMemory<byte> input, NotebookImportContext context) {
        string script = NotebookImportContext.DecodeUtf8(input);
        using var lines = new StringReader(script);
        string? line = lines.ReadLine();
        int lineNumber = 1;
        PercentHeader header = new(null, false, true);
        if (line?.TrimEnd() == "# ---") header = ReadPercentHeader(lines, context, ref line, ref lineNumber);
        string language = NormalizePercentLanguage(context.Options.ScriptLanguage ?? header.Language);
        SetLanguageInfo(context.Document, language);

        using var fenceLines = new StringReader(script);
        var fences = new PercentFenceIndex(fenceLines, context.CancellationToken);
        var framer = new PercentFramer(language, fences);
        var body = new StringBuilder();
        List<string?>? explicitIds = context.Options.PairedNotebook == null ? null : new List<string?>();
        PercentMarker? marker = null;
        int trailingEmpty = 0;
        int trailingBlank = 0;
        bool sawContent = false;
        bool sawMarker = false;
        bool hasUnparsedMetadata = false;
        while (line != null) {
            context.CancellationToken.ThrowIfCancellationRequested();
            if (framer.AtCellBoundary && TryMarker(line.AsSpan(), out ReadOnlySpan<char> options)) {
                if (sawContent || marker != null) AddPercentCell(context, marker, body, trailingEmpty, trailingBlank, language, explicitIds);
                marker = ParseMarker(options, context, lineNumber);
                hasUnparsedMetadata |= marker.Metadata.ContainsKey("officeimo_jupytext_marker_metadata");
                framer = new PercentFramer(language, fences, marker.Kind, explicitCell: true, uncomment: !IsActiveLanguage(marker.Metadata, language) ||
                    (marker.Kind == NotebookCellKind.Code && marker.Language != null &&
                        !SameLanguage(marker.Language, language)));
                body.Clear();
                trailingEmpty = 0;
                trailingBlank = 0;
                sawContent = false;
                sawMarker = true;
            } else if (sawContent || marker != null || !string.IsNullOrWhiteSpace(line)) {
                string content = framer.Read(line, lineNumber, header.CommentMagics);
                if (sawContent) body.Append('\n');
                body.Append(content);
                trailingEmpty = line.Length == 0 ? trailingEmpty + 1 : 0;
                trailingBlank = string.IsNullOrWhiteSpace(line) ? trailingBlank + 1 : 0;
                sawContent = true;
            }
            line = lines.ReadLine();
            lineNumber++;
        }

        if (sawContent || marker != null) AddPercentCell(context, marker, body, trailingEmpty, trailingBlank, language, explicitIds);
        if (!sawMarker && !header.DeclaresPercent) {
            throw new FormatException("The selected Jupytext percent format requires a percent cell marker or an explicit percent header.");
        }

        if (context.Options.PairedNotebook is { } companion) {
            if (hasUnparsedMetadata) ReportPairMismatch(context, "The script contains cell identity metadata that could not be verified.");
            else AttachCompanion(context, companion, explicitIds!);
        }
    }

    private static void AddPercentCell(NotebookImportContext context, PercentMarker? marker, StringBuilder body, int trailingEmpty,
        int trailingBlank, string language, List<string?>? explicitIds) {
        NotebookCellKind kind = marker?.Kind ?? NotebookCellKind.Code;
        int separatorLines = trailingBlank == 2 ? 2 : trailingEmpty > 0 ? 1 : 0;
        int end = body.Length;
        for (int i = 0; i < separatorLines && end > 0; i++) {
            while (end > 0 && body[end - 1] != '\n') end--;
            if (end > 0) end--;
        }

        string source = body.ToString(0, end);
        string? cellMarker = null;
        if (kind != NotebookCellKind.Code && language == "python") {
            source = UnwrapPercentText(source, out cellMarker);
        }
        string? cellLanguage = kind == NotebookCellKind.Code ? marker?.Language ?? language : null;
        if (cellLanguage != null && !SameLanguage(cellLanguage, language)) {
            if (IsPercentLanguageMagic(cellLanguage)) {
                string? arguments = marker is null ? null : MetadataString(marker.Metadata, "magic_args");
                source = "%%" + cellLanguage + (arguments == null ? "" : " " + arguments) + "\n" + source;
            } else {
                context.Report("notebook-import-script-language", "The mixed-language cell was retained as source without a custom Jupyter magic.",
                    "cells[" + context.Document.Cells.Count.ToString(CultureInfo.InvariantCulture) + "]", OfficeConversionLossKind.Approximation);
            }
        }
        NotebookCell cell = context.AddCell(kind, source, cellLanguage, marker?.Id);
        explicitIds?.Add(marker?.Id);
        if (marker != null) {
            foreach (KeyValuePair<string, JsonElement> property in marker.Metadata) cell.SetMetadata(property.Key, property.Value);
        }
        if (cellMarker != null) cell.SetMetadata("cell_marker", NotebookImportContext.JsonString(cellMarker));
    }

    private static string NormalizePercentLanguage(string? language) =>
        language?.Trim().ToLowerInvariant() switch {
            null or "" or "py" or "python" or "python2" or "python3" or "ipython" => "python",
            "r" or "ir" => "r",
            "julia" or "jl" => "julia",
            _ => throw new NotSupportedException("The Jupytext percent importer supports Python, R and Julia scripts.")
        };

    private static bool TryMarker(ReadOnlySpan<char> line, out ReadOnlySpan<char> options) {
        options = default;
        ReadOnlySpan<char> prefix = line.TrimStart();
        if (prefix.Length < 3 || prefix[0] != '#') return false;
        prefix = prefix.Slice(1).TrimStart();
        if (prefix.Length < 2 || prefix[0] != '%' || prefix[1] != '%') return false;
        int i = 2;
        while (i < prefix.Length && prefix[i] == '%') i++;
        if (i == prefix.Length ? i != 2 : !char.IsWhiteSpace(prefix[i])) return false;
        options = prefix.Slice(2).Trim();
        return true;
    }

    private sealed class PercentFramer(string language, PercentFenceIndex fences, NotebookCellKind kind = NotebookCellKind.Code,
        bool explicitCell = false, bool uncomment = false) {
        private char _quote;
        private int _quoteLength;
        private char _fence;
        private int _fenceLength;
        private int _juliaCommentDepth;
        private bool _magicContinues;

        internal bool AtCellBoundary => _quote == '\0' && _fence == '\0' && _juliaCommentDepth == 0;

        internal string Read(string line, int lineNumber, bool commentMagics) {
            if (kind != NotebookCellKind.Code && _quote == '\0') {
                ReadOnlySpan<char> text = RemoveComment(line.AsSpan());
                (char fence, int length, bool canClose) = ReadFence(text);
                if (_fence != '\0') {
                    if (canClose && fence == _fence && length >= _fenceLength) _fence = '\0';
                    return text.ToString();
                }
                if (fence != '\0' && text.Length != line.Length && fences.HasCloser(fence, length, lineNumber)) {
                    _fence = fence;
                    _fenceLength = length;
                }
                if (text.Length != line.Length) return text.ToString();
            }

            bool mayUnescape = AtCellBoundary;
            ReadQuotes(line.AsSpan());
            if (kind == NotebookCellKind.Code && mayUnescape) {
                if (uncomment) return RemoveComment(line.AsSpan()).ToString();
                return UnescapePercentCode(line, language, commentMagics, explicitCell, ref _magicContinues);
            }
            return line;
        }

        private void ReadQuotes(ReadOnlySpan<char> text) {
            bool julia = language == "julia";
            for (int i = 0; i < text.Length; i++) {
                char c = text[i];
                if (julia && _quote == '\0' && i + 1 < text.Length) {
                    if (c == '#' && text[i + 1] == '=') {
                        _juliaCommentDepth++;
                        i++;
                        continue;
                    }
                    if (_juliaCommentDepth > 0 && c == '=' && text[i + 1] == '#') {
                        _juliaCommentDepth--;
                        i++;
                        continue;
                    }
                }
                if (_juliaCommentDepth > 0) continue;
                if (_quote == '\0' && c == '#') break;
                if (c == '\\' && _quote != '\0') {
                    i++;
                    continue;
                }
                int delimiter = c is '\'' or '"' && language != "r" && i + 2 < text.Length && text[i + 1] == c && text[i + 2] == c
                    ? 3 : 1;
                if (_quote == '\0' && c is '\'' or '"' or '`') {
                    if (c == '`' && language == "python") continue;
                    if (julia && c == '\'') {
                        if (IsJuliaAdjoint(text, i)) continue;
                        // Julia also accepts an unescaped apostrophe as the one character in a literal.
                        if (delimiter == 3) {
                            i += 2;
                            continue;
                        }
                    }
                    _quote = c;
                    _quoteLength = delimiter;
                    i += delimiter - 1;
                } else if (c == _quote && delimiter >= _quoteLength) {
                    i += _quoteLength - 1;
                    _quote = '\0';
                }
            }
            if (language == "python" && _quoteLength == 1 && (text.IsEmpty || text[text.Length - 1] != '\\')) _quote = '\0';
        }

        private static bool IsJuliaAdjoint(ReadOnlySpan<char> text, int apostrophe) {
            if (apostrophe == 0) return false;
            if (text[apostrophe - 1] is ')' or ']' or '}' or '.' or '\'' or '"' or '`') return true;

            int start = apostrophe;
            while (start > 0 && IsJuliaIdentifierPart(text[start - 1])) start--;
            // A leading bang is an operator, while a trailing one may be part of the identifier.
            while (start < apostrophe && text[start] == '!') start++;
            if (start == apostrophe) return false;

            ReadOnlySpan<char> token = text.Slice(start, apostrophe - start);
            // Word operators can also be used as values or as qualified names.
            if (token is "in" or "isa" or "where") {
                return (start > 0 && text[start - 1] is '.' or ':' or '@')
                    || !StartsJuliaCharacterLiteral(text.Slice(apostrophe + 1));
            }
            // These keywords can introduce a character literal without a separator.
            return token is not ("begin" or "catch" or "else" or "elseif" or "finally"
                or "for" or "if" or "let" or "quote" or "return" or "try" or "while");
        }

        private static bool StartsJuliaCharacterLiteral(ReadOnlySpan<char> text) {
            if (text.Length < 2) return false;
            int closing = char.IsSurrogatePair(text[0], text[1]) ? 2 : 1;
            if (text[0] == '\\') {
                closing = 2;
                if (text[1] is 'x' or 'u' or 'U' or >= '0' and <= '7') {
                    // The longest Unicode escape has a prefix and eight hexadecimal digits.
                    while (closing < text.Length && closing < 10 && Uri.IsHexDigit(text[closing])) closing++;
                }
            }
            return closing < text.Length && text[closing] == '\'';
        }

        private static bool IsJuliaIdentifierPart(char c) => char.IsLetterOrDigit(c) || char.IsSurrogate(c)
            || c is '!' or '\u2032' or '\u2033' or '\u2034' or '\u2057'
            || char.GetUnicodeCategory(c) is UnicodeCategory.LetterNumber or UnicodeCategory.OtherNumber
                or UnicodeCategory.NonSpacingMark or UnicodeCategory.SpacingCombiningMark or UnicodeCategory.EnclosingMark
                or UnicodeCategory.ConnectorPunctuation or UnicodeCategory.CurrencySymbol
                or UnicodeCategory.ModifierSymbol or UnicodeCategory.OtherSymbol;
    }

    private static ReadOnlySpan<char> RemoveComment(ReadOnlySpan<char> line) {
        if (line.IsEmpty || line[0] != '#') return line;
        return line.Slice(line.Length > 1 && line[1] == ' ' ? 2 : 1);
    }

    private static string UnwrapPercentText(string source, out string? cellMarker) {
        cellMarker = null;
        ReadOnlySpan<char> text = source.AsSpan().Trim();
        int start = !text.IsEmpty && text[0] is 'r' or 'R' ? 1 : 0;
        if (text.Length < start + 6) return source;
        ReadOnlySpan<char> delimiter = text.Slice(start, 3);
        if ((!delimiter.SequenceEqual("\"\"\"".AsSpan()) && !delimiter.SequenceEqual("'''".AsSpan()))
            || !text.Slice(text.Length - 3).SequenceEqual(delimiter)) return source;
        ReadOnlySpan<char> content = text.Slice(start + 3, text.Length - start - 6);
        bool ownOpeningLine = !content.IsEmpty && content[0] == '\n';
        if (ownOpeningLine) content = content.Slice(1);
        bool ownClosingLine = !content.IsEmpty && content[content.Length - 1] == '\n';
        if (ownClosingLine) content = content.Slice(0, content.Length - 1);
        cellMarker = text.Slice(0, start + 3).ToString();
        if (!ownOpeningLine || !ownClosingLine) {
            cellMarker += (ownOpeningLine ? "\n," : ",") + (ownClosingLine ? "\n" : string.Empty) + delimiter.ToString();
        }
        return content.ToString();
    }

    private static string LineLocation(int line) => "line " + line.ToString(CultureInfo.InvariantCulture);
}
