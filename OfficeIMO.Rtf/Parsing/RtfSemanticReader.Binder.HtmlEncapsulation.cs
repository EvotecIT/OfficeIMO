using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

internal static partial class RtfSemanticReader {
    private sealed partial class Binder {
        private RtfHtmlEncapsulation? ReadHtmlEncapsulation(RtfGroup root, int ansiCodePage, int unicodeSkipCount) {
            RtfControlWord? fromHtml = root.Children.OfType<RtfControlWord>()
                .FirstOrDefault(control => control.Name == "fromhtml");
            if (fromHtml == null) return null;

            var html = new StringBuilder();
            var state = new HtmlTextState {
                AnsiCodePage = ResolveFontCodePage(_document.Settings.DefaultFontId, ansiCodePage),
                DocumentCodePage = ansiCodePage,
                UnicodeSkipCount = unicodeSkipCount
            };
            AppendEncapsulatedHtml(root, html, state, isRoot: true);
            return new RtfHtmlEncapsulation(fromHtml.Parameter ?? 1, html.ToString());
        }

        private void AppendEncapsulatedHtml(RtfGroup group, StringBuilder html, HtmlTextState parent, bool isRoot = false) {
            _limits.CheckCancellation();
            string? destination = group.Destination;
            bool htmlTag = destination == "htmltag";
            if (!isRoot && !htmlTag && (RtfDestinationRegistry.IsIgnorableDestinationGroup(group) ||
                RtfDestinationRegistry.ShouldSkipSemanticBinding(destination) ||
                destination == "pict" || destination == "object" || destination == "shp" ||
                destination == "fldinst" || destination == "pn" ||
                TryGetHeaderFooterKind(destination).HasValue || TryGetNoteKind(destination).HasValue)) return;

            HtmlTextState state = parent.Clone();
            state.SkipCharacters = 0;
            if (htmlTag) {
                state.InHtmlTag = true;
                state.Suppressed = false;
                state.AnsiCodePage = state.DocumentCodePage;
            }
            foreach (RtfNode node in group.Children) {
                _limits.CheckCancellation();
                if (node is RtfGroup child) {
                    state.SkipCharacters = 0;
                    RtfGroup? unicodeAlternative = child.Destination == "upr" ? FindUnicodeAlternative(child) : null;
                    AppendEncapsulatedHtml(unicodeAlternative ?? child, html, state);
                    continue;
                }
                if (node is RtfControlWord control) {
                    if (!state.InHtmlTag && control.Name == "htmlrtf") {
                        state.Suppressed = control.Parameter != 0;
                        continue;
                    }
                    // Font selection remains meaningful even in a suppressed RTF approximation.
                    if (!state.InHtmlTag && control.Name == "f") {
                        state.FontId = control.Parameter;
                        state.AnsiCodePage = ResolveFontCodePage(state.FontId, state.DocumentCodePage);
                        continue;
                    }
                    if (state.Suppressed && !state.InHtmlTag) continue;
                    if (RtfTextDecoder.ConsumeFallback(state)) continue;
                    switch (control.Name) {
                        case "ansicpg" when !state.InHtmlTag && control.Parameter.HasValue:
                            state.DocumentCodePage = control.Parameter.Value;
                            state.AnsiCodePage = ResolveFontCodePage(state.FontId, state.DocumentCodePage);
                            break;
                        case "plain" when !state.InHtmlTag:
                            state.FontId = _document.Settings.DefaultFontId;
                            state.AnsiCodePage = ResolveFontCodePage(state.FontId, state.DocumentCodePage);
                            break;
                        case "uc" when control.Parameter >= 0:
                            state.UnicodeSkipCount = control.Parameter!.Value;
                            break;
                        case "u" when control.Parameter.HasValue:
                            AppendHtmlText(html, RtfTextDecoder.DecodeUnicode(control.Parameter.Value, state), state, node.Position);
                            break;
                        case "par":
                        case "line":
                            AppendHtmlText(html, RtfTextDecoder.Flush(state) + "\r\n", state, node.Position);
                            break;
                        case "tab":
                            AppendHtmlText(html, RtfTextDecoder.Flush(state) + "\t", state, node.Position);
                            break;
                        default:
                            if (IsSpecialCharacterControl(control.Name))
                                AppendHtmlText(html, RtfTextDecoder.Flush(state) + GetSpecialCharacterText(control.Name), state, node.Position);
                            break;
                    }
                    continue;
                }
                if (state.Suppressed && !state.InHtmlTag) continue;
                if (node is RtfText text) {
                    AppendHtmlText(html, RtfTextDecoder.DecodeAnsiText(text.Text, state), state, node.Position);
                } else if (node is RtfControlSymbol symbol) {
                    string decoded;
                    if (symbol.Symbol == '\'' && symbol.Parameter.HasValue) {
                        decoded = RtfTextDecoder.DecodeAnsiByte(symbol.Parameter.Value, state);
                    } else {
                        decoded = symbol.Symbol switch {
                            '\\' or '{' or '}' => RtfTextDecoder.DecodeLiteral(symbol.Symbol.ToString(), state),
                            '~' => RtfTextDecoder.DecodeLiteral("\u00A0", state),
                            '_' => RtfTextDecoder.DecodeLiteral("\u2011", state),
                            '-' => RtfTextDecoder.DecodeLiteral("\u00AD", state),
                            _ => string.Empty
                        };
                    }
                    AppendHtmlText(html, decoded, state, node.Position);
                } else if (node is RtfBinary) {
                    RtfTextDecoder.ConsumeFallback(state);
                }
            }
            if (!state.Suppressed || state.InHtmlTag)
                AppendHtmlText(html, RtfTextDecoder.Flush(state), state, group.Position);
        }

        private void AppendHtmlText(StringBuilder html, string text, HtmlTextState state, int position) {
            _limits.AddTextCharacters(text.Length, position);
            html.Append(state.InHtmlTag ? text : System.Net.WebUtility.HtmlEncode(text));
        }

        private sealed class HtmlTextState : RtfTextDecodingState {
            public int DocumentCodePage { get; set; }
            public int? FontId { get; set; }
            public bool InHtmlTag { get; set; }
            public bool Suppressed { get; set; }
            public HtmlTextState Clone() => (HtmlTextState)MemberwiseClone();
        }
    }
}
