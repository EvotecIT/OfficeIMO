namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    /// <summary>
    /// Recognizes common inline TOC placeholders and converts them into a TocPlaceholderBlock so
    /// downstream rendering can generate a proper Table of Contents. Supported forms:
    ///   [TOC]
    ///   [[TOC]]
    ///   [TOC min=2 max=3 ordered=true layout=sidebar-right sticky=true scrollspy=true title="On this page"]
    ///   {:toc}
    ///   <!-- TOC --> or <!-- toc -->
    /// </summary>
    internal sealed class TocParser : IMarkdownBlockParser {
        public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            var raw = lines[i]; if (string.IsNullOrWhiteSpace(raw)) return false;
            var t = raw.Trim();

            // Simple markers
            if (t.Equals("[TOC]", System.StringComparison.OrdinalIgnoreCase) ||
                t.Equals("[[TOC]]", System.StringComparison.OrdinalIgnoreCase) ||
                t.Equals("[toc]", System.StringComparison.OrdinalIgnoreCase) ||
                t.Equals("[[toc]]", System.StringComparison.OrdinalIgnoreCase) ||
                t.Equals("{:toc}", System.StringComparison.OrdinalIgnoreCase) ||
                t.Equals("<!-- TOC -->", System.StringComparison.OrdinalIgnoreCase) ||
                t.Equals("<!-- toc -->", System.StringComparison.OrdinalIgnoreCase)) {
                var opts = new TocOptions();
                var ph = new TocPlaceholderBlock(opts);
                doc.Add(ph); i++; return true;
            }

            // Parameterized: [TOC key=value ...]
            if (t.StartsWith("[TOC", System.StringComparison.OrdinalIgnoreCase) && t.EndsWith("]")) {
                var inner = t.Substring(4, t.Length - 5).Trim(); // after [TOC and before ]
                var opts = new TocOptions();
                if (!string.IsNullOrWhiteSpace(inner)) {
                    try { ApplyAttributes(inner, opts); } catch { /* ignore malformed attributes; fall back to defaults */ }
                }
                // Clamp levels and sanitize options
                if (opts.MinLevel < 1) opts.MinLevel = TocOptions.DefaultMinLevel;
                if (opts.MaxLevel < opts.MinLevel) opts.MaxLevel = opts.MinLevel;
                if (opts.MaxLevel > 6) opts.MaxLevel = 6;
                if (opts.MinLevel > 6) opts.MinLevel = 6;
                if (opts.TitleLevel < 1) opts.TitleLevel = TocOptions.DefaultTitleLevel;
                if (opts.TitleLevel > 6) opts.TitleLevel = 6;
                if (opts.WidthPx.HasValue && opts.WidthPx.Value <= 0) opts.WidthPx = TocOptions.DefaultSidebarWidthPx;
                var ph = new TocPlaceholderBlock(opts);
                doc.Add(ph); i++; return true;
            }
            return false;
        }

        private static void ApplyAttributes(string inner, TocOptions o) {
            foreach (var tok in Tokenize(inner)) {
                int eq = tok.IndexOf('=');
                string key = eq > 0 ? tok.Substring(0, eq).Trim() : tok.Trim();
                string val = eq > 0 ? tok.Substring(eq + 1).Trim().Trim('"') : "true";
                switch (key.ToLowerInvariant()) {
                    case "min": if (int.TryParse(val, out var mi)) o.MinLevel = mi; break;
                    case "max": if (int.TryParse(val, out var ma)) o.MaxLevel = ma; break;
                    case "ordered": case "ol": if (Bool(val)) o.Ordered = true; break;
                    case "title": o.Title = val; o.IncludeTitle = !string.IsNullOrWhiteSpace(val); break;
                    case "titlelevel": if (int.TryParse(val, out var tl)) o.TitleLevel = tl; break;
                    case "layout":
                        var v = val.ToLowerInvariant();
                        if (v is "panel") o.Layout = TocLayout.Panel;
                        else if (v is "sidebar-right" or "sidebarright" or "right") o.Layout = TocLayout.SidebarRight;
                        else if (v is "sidebar-left" or "sidebarleft" or "left") o.Layout = TocLayout.SidebarLeft;
                        else o.Layout = TocLayout.List; break;
                    case "scrollspy": if (Bool(val)) o.ScrollSpy = true; break;
                    case "sticky": if (Bool(val)) o.Sticky = true; break;
                    case "aria": case "arialabel": o.AriaLabel = val; break;
                    case "width": case "widthpx": if (int.TryParse(val, out var w)) o.WidthPx = w; break;
                    case "chrome":
                        var c = val.ToLowerInvariant();
                        if (c == "none" || c == "no" || c == "ghost") o.Chrome = TocChrome.None;
                        else if (c == "outline") o.Chrome = TocChrome.Outline;
                        else if (c == "panel") o.Chrome = TocChrome.Panel;
                        else o.Chrome = TocChrome.Default; break;
                    case "hideonnarrow": if (Bool(val)) o.HideOnNarrow = true; break;
                    case "requiretoplevel": if (!Bool(val)) o.RequireTopLevel = false; break;
                    case "normalize": case "normalizetominlevel": if (!Bool(val)) o.NormalizeToMinLevel = false; break;
                    case "scope":
                        var sv = val.ToLowerInvariant();
                        if (sv.StartsWith("prev")) o.Scope = TocScope.PreviousHeading;
                        else if (sv.StartsWith("head")) o.Scope = TocScope.HeadingTitle; // requires scopeheading
                        else o.Scope = TocScope.Document;
                        break;
                    case "scopeheading": case "scopeheadingtitle": o.ScopeHeadingTitle = val; break;
                }
            }

            static bool Bool(string s) => s.Equals("true", System.StringComparison.OrdinalIgnoreCase) || s.Equals("1");
        }

        private static System.Collections.Generic.IEnumerable<string> Tokenize(string inner) {
            var tokens = new System.Collections.Generic.List<string>();
            if (string.IsNullOrWhiteSpace(inner)) return tokens;
            var sb = new System.Text.StringBuilder(); bool inQuotes = false;
            for (int i = 0; i < inner.Length; i++) {
                char ch = inner[i];
                if (ch == '"') { inQuotes = !inQuotes; sb.Append(ch); continue; }
                if (!inQuotes && char.IsWhiteSpace(ch)) { if (sb.Length > 0) { tokens.Add(sb.ToString()); sb.Clear(); } continue; }
                sb.Append(ch);
            }
            if (sb.Length > 0) tokens.Add(sb.ToString());
            return tokens;
        }
    }
}
