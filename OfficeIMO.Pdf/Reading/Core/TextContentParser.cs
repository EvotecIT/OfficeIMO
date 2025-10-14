using System.Globalization;
using System.Text;

namespace OfficeIMO.Pdf;

internal static class TextContentParser {
    public static List<PdfTextSpan> Parse(
        string content,
        System.Func<string, byte[], string> decodeWithFont,
        System.Func<string, byte[], double> sumWidth1000ForFont,
        bool adjustKerningFromTJ = true) {
        var spans = new List<PdfTextSpan>();
        // Text state
        bool inText = false;
        string font = "F1"; double size = 12; double x = 0, y = 0; double leading = size * 1.2; double charSpacing = 0, wordSpacing = 0; double hScale = 1.0; double textRise = 0;
        // Graphics state (CTM) and stack
        Matrix2D ctm = Matrix2D.Identity; var gstack = new System.Collections.Generic.Stack<Matrix2D>();
        // Operand buffer (tokens collected since last operator)
        var args = new List<object>(8);
        int i = 0; int n = content.Length;
        // Kerning state between text runs in TJ arrays (points) and rolling output buffer for gap checks
        double pendingGapPt = 0;
        var sbOutGlobal = new StringBuilder();
        while (i < n) {
            SkipWs(); if (i >= n) break;
            char c = content[i];
            if (c == '%') { // comment till end of line
                while (i < n && content[i] != '\n' && content[i] != '\r') i++;
                continue;
            }
            if (c == '/') { args.Add(ReadName()); continue; }
            if (c == '(') { args.Add(ReadLiteralStringBytes()); continue; }
            if (c == '<') {
                if (i + 1 < n && content[i + 1] == '<') { i += 2; continue; } // ignore dictionaries in content streams
                args.Add(ReadHexStringBytes()); continue;
            }
            if (c == '[') { args.Add(ReadArray()); continue; }
            if (c == ']' || c == '>') { i++; continue; }
            if (IsNumberStart(c)) { args.Add(ReadNumber()); continue; }
            // operator (BT, ET, Tf, Tm, Td, TD, T*, TL, Tc, Tw, Tz, Ts, cm, q, Q, Tj, TJ, ', ")
            string op = ReadOperator();
            if (op.Length == 0) { i++; continue; }

            switch (op) {
                case "BT": inText = true; pendingGapPt = 0; args.Clear(); break;
                case "ET": inText = false; pendingGapPt = 0; args.Clear(); break;
                case "Tf": if (args.Count >= 2) { size = ToDouble(args[args.Count - 1]); font = ToName(args[args.Count - 2]); args.Clear(); } break;
                case "Tm": if (args.Count >= 6) { x = ToDouble(args[args.Count - 2]); y = ToDouble(args[args.Count - 1]); args.Clear(); } break;
                case "Td": if (args.Count >= 2) { x += ToDouble(args[args.Count - 2]); y += ToDouble(args[args.Count - 1]); args.Clear(); } break;
                case "TD": if (args.Count >= 2) { double tx = ToDouble(args[args.Count - 2]); double ty = ToDouble(args[args.Count - 1]); x += tx; y += ty; leading = -ty; args.Clear(); } break;
                case "TL": if (args.Count >= 1) { leading = ToDouble(args[args.Count - 1]); args.Clear(); } break;
                case "T*": y -= leading; args.Clear(); break;
                case "Tc": if (args.Count >= 1) { charSpacing = ToDouble(args[args.Count - 1]); args.Clear(); } break;
                case "Tw": if (args.Count >= 1) { wordSpacing = ToDouble(args[args.Count - 1]); args.Clear(); } break;
                case "Tz": if (args.Count >= 1) { hScale = ToDouble(args[args.Count - 1]) / 100.0; args.Clear(); } break;
                case "Ts": if (args.Count >= 1) { textRise = ToDouble(args[args.Count - 1]); args.Clear(); } break;
                case "q": gstack.Push(ctm); args.Clear(); break;
                case "Q": ctm = gstack.Count > 0 ? gstack.Pop() : Matrix2D.Identity; args.Clear(); break;
                case "cm": if (args.Count >= 6) { var m2 = new Matrix2D(ToDouble(args[args.Count - 6]), ToDouble(args[args.Count - 5]), ToDouble(args[args.Count - 4]), ToDouble(args[args.Count - 3]), ToDouble(args[args.Count - 2]), ToDouble(args[args.Count - 1])); ctm = Matrix2D.Multiply(ctm, m2); args.Clear(); } break;
                case "'": // move to next line and show text
                    if (args.Count >= 1) { y -= leading; ShowTextRun(ToBytes(args[args.Count - 1])); pendingGapPt = 0; }
                    args.Clear();
                    break;
                case "\"": // set spacing and show text
                    if (args.Count >= 3) { wordSpacing = ToDouble(args[args.Count - 3]); charSpacing = ToDouble(args[args.Count - 2]); ShowTextRun(ToBytes(args[args.Count - 1])); pendingGapPt = 0; }
                    args.Clear();
                    break;
                case "Tj": if (args.Count >= 1) { ShowTextRun(ToBytes(args[args.Count - 1])); pendingGapPt = 0; args.Clear(); } break;
                case "TJ": if (args.Count >= 1) { ShowTextArray(args[args.Count - 1]); args.Clear(); } break;
                default: args.Clear(); break;
            }
        }
        return spans;

        // Helpers
        void MaybeInsertSpaceBeforeRun() {
            // Insert a space depending on kerning gap accumulated from TJ array numbers
            if (pendingGapPt <= 0) return;
            double prevAvg = Math.Max(1.0, size * 0.5); // fallback if we can't infer
            double emThreshold = size * 0.24; // about quarter em
            double glyphThreshold = prevAvg * 0.6;
            double threshold = Math.Max(emThreshold, glyphThreshold);
            // Tighten when previous char is wordish
            bool prevWord = sbOutGlobal.Length > 0 && (char.IsLetterOrDigit(sbOutGlobal[sbOutGlobal.Length - 1]) || sbOutGlobal[sbOutGlobal.Length - 1] == '\'' || sbOutGlobal[sbOutGlobal.Length - 1] == '-' || sbOutGlobal[sbOutGlobal.Length - 1] == '/');
            if (prevWord) threshold = Math.Min(threshold, 2.0);
            if (pendingGapPt >= threshold) sbOutGlobal.Append(' ');
            pendingGapPt = 0;
        }
        void ShowTextRun(byte[] bytes) {
            if (!inText || bytes == null || bytes.Length == 0) return;
            MaybeInsertSpaceBeforeRun();
            // Detect 2-byte CIDs (Identity-H) vs single-byte
            bool twoByte = false;
            if (bytes.Length >= 2) {
                string one = decodeWithFont(font, new byte[] { bytes[0] });
                string two = decodeWithFont(font, new byte[] { bytes[0], bytes[1] });
                twoByte = string.IsNullOrEmpty(one) && !string.IsNullOrEmpty(two);
            }
            var sbOut = new StringBuilder(bytes.Length);
            double advTotal = 0;
            char prevChar = '\0';
            for (int idx = 0; idx < bytes.Length;) {
                int step = twoByte ? (idx + 1 < bytes.Length ? 2 : 1) : 1;
                byte[] g = step == 1 ? new byte[] { bytes[idx] } : new byte[] { bytes[idx], bytes[idx + 1] };
                string t = decodeWithFont(font, g) ?? string.Empty;
                // Normalize common ligatures if ToUnicode is absent or produced U+FBxx
                if (t.Length > 0) {
                    t = t
                        .Replace("\uFB00", "ff") // ﬀ
                        .Replace("\uFB01", "fi") // ﬁ
                        .Replace("\uFB02", "fl") // ﬂ
                        .Replace("\uFB03", "ffi") // ﬃ
                        .Replace("\uFB04", "ffl"); // ﬄ
                }
                char ch = (t.Length > 0) ? t[0] : '\0';
                double w1000 = sumWidth1000ForFont(font, g);
                double advGlyph = ((w1000 / 1000.0) * size + charSpacing + (ch == ' ' ? wordSpacing : 0)) * hScale;
                // Drop thin spaces between letters/digits (visual join) but still advance
                double thinSpacePt = Math.Max(1.0, size * 0.12);
                bool dropSpace = false;
                if (ch == ' ') {
                    // Within-word space? Look ahead and behind
                    char nextChar = '\0';
                    int stepNext = twoByte ? (idx + step + 1 < bytes.Length ? 2 : 1) : 1;
                    if (idx + step < bytes.Length) {
                        byte[] gn = stepNext == 1 ? new byte[] { bytes[idx + step] } : new byte[] { bytes[idx + step], bytes[idx + step + 1] };
                        string tn = decodeWithFont(font, gn) ?? string.Empty;
                        if (tn.Length > 0) nextChar = tn[0];
                    }
                    bool prevWord = char.IsLetterOrDigit(prevChar) || prevChar=='\'' || prevChar=='-' || prevChar=='/';
                    bool nextWord = char.IsLetterOrDigit(nextChar) || nextChar=='\'' || nextChar=='-' || nextChar=='/';
                    if (prevWord && nextWord) {
                        double withinWordThreshold = Math.Max(size * 0.60, 1.6); // ~0.60em or >=1.6pt
                        if (advGlyph <= withinWordThreshold) dropSpace = true;
                    } else if (advGlyph <= thinSpacePt && prevChar != '\0') {
                        dropSpace = true;
                    }
                }
                if (dropSpace) {
                    // do not append, but keep advance
                } else if (ch != '\0') {
                    sbOut.Append(ch);
                    prevChar = ch;
                }
                advTotal += advGlyph;
                idx += step;
            }
            if (sbOut.Length == 0) return;
            string textOut = NormalizeShatteredSpan(sbOut.ToString());
            var (dx, dy) = ctm.Transform(x, y + textRise);
            spans.Add(new PdfTextSpan(textOut, font, size, dx, dy, advTotal));
            sbOutGlobal.Append(textOut);
            x += advTotal;
        }

        void ShowTextArray(object arrObj) {
            if (!inText || arrObj == null) return;
            var list = arrObj as List<object>;
            if (list == null) return;
            for (int j = 0; j < list.Count; j++) {
                var it = list[j];
                if (it is byte[] b) { ShowTextRun(b); }
                else if (adjustKerningFromTJ && it is double num) {
                    double delta = -num / 1000.0 * size * hScale;
                    x += delta;
                    // Only positive visual gap should suggest a space
                    if (delta > 0) pendingGapPt += delta; else pendingGapPt = 0;
                }
            }
        }

        void SkipWs() { while (i < n && char.IsWhiteSpace(content[i])) i++; }
        static bool IsDigit(char ch) => ch >= '0' && ch <= '9';
        bool IsNumberStart(char ch) => ch == '-' || ch == '+' || ch == '.' || IsDigit(ch);

        double ReadNumber() {
            int start = i; i++;
            while (i < n) { char ch = content[i]; if (!(IsDigit(ch) || ch == '.' || ch == 'E' || ch == 'e' || ch == '-' || ch == '+')) break; i++; }
            var s = content.Substring(start, i - start);
            if (!double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var v)) v = 0;
            return v;
        }

        string ReadName() {
            i++; int start = i;
            while (i < n) { char ch = content[i]; if (char.IsWhiteSpace(ch) || ch == '/' || ch == '[' || ch == ']' || ch == '(' || ch == ')' || ch == '<' || ch == '>') break; i++; }
            return content.Substring(start, i - start);
        }

        byte[] ReadLiteralStringBytes() {
            int start = ++i; int depth = 1; bool esc = false; var sb = new StringBuilder();
            while (i < n && depth > 0) {
                char ch = content[i++];
                if (esc) { sb.Append(ch); esc = false; }
                else if (ch == '\\') esc = true;
                else if (ch == '(') { depth++; sb.Append(ch); }
                else if (ch == ')') { depth--; if (depth > 0) sb.Append(ch); }
                else sb.Append(ch);
            }
            return PdfStringParser.ParseLiteralToBytes(sb.ToString());
        }

        byte[] ReadHexStringBytes() {
            i++; int start = i; while (i < n && content[i] != '>') i++; int end = i; if (i < n && content[i] == '>') i++;
            string hex = content.Substring(start, end - start);
            var sb = new StringBuilder(hex.Length);
            for (int k = 0; k < hex.Length; k++) { char ch = hex[k]; if (!char.IsWhiteSpace(ch)) sb.Append(ch); }
            hex = sb.ToString();
            if (hex.Length % 2 == 1) hex += "0";
            var bytes = new byte[hex.Length / 2];
            for (int k = 0; k < bytes.Length; k++) bytes[k] = byte.Parse(hex.Substring(k * 2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            return bytes;
        }

        List<object> ReadArray() {
            var list = new List<object>();
            i++; // skip [
            while (i < n) {
                SkipWs(); if (i >= n) break; char ch = content[i]; if (ch == ']') { i++; break; }
                if (ch == '(') { list.Add(ReadLiteralStringBytes()); continue; }
                if (ch == '<') { if (i + 1 < n && content[i + 1] == '<') { i += 2; continue; } list.Add(ReadHexStringBytes()); continue; }
                if (IsNumberStart(ch)) { list.Add(ReadNumber()); continue; }
                if (ch == '/') { list.Add(ReadName()); continue; }
                if (ch == '[') { i++; continue; } // ignore nested
                // unknown token inside array -> treat as operator and skip
                ReadOperator();
            }
            return list;
        }

        string ReadOperator() {
            int start = i; char ch = content[i++];
            if (ch == '\'' || ch == '"') return ch.ToString();
            while (i < n) {
                char c2 = content[i];
                if (char.IsWhiteSpace(c2) || c2 == '(' || c2 == '[' || c2 == '/' || c2 == '<' || c2 == '>') break;
                i++;
            }
            return content.Substring(start, i - start);
        }

        static double ToDouble(object o) { return o is double d ? d : 0.0; }
        static string ToName(object o) { return o as string ?? string.Empty; }
        static byte[] ToBytes(object o) { return o as byte[] ?? Array.Empty<byte>(); }

        // Helpers (left empty for future metrics)
        // NormalizeThinSpaces removed in favor of per-glyph join logic

        static bool Wordish(char c) => char.IsLetter(c) || char.IsDigit(c) || c == '\'' || c == '-' || c == '/';
        static bool AllLetters(string s) { for (int k = 0; k < s.Length; k++) if (!Wordish(s[k])) return false; return s.Length > 0; }
        static bool ShortAbbrev(string s) { if (s.Length == 0 || s.Length > 3) return false; for (int k = 0; k < s.Length; k++) if (!char.IsUpper(s[k])) return false; return true; }
        static string NormalizeShatteredSpan(string s) {
            if (string.IsNullOrEmpty(s)) return s;
            // Collapse whitespace sequences
            s = System.Text.RegularExpressions.Regex.Replace(s, "\\s+", " ").Trim();
            var parts = s.Split(' ');
            if (parts.Length <= 2) {
                if (parts.Length == 2 && AllLetters(parts[0]) && AllLetters(parts[1])) {
                    if (parts[0].Length == 1 && parts[1].Length >= 3) return parts[0] + parts[1];
                    if (parts[1].Length <= 2 || parts[0].Length <= 2) return parts[0] + parts[1];
                }
                return s;
            }
            int shortCount = parts.Count(p => p.Length <= 2 && AllLetters(p));
            if (!(shortCount >= 2 || shortCount * 4 >= parts.Length)) return s; // mostly healthy span
            var sb = new StringBuilder(s.Length);
            sb.Append(parts[0]);
            for (int ii = 1; ii < parts.Length; ii++) {
                string prev = parts[ii - 1]; string cur = parts[ii];
                bool upperSinglesJoin = prev.Length == 1 && cur.Length == 1 && char.IsUpper(prev[0]) && char.IsUpper(cur[0]);
                bool leadingLetterJoin = AllLetters(prev) && AllLetters(cur) && prev.Length == 1 && cur.Length >= 3;
                bool joinSmall = AllLetters(prev) && AllLetters(cur) && ((prev.Length <= 2 || cur.Length <= 2) || leadingLetterJoin || upperSinglesJoin) && !(ShortAbbrev(prev) && ShortAbbrev(cur) && !upperSinglesJoin);
                bool nextShort = (ii + 1 < parts.Length) && parts[ii + 1].Length <= 2 && AllLetters(parts[ii + 1]) && !ShortAbbrev(parts[ii + 1]);
                if (joinSmall || (AllLetters(cur) && cur.Length <= 2 && nextShort)) sb.Append(cur);
                else sb.Append(' ').Append(cur);
            }
            string joined = sb.ToString().Replace("  ", " ");
            // Secondary pass: join common suffix fragments
            var suffixes = new System.Collections.Generic.HashSet<string>(new [] { "ion","ions","ing","ment","tion","sion","iation","ization","ability","ality","able","ible","ance","ence","al","ally","er","ers","ed","ly","ology","ologies" });
            var toks = joined.Split(' ');
            if (toks.Length > 1) {
                var sb2 = new StringBuilder(joined.Length);
                sb2.Append(toks[0]);
                for (int i = 1; i < toks.Length; i++) {
                    string prev = toks[i - 1]; string cur = toks[i];
                    if (AllLetters(prev) && AllLetters(cur) && suffixes.Contains(cur.ToLowerInvariant())) sb2.Append(cur);
                    else sb2.Append(' ').Append(cur);
                }
                joined = sb2.ToString();
            }
            return joined;
        }
    }
}
