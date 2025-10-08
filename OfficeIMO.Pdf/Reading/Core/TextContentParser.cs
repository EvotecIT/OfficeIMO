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
                case "BT": inText = true; args.Clear(); break;
                case "ET": inText = false; args.Clear(); break;
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
                    if (args.Count >= 1) { y -= leading; ShowText(ToBytes(args[args.Count - 1])); }
                    args.Clear();
                    break;
                case "\"": // set spacing and show text
                    if (args.Count >= 3) { wordSpacing = ToDouble(args[args.Count - 3]); charSpacing = ToDouble(args[args.Count - 2]); ShowText(ToBytes(args[args.Count - 1])); }
                    args.Clear();
                    break;
                case "Tj": if (args.Count >= 1) { ShowText(ToBytes(args[args.Count - 1])); args.Clear(); } break;
                case "TJ": if (args.Count >= 1) { ShowTextArray(args[args.Count - 1]); args.Clear(); } break;
                default: args.Clear(); break;
            }
        }
        return spans;

        // Helpers
        void ShowText(byte[] bytes) {
            if (!inText || bytes == null || bytes.Length == 0) return;
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
                string t = decodeWithFont(font, g);
                char ch = (t != null && t.Length > 0) ? t[0] : '\0';
                double w1000 = sumWidth1000ForFont(font, g);
                double advGlyph = ((w1000 / 1000.0) * size + charSpacing + (ch == ' ' ? wordSpacing : 0)) * hScale;
                // Drop thin spaces between letters/digits (visual join) but still advance
                bool dropSpace = (ch == ' ' && advGlyph <= 1.5 && prevChar != '\0');
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
            var (dx, dy) = ctm.Transform(x, y + textRise);
            spans.Add(new PdfTextSpan(sbOut.ToString(), font, size, dx, dy, advTotal));
            x += advTotal;
        }

        void ShowTextArray(object arrObj) {
            if (!inText || arrObj == null) return;
            var list = arrObj as List<object>;
            if (list == null) return;
            for (int j = 0; j < list.Count; j++) {
                var it = list[j];
                if (it is byte[] b) { ShowText(b); }
                else if (adjustKerningFromTJ && it is double num) { x += -num / 1000.0 * size * hScale; }
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
    }
}
