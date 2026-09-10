using System.Text;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Project;

/// <summary>MPX 4.0 record framing. Format-specific quoting is independent of field interpretation.</summary>
internal sealed class ProjectMpxRecords {
    internal char Separator { get; private set; }
    internal int CodePage { get; private set; }
    internal List<string[]> Records { get; } = new List<string[]>();

    internal static bool IsMpx(byte[] bytes) => bytes.Length >= 4 && bytes[0] == 'M' && bytes[1] == 'P' && bytes[2] == 'X';

    internal static bool HasAmbiguousDependencySeparator(char separator) => separator == '+' || separator == '-' || separator == '.' || separator == '%' || separator == '?';

    internal static ProjectMpxRecords Read(byte[] bytes, ProjectLoadOptions options, CancellationToken token) {
        if (!IsMpx(bytes)) throw new InvalidDataException("The MPX file creation record is missing.");
        char separator = (char)bytes[3];
        if (separator < 33 || separator > 126 || char.IsLetterOrDigit(separator) || separator == '"')
            throw new InvalidDataException("Invalid MPX list separator.");
        // The declaration itself is ASCII. Decode provisionally only to find its code page.
        int end = Array.FindIndex(bytes, b => b == 10 || b == 13);
        if (end < 0) end = bytes.Length;
        var first = Split(Encoding.ASCII.GetString(bytes, 0, end), separator, options, token).Single();
        if (first.Length < 3 || (first[2] != "4.0" && first[2] != "4.1"))
            throw new NotSupportedException("Qualified MPX input uses version 4.0 or 4.1.");
        int codePage = (first.Length < 4 ? "ANSI" : first[3].ToUpperInvariant()) switch {
            "" or "ANSI" => 1252, "437" => 437, "850" => 850, "MAC" => 10000,
            _ => throw new NotSupportedException("Unknown MPX code page declaration.")
        };
        if (bytes.Length > options.MaxCharacters) throw new InvalidDataException("MPX exceeds MaxCharacters.");
        string text = OfficeLegacySingleByteEncoding.Decode(bytes, 0, bytes.Length, codePage);
        var result = new ProjectMpxRecords { Separator = separator, CodePage = codePage };
        result.Records.AddRange(Split(text, separator, options, token));
        return result;
    }

    internal static IEnumerable<string[]> Split(string text, char separator, ProjectLoadOptions options, CancellationToken token) {
        var fields = new List<string>(); var field = new StringBuilder();
        bool quoted = false, closed = false, started = false; int fieldCount = 0, recordCount = 0;
        void FinishField() {
            if (++fieldCount > options.MaxElements) throw new InvalidDataException("MPX exceeds its field limit.");
            fields.Add(field.ToString().Trim(' ', '\t')); field.Clear(); closed = false; started = false;
        }
        for (int i = 0; i < text.Length; i++) {
            if ((i & 4095) == 0) token.ThrowIfCancellationRequested();
            char c = text[i];
            if (c == '\0') throw new InvalidDataException("MPX contains a NUL character.");
            if (quoted) {
                if (c == '"') {
                    if (i + 1 < text.Length && text[i + 1] == '"') { field.Append('"'); i++; }
                    else { quoted = false; closed = true; }
                } else if (c == '\r' || c == '\n') throw new InvalidDataException("MPX notes use ASCII 127, not a physical newline inside a field.");
                else field.Append(c);
                continue;
            }
            if (c == separator || c == '\r' || c == '\n') {
                FinishField();
                if (c != separator) {
                    if (c == '\r' && i + 1 < text.Length && text[i + 1] == '\n') i++;
                    if (fields.Count != 1 || fields[0].Length != 0) {
                        if (++recordCount > options.MaxElements) throw new InvalidDataException("MPX exceeds its record limit.");
                        yield return fields.ToArray();
                    }
                    fields.Clear();
                }
                continue;
            }
            if (c == '"') {
                if (started || closed) throw new InvalidDataException("Unexpected quotation mark in MPX field.");
                quoted = true; started = true; field.Clear(); continue;
            }
            if (closed && c != ' ' && c != '\t') throw new InvalidDataException("Content follows a closed MPX quoted field.");
            field.Append(c); if (c != ' ' && c != '\t') started = true;
        }
        if (quoted) throw new InvalidDataException("Unterminated MPX quoted field.");
        if (field.Length != 0 || fields.Count != 0 || closed) { FinishField(); yield return fields.ToArray(); }
    }

    internal static string WriteRecord(IEnumerable<string?> fields, char separator = ',') =>
        string.Join(separator.ToString(), fields.Select(value => {
            string text = value ?? "";
            if (text.IndexOfAny(new[] { '\r', '\n', '\0' }) >= 0) throw new InvalidDataException("MPX fields cannot contain physical newlines or NUL.");
            return text.IndexOf(separator) >= 0 || text.IndexOf('"') >= 0 ? "\"" + text.Replace("\"", "\"\"") + "\"" : text;
        })) + "\r\n";
}
