using System.Text.Json;
using System.Text.RegularExpressions;
using System.Globalization;
using OfficeIMO.Pdf;

/// <summary>Cross-language PDF structure and text preservation against the browser's independent expectations.</summary>
internal static class PdfExportVerifier {
    internal static object Verify(string path) {
        using JsonDocument expected = JsonDocument.Parse(File.ReadAllText(path + ".json"));
        JsonElement contract = expected.RootElement;
        PdfReadDocument document = PdfReadDocument.Open(path);
        if (document.RepairReport.HasRepairs)
            throw new InvalidDataException("PDF required reader repairs: " + JsonSerializer.Serialize(document.RepairReport.Diagnostics));
        string[] texts = document.Pages.Select(p => string.Concat(p.GetTextSpans().Select(s => s.Text))).ToArray();
        string all = string.Concat(texts);
        foreach (JsonElement value in contract.GetProperty("required").EnumerateArray()) {
            string text = value.GetString()!;
            if (!all.Contains(text, StringComparison.Ordinal)) throw new InvalidDataException(Path.GetFileName(path) + " lost required text (length " + text.Length + "): " + text.Substring(0, Math.Min(text.Length, 100)) + "; extracted prefix: " + all.Substring(0, Math.Min(all.Length, 200)));
        }
        if (contract.TryGetProperty("forbidden", out JsonElement forbidden))
            foreach (JsonElement value in forbidden.EnumerateArray())
                if (all.Contains(value.GetString()!, StringComparison.Ordinal)) throw new InvalidDataException("PDF retained explicitly omitted metadata.");
        if (contract.TryGetProperty("repeated", out JsonElement repeated))
            foreach (string page in texts) foreach (JsonElement value in repeated.EnumerateArray())
                if (!page.Contains(value.GetString()!, StringComparison.Ordinal)) throw new InvalidDataException("A PDF page lost a repeated heading.");
        if (contract.TryGetProperty("firstPageOnly", out JsonElement firstPageOnly))
            foreach (JsonElement value in firstPageOnly.EnumerateArray()) {
                string heading = value.GetString()!;
                if (!texts[0].Contains(heading, StringComparison.Ordinal) || texts.Skip(1).Any(page => page.Contains(heading, StringComparison.Ordinal)))
                    throw new InvalidDataException("PDF repeated a table heading outside the table.");
            }
        if (contract.TryGetProperty("bodyText", out JsonElement bodyText)) {
            string body = all;
            foreach (JsonElement heading in contract.GetProperty("repeated").EnumerateArray()) body = body.Replace(heading.GetString()!, "", StringComparison.Ordinal);
            if (body != bodyText.GetString()) throw new InvalidDataException("PDF long text was lost, duplicated or reordered.");
        }
        if (contract.TryGetProperty("rows", out JsonElement rows)) {
            int count = rows.GetInt32();
            int[] ids = Regex.Matches(all, @"Row(\d{6})").Select(m => int.Parse(m.Groups[1].Value)).ToArray();
            if (!ids.SequenceEqual(Enumerable.Range(0, count))) throw new InvalidDataException("PDF row IDs were lost, duplicated or reordered.");
            if (contract.TryGetProperty("columns", out JsonElement columns)) {
                int width = columns.GetInt32(), ordinal = 0;
                foreach (string cell in document.Pages.SelectMany(p => p.GetTextSpans()).Select(s => s.Text).Where(t => !t.StartsWith("Column", StringComparison.Ordinal))) {
                    int row = ordinal / width, column = ordinal % width;
                    string wanted = column == 0 ? "Row" + row.ToString("D6", CultureInfo.InvariantCulture) : (row + column).ToString(CultureInfo.InvariantCulture);
                    if (cell != wanted) throw new InvalidDataException($"PDF cell {row}/{column} differs: {cell} != {wanted}.");
                    ordinal++;
                }
                if (ordinal != (long)count * width) throw new InvalidDataException("PDF cell count differs.");
            }
        }
        return new { file = Path.GetFileName(path), pages = document.Pages.Count, bytes = new FileInfo(path).Length, repairs = 0 };
    }
}
