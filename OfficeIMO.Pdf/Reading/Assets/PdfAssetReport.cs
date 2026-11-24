using System.Globalization;
using System.Net;

namespace OfficeIMO.Pdf;

internal sealed class PdfAssetReport {
    private const string Schema = "officeimo.pdf.assets.v1";

    internal Table Pages { get; } = new("pages", "Page SVGs", ("page", "Page"), ("file", "File"));
    internal Table Boxes { get; } = new("pageBoxes", "Page boxes in millimetres",
        ("page", "Page"), ("box", "Box"), ("leftMm", "Left"), ("bottomMm", "Bottom"), ("rightMm", "Right"),
        ("topMm", "Top"), ("widthMm", "Width"), ("heightMm", "Height"));
    internal Table Images { get; } = new("images", "Images",
        ("page", "Page"), ("resource", "Resource"), ("objectNumber", "Object"), ("file", "File"),
        ("widthPixels", "Width (px)"), ("heightPixels", "Height (px)"), ("colorSpace", "Color space"),
        ("filter", "Filter"), ("renderingIntent", "Rendering intent"), ("maskKind", "Mask"),
        ("maskResolved", "Mask resolved"), ("rgbaPng", "RGBA PNG"), ("imageFile", "Image file"), ("sha256", "SHA-256"));
    internal Table Placements { get; } = new("imagePlacements", "Image placements",
        ("page", "Page"), ("resource", "Resource"), ("objectNumber", "Object"), ("file", "File"),
        ("effectiveDpiX", "X DPI"), ("effectiveDpiY", "Y DPI"));
    internal Table Fonts { get; } = new("fonts", "Fonts",
        ("name", "Base font"), ("family", "Family"), ("subtype", "PDF subtype"), ("embedded", "Embedded"),
        ("subset", "Subset"), ("subsetTag", "Subset tag"), ("programSubtype", "Program"),
        ("pages", "Pages"), ("file", "File"), ("nameTableRepaired", "Name repaired"), ("sourceSha256", "Source SHA-256"), ("sha256", "File SHA-256"));
    internal Table Profiles { get; } = new("iccProfiles", "ICC profiles",
        ("page", "Page"), ("source", "Source"), ("file", "File"), ("sizeBytes", "Bytes"), ("sha256", "SHA-256"));
    internal Table Inks { get; } = new("spotColorants", "Spot colorants",
        ("page", "Page"), ("colorSpace", "Color space"), ("name", "Name"));
    internal Table Diagnostics { get; } = new("diagnostics", "Diagnostics",
        ("page", "Page"), ("source", "Source"), ("code", "Code"), ("message", "Message"));

    private IEnumerable<Table> Tables {
        get {
            yield return Pages;
            yield return Boxes;
            yield return Images;
            yield return Placements;
            yield return Fonts;
            yield return Profiles;
            yield return Inks;
            yield return Diagnostics;
        }
    }

    internal void WriteJson(TextWriter writer, CancellationToken token) {
        writer.Write("{\n  \"schema\": \"" + Schema + "\"");
        var output = new StringBuilder();
        foreach (Table table in Tables) {
            writer.Write(",\n  \"" + table.Key + "\": [");
            for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
                token.ThrowIfCancellationRequested();
                output.Clear();
                output.Append(rowIndex == 0 ? "\n    {" : ",\n    {");
                object?[] row = table.Rows[rowIndex];
                for (int column = 0; column < row.Length; column++) {
                    if (column != 0) output.Append(", ");
                    output.Append('"').Append(table.Columns[column].Key).Append("\": ");
                    AppendJsonValue(output, row[column]);
                }
                writer.Write(output.Append('}').ToString());
            }
            if (table.Rows.Count != 0) writer.Write("\n  ");
            writer.Write(']');
        }
        writer.Write("\n}\n");
    }

    internal void WriteHtml(TextWriter writer, CancellationToken token) {
        writer.Write("""
            <!doctype html>
            <html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width">
            <title>PDF asset report</title>
            <style>
            body{font:16px system-ui,sans-serif;max-width:100rem;margin:2rem auto;padding:0 1rem;color:#172230}
            .table{overflow:auto}table{border-collapse:collapse;width:100%;font-size:.875rem}
            th,td{text-align:left;vertical-align:top;padding:.5rem;border-bottom:1px solid #d5dee7}
            th{background:#eef3f8}a{color:#174c84}section{margin-top:2rem}
            </style></head><body>
            <h1>PDF asset report</h1><p><a href="report.json">Download JSON report</a></p>
            """);
        var output = new StringBuilder();
        foreach (Table table in Tables) {
            writer.Write("<section><h2>" + table.Title + "</h2>");
            if (table.Rows.Count == 0) {
                writer.Write("<p>None found.</p></section>");
                continue;
            }
            writer.Write("<div class=\"table\"><table><thead><tr>");
            foreach (var column in table.Columns) writer.Write("<th scope=\"col\">" + column.Label + "</th>");
            writer.Write("</tr></thead><tbody>");
            foreach (object?[] row in table.Rows) {
                token.ThrowIfCancellationRequested();
                output.Clear();
                output.Append("<tr>");
                for (int column = 0; column < row.Length; column++) {
                    string value = WebUtility.HtmlEncode(Format(row[column]));
                    output.Append("<td>");
                    if (row[column] is null) output.Append('—');
                    else if (table.Columns[column].Key == "file") {
                        output.Append("<a href=\"").Append(value).Append("\">").Append(value).Append("</a>");
                    } else output.Append(value);
                    output.Append("</td>");
                }
                writer.Write(output.Append("</tr>").ToString());
            }
            writer.Write("</tbody></table></div></section>");
        }
        writer.Write("</body></html>\n");
    }

    private static void AppendJsonValue(StringBuilder output, object? value) {
        if (value is null) {
            output.Append("null");
        } else if (value is string text) {
            output.Append('"');
            foreach (char character in text) {
                if (character is '"' or '\\') output.Append('\\').Append(character);
                else if (character < ' ' || char.IsSurrogate(character)) {
                    output.Append("\\u").Append(((int)character).ToString("x4", CultureInfo.InvariantCulture));
                } else output.Append(character);
            }
            output.Append('"');
        } else if (value is int[] pages) {
            output.Append('[').Append(Format(pages)).Append(']');
        } else {
            output.Append(Format(value));
        }
    }

    private static string Format(object? value) => value switch {
        double number when double.IsNaN(number) || double.IsInfinity(number) => "null",
        double number => number.ToString("G", CultureInfo.InvariantCulture),
        bool flag => flag ? "true" : "false",
        int[] pages => string.Join(", ", pages.Select(static page => page.ToString(CultureInfo.InvariantCulture))),
        _ => Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty
    };

    internal sealed class Table {
        internal const int MaximumRows = 100_000;
        internal Table(string key, string title, params (string Key, string Label)[] columns) {
            Key = key;
            Title = title;
            Columns = columns;
        }

        internal string Key { get; }
        internal string Title { get; }
        internal (string Key, string Label)[] Columns { get; }
        internal List<object?[]> Rows { get; } = new();

        internal void Add(params object?[] values) {
            if (Rows.Count >= MaximumRows) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.UnderstandingArtifacts, MaximumRows, Rows.Count + 1);
            }
            Rows.Add(values);
        }
    }
}
