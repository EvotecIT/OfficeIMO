using System.Globalization;
using System.Threading;

namespace OfficeIMO.Excel.Legacy;

/// <summary>Imports DIF row/value datasets without guessing formulas or cell formatting.</summary>
internal sealed class DifAdapter : TextSpreadsheetAdapterBase {
    public override LegacySpreadsheetFormat Format => LegacySpreadsheetFormat.Dif;
    public override string ProfileId => "dif-row-values";

    public override int Probe(byte[] data, string? sourceName, OfficeLegacyImportLimits limits, CancellationToken cancellationToken, out string reason) {
        cancellationToken.ThrowIfCancellationRequested();
        bool matches = TextSpreadsheetReader.HasHeader(data, "TABLE\r") || TextSpreadsheetReader.HasHeader(data, "TABLE\n");
        reason = matches ? "DIF TABLE topic signature." : "No DIF TABLE topic.";
        return matches ? 95 : 0;
    }

    protected override LegacySpreadsheetModel Parse(TextSpreadsheetReader reader) {
        if (reader.RequiredLine() != "TABLE") throw new InvalidDataException("DIF must start with its TABLE topic.");
        if (reader.RequiredLine() != "0,1") throw new InvalidDataException("Invalid DIF TABLE tuple.");
        string title = ReadText(reader);
        var model = new LegacySpreadsheetModel { Quality = OfficeLegacyImportQuality.Structured };
        if (title.Length > 0) {
            reader.RetainText(title);
            model.Metadata["SourceTableName"] = title;
        }
        var sheet = new LegacySpreadsheetSheet("Sheet1");
        model.Sheets.Add(sheet);
        int? vectors = null, tuples = null;
        var omittedTopics = new HashSet<string>(StringComparer.Ordinal);
        while (true) {
            string topic = reader.RequiredLine();
            (int type, string raw) = ParseTuple(reader.RequiredLine());
            if (type != 0) throw new InvalidDataException("Invalid DIF header tuple.");
            ReadText(reader);
            if (topic == "DATA") {
                if (raw != "0") throw new InvalidDataException("Invalid DIF DATA tuple.");
                break;
            }
            if (topic == "VECTORS" || topic == "TUPLES") {
                if (!int.TryParse(raw, NumberStyles.None, CultureInfo.InvariantCulture, out int count) || count > 1_048_576)
                    throw new InvalidDataException("DIF dimensions are outside the bounded worksheet profile.");
                if (topic == "VECTORS") { if (vectors.HasValue) throw new InvalidDataException("Duplicate DIF VECTORS topic."); vectors = count; }
                else { if (tuples.HasValue) throw new InvalidDataException("Duplicate DIF TUPLES topic."); tuples = count; }
            } else {
                if (omittedTopics.Count < 64) omittedTopics.Add(topic.Length > 16 ? topic.Substring(0, 16) : topic);
                else omittedTopics.Add("Additional topic kinds");
            }
        }
        if (!vectors.HasValue || !tuples.HasValue) throw new InvalidDataException("DIF requires VECTORS and TUPLES topics.");
        int row = 0, column = 0, maximumColumn = 0, errorCount = 0;
        while (true) {
            (int type, string raw) = ParseTuple(reader.RequiredLine());
            string payload = reader.RequiredLine();
            if (type == -1) {
                if (raw != "0") throw new InvalidDataException("Invalid DIF control tuple.");
                if (payload == "EOD") {
                    reader.RequireEnd();
                    if (vectors != maximumColumn || tuples != row) {
                        if (tuples == maximumColumn && vectors == row) model.Metadata["DimensionsProfile"] = "Transposed header counts; row order preserved";
                        else model.Findings.Add(Loss("DIF_DIMENSIONS_MISMATCH", "Structure", "DIF header dimensions differ from the recovered data; only explicitly present cells were imported in source row order."));
                    }
                    if (errorCount > 0) {
                        model.Metadata["ErrorValueCount"] = errorCount.ToString(CultureInfo.InvariantCulture);
                        model.Findings.Add(Loss("DIF_ERROR_AS_TEXT", "Cell", "DIF NA/ERROR markers were retained as literal text, without inventing a specific Excel error value."));
                    }
                    if (omittedTopics.Count > 0) model.Findings.Add(Loss("DIF_TOPICS_OMITTED", "Structure", "Unsupported DIF topics were omitted: " + string.Join(", ", omittedTopics.OrderBy(static item => item, StringComparer.Ordinal)) + "."));
                    return model;
                }
                if (payload != "BOT") throw new InvalidDataException("Unsupported DIF control marker.");
                if (++row > 1_048_576) throw new InvalidDataException("DIF exceeds Excel row bounds.");
                column = 0;
                continue;
            }
            if (row == 0) throw new InvalidDataException("DIF cell appears before the first BOT marker.");
            object value;
            if (type == 1) {
                if (raw != "0") throw new InvalidDataException("Invalid DIF string tuple.");
                value = ReadText(reader, payload);
            } else if (type == 0) {
                double number = TextSpreadsheetReader.Number(raw);
                switch (payload) {
                    case "V": value = number; break;
                    case "TRUE": if (number != 1) throw new InvalidDataException("Invalid DIF TRUE value."); value = true; break;
                    case "FALSE": if (number != 0) throw new InvalidDataException("Invalid DIF FALSE value."); value = false; break;
                    case "NA": case "ERROR": value = payload; errorCount++; break;
                    default: throw new InvalidDataException("Unsupported DIF numeric-value marker.");
                }
            } else throw new InvalidDataException("Unsupported DIF dataset type.");
            reader.AddCell(model, sheet, row, ++column, value);
            maximumColumn = Math.Max(maximumColumn, column);
        }
    }

    private static (int Type, string Value) ParseTuple(string line) {
        int comma = line.IndexOf(',');
        if (comma < 1 || !int.TryParse(line.Substring(0, comma), NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out int type))
            throw new InvalidDataException("Invalid DIF dataset tuple.");
        return (type, line.Substring(comma + 1));
    }

    private static string ReadText(TextSpreadsheetReader reader, string? firstLine = null) {
        string line = firstLine ?? reader.RequiredLine();
        if (!line.StartsWith("\"", StringComparison.Ordinal)) throw new InvalidDataException("DIF text must be quoted.");
        var text = new System.Text.StringBuilder(line);
        while (!HasClosingQuote(text)) {
            string continuation = reader.RequiredLine();
            if ((long)text.Length + continuation.Length + 1 > (long)reader.Limits.MaxTextCharacters + 2)
                throw new InvalidDataException("DIF text exceeds the configured text limit.");
            text.Append('\n').Append(continuation);
        }
        string value = text.ToString(1, text.Length - 2).Replace("\"\"", "\"");
        if (value.Length > reader.Limits.MaxTextCharacters) throw new InvalidDataException("DIF text exceeds the configured text limit.");
        return value;
    }

    private static bool HasClosingQuote(System.Text.StringBuilder text) {
        int quotes = 0;
        for (int index = text.Length - 1; index > 0 && text[index] == '"'; index--) quotes++;
        return (quotes & 1) == 1;
    }
}
