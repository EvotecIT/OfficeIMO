using System.Globalization;
using System.Text;
using DBAClientX.Dbf;

namespace OfficeIMO.Reader.Dbf {
    internal static class DbfReaderAdapter {
        internal static IEnumerable<ReaderChunk> Read(string path, ReaderOptions readerOptions, ReaderDbfOptions options, CancellationToken token) {
            ApplyLimit(readerOptions, options);
            if (options.AllowMemoSidecarReads) {
                using var reader = DbfDataReader.Open(path, options.ReadOptions, token);
                foreach (ReaderChunk chunk in Project(reader, path, readerOptions, options, token)) yield return chunk;
            } else {
                using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
                using var reader = DbfDataReader.Open(stream, options.ReadOptions, cancellationToken: token);
                foreach (ReaderChunk chunk in Project(reader, path, readerOptions, options, token)) yield return chunk;
            }
        }

        internal static IEnumerable<ReaderChunk> Read(Stream stream, string? name, ReaderOptions readerOptions, ReaderDbfOptions options, CancellationToken token) {
            ApplyLimit(readerOptions, options);
            using var reader = DbfDataReader.Open(stream, options.ReadOptions, cancellationToken: token);
            foreach (ReaderChunk chunk in Project(reader, name ?? "document.dbf", readerOptions, options, token)) yield return chunk;
        }

        internal static bool Probe(Stream stream, ReaderOptions readerOptions, ReaderDbfOptions options, CancellationToken token) {
            ApplyLimit(readerOptions, options);
            try { using var reader = DbfDataReader.Open(stream, options.ReadOptions, cancellationToken: token); return true; }
            catch (InvalidDataException) { return false; }
            catch (NotSupportedException) { return false; }
        }

        private static void ApplyLimit(ReaderOptions reader, ReaderDbfOptions options) {
            if (reader.MaxInputBytes.HasValue)
                options.ReadOptions.MaxInputBytes = Math.Min(options.ReadOptions.MaxInputBytes, reader.MaxInputBytes.Value);
        }

        private static IEnumerable<ReaderChunk> Project(DbfDataReader reader, string path, ReaderOptions readerOptions, ReaderDbfOptions options, CancellationToken token) {
            string[] columns = reader.Schema.Columns.Select(column => column.Name).ToArray();
            int chunkRows = Math.Min(options.ChunkRows, Math.Max(1, readerOptions.MaxTableRows));
            int chunkIndex = 0, firstRecord = 0;
            var rows = new List<IReadOnlyList<string>>(chunkRows);
            var warnings = DefaultWarnings();
            while (reader.Read()) {
                token.ThrowIfCancellationRequested();
                if (rows.Count == 0) firstRecord = checked((int)reader.SourceRecordNumber) - 1;
                var values = new string[reader.FieldCount];
                for (int ordinal = 0; ordinal < values.Length; ordinal++) {
                    object value = reader.GetValue(ordinal);
                    values[ordinal] = Format(value, reader.Schema.Columns[ordinal]);
                    if (value == DBNull.Value) warnings.Add("DBF null values are represented as empty Reader cells.");
                    if (value is byte[]) warnings.Add("DBF binary values are represented as Base64 text and remain inert.");
                }
                if (reader.IsDeleted) warnings.Add("This chunk includes records marked deleted in the source DBF.");
                rows.Add(values);
                if (rows.Count < chunkRows) continue;
                yield return Build(path, reader.Schema, columns, rows, chunkIndex++, firstRecord, options.IncludeMarkdown, warnings);
                rows = new List<IReadOnlyList<string>>(chunkRows);
                warnings = DefaultWarnings();
            }
            if (rows.Count > 0 || chunkIndex == 0)
                yield return Build(path, reader.Schema, columns, rows, chunkIndex, firstRecord, options.IncludeMarkdown, warnings);
        }

        private static HashSet<string> DefaultWarnings() => new(StringComparer.Ordinal) {
            "DBF indexes, database backlinks and embedded object payloads are not interpreted.",
            "DBF source-bundle byte hash is unavailable; chunk hashes cover the emitted projection."
        };

        private static string Format(object value, DbfColumn column) => value switch {
            DBNull => string.Empty,
            byte[] bytes => Convert.ToBase64String(bytes),
            DateTime date => date.ToString(column.NativeType == 'D' ? "yyyy-MM-dd" : "yyyy-MM-dd'T'HH:mm:ss.fff", CultureInfo.InvariantCulture),
            double number => number.ToString("R", CultureInfo.InvariantCulture),
            bool boolean => boolean ? "true" : "false",
            IFormattable formatted => formatted.ToString(null, CultureInfo.InvariantCulture) ?? string.Empty,
            _ => value.ToString() ?? string.Empty
        };

        private static ReaderChunk Build(string path, DbfTableSchema schema, string[] columns, List<IReadOnlyList<string>> rows,
            int chunkIndex, int firstRecord, bool markdown, IEnumerable<string> warnings) {
            var location = new ReaderLocation { Path = path, BlockIndex = chunkIndex, SourceBlockIndex = firstRecord };
            var table = new ReaderTable {
                Title = Path.GetFileName(path), Kind = "dbf", Summary = $"{schema.Dialect}; {schema.EncodingName}",
                Location = location, Columns = columns, Rows = rows.ToArray(), TotalRowCount = rows.Count,
                ColumnProfiles = ReaderTableProfiler.CreateProfiles(columns, rows)
            };
            var text = new StringBuilder().AppendLine(string.Join("\t", columns));
            foreach (IReadOnlyList<string> row in rows) text.AppendLine(string.Join("\t", row));
            string? markdownText = null;
            if (markdown) {
                var builder = new StringBuilder();
                AppendMarkdownRow(builder, columns);
                AppendMarkdownRow(builder, columns.Select(_ => "---"));
                foreach (IReadOnlyList<string> row in rows) AppendMarkdownRow(builder, row);
                markdownText = builder.ToString();
            }
            return new ReaderChunk {
                Id = "dbf-" + chunkIndex.ToString("D4", CultureInfo.InvariantCulture), Kind = ReaderInputKind.Dbf,
                Location = location, Text = text.ToString(), Markdown = markdownText, Tables = new[] { table },
                Warnings = warnings.OrderBy(value => value, StringComparer.Ordinal).ToArray()
            };
        }

        private static void AppendMarkdownRow(StringBuilder builder, IEnumerable<string> cells) {
            builder.Append('|');
            foreach (string cell in cells) {
                builder.Append(' ');
                foreach (char value in cell) {
                    if (value == '\r') continue;
                    if (value == '\n') { builder.Append(" "); continue; }
                    if (value is >= '!' and <= '/' or >= ':' and <= '@' or >= '[' and <= '`' or >= '{' and <= '~') builder.Append('\\');
                    builder.Append(value);
                }
                builder.Append(" |");
            }
            builder.Append('\n');
        }
    }
}
