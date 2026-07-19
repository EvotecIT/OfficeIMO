using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;

namespace OfficeIMO.Reader;

/// <summary>
/// Dependency-light routing engine shared by all modular Reader packages.
/// </summary>
internal static partial class DocumentReaderEngine {
    private static readonly ReaderHandlerRegistrySnapshot EmptyHandlerRegistry =
        new ReaderHandlerRegistry().CaptureSnapshot();

    internal static IReadOnlyList<ReaderHandlerCapability> GetCapabilities(ReaderHandlerRegistrySnapshot snapshot) {
        if (snapshot == null) throw new ArgumentNullException(nameof(snapshot));
        return snapshot.Handlers
            .Select(static handler => handler.ToCapability())
            .OrderBy(static capability => capability.Id, StringComparer.Ordinal)
            .ToArray();
    }

    internal static ReaderCapabilityManifest GetCapabilityManifest(ReaderHandlerRegistrySnapshot snapshot) {
        return new ReaderCapabilityManifest {
            SchemaId = ReaderCapabilitySchema.Id,
            SchemaVersion = ReaderCapabilitySchema.Version,
            Handlers = GetCapabilities(snapshot)
        };
    }

    internal static ReaderInputKind DetectKind(string path) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        if (path.Length == 0) throw new ArgumentException("Path cannot be empty.", nameof(path));

        string extension = NormalizeExtension(ReaderLogicalPath.GetExtension(path));
        if (string.Equals(ReaderLogicalPath.GetFileName(path), "winmail.dat", StringComparison.OrdinalIgnoreCase)) {
            return ReaderInputKind.Email;
        }
        if (extension.Length > 0 && TryResolveCustomHandlerByExtension(extension, out ReaderHandlerDescriptor handler)) {
            return handler.Kind;
        }

        return extension switch {
            ".docx" or ".docm" or ".doc" => ReaderInputKind.Word,
            ".xlsx" or ".xlsm" or ".xls" => ReaderInputKind.Excel,
            ".pptx" or ".pptm" or ".ppt" or ".pot" or ".pps" => ReaderInputKind.PowerPoint,
            ".md" or ".markdown" => ReaderInputKind.Markdown,
            ".pdf" => ReaderInputKind.Pdf,
            ".eml" or ".msg" or ".oft" or ".mbox" or ".mbx" or ".tnef" or ".pst" or ".ost" or ".olm" or ".emlx" or ".oab" => ReaderInputKind.Email,
            ".ics" or ".ical" or ".ifb" or ".vcs" => ReaderInputKind.Calendar,
            ".vcf" or ".vcard" => ReaderInputKind.VCard,
            ".txt" or ".log" => ReaderInputKind.Text,
            ".csv" or ".tsv" => ReaderInputKind.Csv,
            ".json" => ReaderInputKind.Json,
            ".xml" => ReaderInputKind.Xml,
            ".yml" or ".yaml" => ReaderInputKind.Yaml,
            ".zip" => ReaderInputKind.Zip,
            ".one" or ".onepkg" or ".onetoc2" => ReaderInputKind.OneNote,
            ".rtf" => ReaderInputKind.Rtf,
            ".htm" or ".html" or ".mht" or ".mhtml" or ".xhtml" => ReaderInputKind.Html,
            ".odt" or ".ods" or ".odp" => ReaderInputKind.OpenDocument,
            ".adoc" or ".asc" or ".asciidoc" => ReaderInputKind.AsciiDoc,
            ".tex" => ReaderInputKind.Latex,
            _ => ReaderInputKind.Unknown
        };
    }

    /// <summary>
    /// Parses GitHub-style pipe tables without taking a dependency on a Markdown engine.
    /// Format-specific Markdown readers may provide richer projections.
    /// </summary>
    internal static IReadOnlyList<ReaderTable> ExtractMarkdownTables(
        string markdown,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        ReaderOptions effective = NormalizeOptions(options);
        string[] lines = (markdown ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        var tables = new List<ReaderTable>();
        for (int index = 0; index + 1 < lines.Length; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            string[] headers = SplitPipeRow(lines[index]);
            if (headers.Length == 0 || !IsPipeSeparator(lines[index + 1], headers.Length)) continue;

            var rows = new List<IReadOnlyList<string>>();
            int sourceRows = 0;
            int cursor = index + 2;
            while (cursor < lines.Length) {
                string[] cells = SplitPipeRow(lines[cursor]);
                if (cells.Length == 0) break;
                sourceRows++;
                if (rows.Count < effective.MaxTableRows) {
                    Array.Resize(ref cells, headers.Length);
                    for (int cell = 0; cell < cells.Length; cell++) cells[cell] ??= string.Empty;
                    rows.Add(cells);
                }
                cursor++;
            }

            string[] columns = headers.Select((header, column) => string.IsNullOrWhiteSpace(header)
                ? "Column" + (column + 1).ToString(CultureInfo.InvariantCulture)
                : header).ToArray();
            tables.Add(new ReaderTable {
                Kind = "markdown-table",
                Columns = columns,
                Rows = rows,
                TotalRowCount = sourceRows,
                Truncated = sourceRows > rows.Count,
                ColumnProfiles = ReaderTableProfiler.CreateProfiles(columns, rows),
                Location = new ReaderLocation { StartLine = index + 1, EndLine = Math.Max(index + 2, cursor) }
            });
            index = cursor - 1;
        }
        return tables;
    }

    private static string[] SplitPipeRow(string? line) {
        if (string.IsNullOrWhiteSpace(line) || line!.IndexOf('|') < 0) return Array.Empty<string>();
        string value = line.Trim();
        if (value.StartsWith("|", StringComparison.Ordinal)) value = value.Substring(1);
        if (value.EndsWith("|", StringComparison.Ordinal)) value = value.Substring(0, value.Length - 1);
        return value.Split('|').Select(static cell => cell.Trim()).ToArray();
    }

    private static bool IsPipeSeparator(string? line, int columnCount) {
        string[] cells = SplitPipeRow(line);
        if (cells.Length != columnCount) return false;
        return cells.All(static cell => {
            string value = cell.Trim().Trim(':');
            return value.Length >= 3 && value.All(static ch => ch == '-');
        });
    }

}
