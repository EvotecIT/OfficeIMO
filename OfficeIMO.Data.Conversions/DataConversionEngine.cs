namespace OfficeIMO.Data.Conversions;

/// <summary>Opens sources and sinks by format and runs one conversion.</summary>
internal static class DataConversionEngine {
    /// <summary>Converts a stream between two hub formats. The output stream receives the result as it is produced.</summary>
    public static OfficeConversionReport Convert(Stream input, DataFormat source, Stream output, DataFormat target,
        DataConversionOptions? options, CancellationToken cancellationToken = default) {
        if (input == null) throw new ArgumentNullException(nameof(input));
        if (output == null) throw new ArgumentNullException(nameof(output));
        DataConversionOptions effective = Normalize(options, target);
        var context = new DataConversionContext(effective, cancellationToken);
        Stream? spooled = null;
        try {
            if (DataFormats.IsTable(target) && DataFormats.IsTable(source)) {
                throw new NotSupportedException(DataFormats.Name(source) + " to " + DataFormats.Name(target)
                    + " is not a data conversion; load the file with CsvDocument and save it with the target's CsvSaveOptions.");
            }

            if (DataFormats.IsTable(target) || DataFormats.IsTable(source) || source == DataFormat.Xml) {
                if (!input.CanSeek) {
                    spooled = Spool(input, cancellationToken);
                    input = spooled;
                }
            }

            if (DataFormats.IsTable(target)) {
                using StructuredDataReader rows = StructuredDataReader.Create(input, source, effective, owned: null, text: true);
                return CsvTableOutput.Write(rows, output, DataFormats.Name(source), target == DataFormat.Tsv);
            }

            long? documents = DocumentCount(source);
            if (NeedsDocumentCount(source, target)) {
                if (!input.CanSeek) {
                    spooled = Spool(input, cancellationToken);
                    input = spooled;
                }

                documents = CountDocuments(input, source, context);
            }

            using DataEventReader reader = OpenReader(input, source, context);
            using DataEventWriter writer = CreateWriter(output, target, context, documents, source);
            return Run(reader, writer, context, DataFormats.Name(source), DataFormats.Name(target));
        } finally {
            spooled?.Dispose();
        }
    }

    /// <summary>Applies the target's implied settings, such as JSON Lines for a <c>.jsonl</c> target.</summary>
    internal static DataConversionOptions Normalize(DataConversionOptions? options, DataFormat target) {
        if (target == DataFormat.JsonLines && options?.JsonLines != true) {
            DataConversionOptions copy = options?.Clone() ?? new DataConversionOptions();
            copy.JsonLines = true;
            return copy;
        }

        return options ?? DataConversionOptions.Default;
    }

    internal static OfficeConversionReport Run(DataEventReader reader, DataEventWriter writer, DataConversionContext context, string source, string target) {
        context.Source = reader;
        DataCopier.Copy(reader, writer);
        context.ThrowIfLossNotAllowed(source, target);
        return context.CreateReport(source, target);
    }

    private static bool NeedsDocumentCount(DataFormat source, DataFormat target) =>
        source == DataFormat.Yaml && target is DataFormat.Json or DataFormat.JsonLines;

    internal static long? DocumentCount(DataFormat source) => source switch {
        DataFormat.JsonLines => null,
        DataFormat.Yaml => null,
        _ => 1
    };

    private static long CountDocuments(Stream input, DataFormat source, DataConversionContext context) {
        long start = input.Position;
        try {
            if (source != DataFormat.Yaml) return 1;
            if (!LooksUtf8(input)) {
                input.Position = start;
                using var counting = new YamlEventReader(input, new DataConversionContext(context.Options, context.CancellationToken));
                long count = 0;
                while (counting.Read()) {
                    if (counting.TokenType == DataTokenType.StartDocument) count++;
                }

                return count;
            }

            input.Position = start;
            return YamlDocumentCounter.Count(input);
        } finally {
            input.Position = start;
        }
    }

    // UTF-16 and UTF-32 YAML have a zero byte among the first four (YAML 1.2.2, section 5.2).
    private static bool LooksUtf8(Stream input) {
        var head = new byte[4];
        int read = input.Read(head, 0, 4);
        for (int i = 0; i < read; i++) {
            if (head[i] == 0) return false;
        }

        return !(read >= 2 && ((head[0] == 0xFE && head[1] == 0xFF) || (head[0] == 0xFF && head[1] == 0xFE)));
    }

    private static Stream Spool(Stream input, CancellationToken cancellationToken) {
        string path = Path.Combine(Path.GetTempPath(), "officeimo-data-" + Guid.NewGuid().ToString("N") + ".tmp");
        var file = new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None, 81920, FileOptions.DeleteOnClose);
        try {
            byte[] buffer = ArrayPool<byte>.Shared.Rent(81920);
            try {
                int read;
                while ((read = input.Read(buffer, 0, buffer.Length)) > 0) {
                    cancellationToken.ThrowIfCancellationRequested();
                    file.Write(buffer, 0, read);
                }
            } finally {
                ArrayPool<byte>.Shared.Return(buffer);
            }

            file.Position = 0;
            return file;
        } catch {
            file.Dispose();
            throw;
        }
    }

    internal static DataEventReader OpenReader(Stream input, DataFormat format, DataConversionContext context) => format switch {
        DataFormat.Json => new JsonEventReader(input, jsonLines: false, context),
        DataFormat.JsonLines => new JsonEventReader(input, jsonLines: true, context),
        DataFormat.Yaml => new YamlEventReader(input, context),
        DataFormat.Xml => XmlEventReader.FromStream(input, context),
        DataFormat.Csv => TableEventReader.OpenCsv(input, context, tsv: false),
        DataFormat.Tsv => TableEventReader.OpenCsv(input, context, tsv: true),
        DataFormat.Toml => new TomlModelEventReader(LoadToml(input, context), context, owned: true),
        _ => throw new NotSupportedException("Reading " + format + " is not supported.")
    };

    internal static DataEventWriter CreateWriter(Stream output, DataFormat format, DataConversionContext context, long? documents, DataFormat source) => format switch {
        DataFormat.Json or DataFormat.JsonLines => new JsonEventWriter(output, documents, context, reportCombined: source == DataFormat.Yaml),
        DataFormat.Xml => XmlEventWriter.Create(output, documents, rootIsNamed: source == DataFormat.Xml, context),
        DataFormat.Yaml => new YamlEventWriter(output, context),
        DataFormat.Toml => new TomlEventWriter(output, context),
        _ => throw new NotSupportedException("Writing " + format + " is not supported.")
    };

    // TOML tables can be defined in any order and must be checked as a whole, so the source is loaded as a document.
    private static OfficeIMO.Toml.TomlDocument LoadToml(Stream input, DataConversionContext context) =>
        OfficeIMO.Toml.TomlDocument.Load(input, context.Options.TomlLoad);

    /// <summary>Writes through a temporary file next to the target and replaces the target only when the conversion succeeds.</summary>
    internal static OfficeConversionReport SaveToPath(string path, Func<Stream, OfficeConversionReport> write) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("A path is required.", nameof(path));
        string full = Path.GetFullPath(path);
        string directory = Path.GetDirectoryName(full) ?? ".";
        Directory.CreateDirectory(directory);
        string temporary = Path.Combine(directory, "." + Path.GetFileName(full) + "." + Guid.NewGuid().ToString("N") + ".tmp");
        OfficeConversionReport report;
        try {
            using (var stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None, 81920)) {
                report = write(stream);
            }

            if (File.Exists(full)) File.Delete(full);
            File.Move(temporary, full);
        } catch {
            try { if (File.Exists(temporary)) File.Delete(temporary); } catch (IOException) { } catch (UnauthorizedAccessException) { }
            throw;
        }

        return new OfficeConversionReport(report.Source, report.Target, full, report.ItemsWritten, report.Elapsed, report.Diagnostics);
    }
}
