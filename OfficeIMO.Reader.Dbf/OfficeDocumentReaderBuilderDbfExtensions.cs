namespace OfficeIMO.Reader.Dbf;

/// <summary>Adds DBF/xBase table support to an isolated Reader builder.</summary>
public static class OfficeDocumentReaderBuilderDbfExtensions {
    /// <summary>Stable DBF handler identifier.</summary>
    public const string HandlerId = "officeimo.reader.dbf";
    /// <summary>Adds bounded DBF path and stream ingestion using the DbaClientX codec.</summary>
    public static OfficeDocumentReaderBuilder AddDbfHandler(this OfficeDocumentReaderBuilder builder,
        ReaderDbfOptions? options = null, bool replaceExisting = true) {
        if (builder == null) throw new ArgumentNullException(nameof(builder));
        ReaderDbfOptions registered = (options ?? new()).Copy();
        return builder.AddHandler(new ReaderHandlerRegistration {
            Origin = ReaderHandlerOrigin.OfficeIMO, Id = HandlerId, DisplayName = "DBF Reader Adapter",
            Description = "Bounded DBF/xBase row projection through DbaClientX.",
            Kind = ReaderInputKind.Dbf, Extensions = new[] { ".dbf" }, UseDetectedKindFallback = false,
            DefaultMaxInputBytes = registered.ReadOptions.MaxInputBytes,
            SupportsIncrementalPath = true, SupportsIncrementalStream = true,
            // The table's bytes alone cannot identify memo-dependent content. Chunk hashes remain available.
            SourceHashBehavior = ReaderSourceHashBehavior.HandlerManaged,
            ReadPath = (path, reader, token) => DbfReaderAdapter.Read(path, reader, registered.Copy(), token),
            ReadStream = (stream, name, reader, token) => DbfReaderAdapter.Read(stream, name, reader, registered.Copy(), token),
            ExtensionValidationProbeStream = (stream, _, reader, token) => DbfReaderAdapter.Probe(stream, reader, registered.Copy(), token),
            WarningBehavior = ReaderWarningBehavior.Mixed,
            FormatQualifications = new[] {
                new ReaderFormatQualification(".dbf", "dbf", ReaderFormatSupport.ReadConvert,
                    "dBASE III-compatible, FoxPro 2 F5 and bounded Visual FoxPro field profiles",
                    preservation: new[] { "Column names and invariant row values; binary values as Base64 text" },
                    limitations: new[] { "No DBF writing, index/database traversal or object activation",
                        "Nulls become empty Reader cells; typed values remain available through DbfDataReader",
                        "Nonempty memos require opt-in path sidecars; no filesystem resolution for stream inputs",
                        "Source bundle byte hash is not supplied; chunk hashes cover the emitted projection" },
                    evidence: new[] { "DBAClientX.Dbf/README.md", "OfficeIMO.Reader.Dbf.Tests" })
            }
        }, replaceExisting);
    }
}
