namespace OfficeIMO.Reader.Publisher;

/// <summary>Adds native Publisher ingestion to an isolated Reader builder.</summary>
public static class OfficeDocumentReaderBuilderPublisherExtensions {
    /// <summary>Stable identifier for the native Publisher handler.</summary>
    public const string HandlerId = "officeimo.reader.publisher";

    /// <summary>Registers bounded PUB story, page-inventory and embedded-image recovery.</summary>
    public static OfficeDocumentReaderBuilder AddPublisherHandler(this OfficeDocumentReaderBuilder builder,
        ReaderPublisherOptions? options = null, bool replaceExisting = false) {
        if (builder == null) throw new ArgumentNullException(nameof(builder));
        ReaderPublisherOptions registered = (options ?? new ReaderPublisherOptions()).Clone();
        long maximumBytes = registered.ReadOptions!.Limits.MaxInputBytes;
        return builder.AddHandler(new ReaderHandlerRegistration {
            Origin = ReaderHandlerOrigin.OfficeIMO,
            Id = HandlerId, DisplayName = "Microsoft Publisher Reader Adapter",
            Description = "Native complete-story recovery, physical page inventory and embedded images with source losses.",
            Kind = ReaderInputKind.Publisher, Extensions = new[] { ".pub" },
            FormatQualifications = new[] { new ReaderFormatQualification(".pub", "Publisher.Native", ReaderFormatSupport.ReadConvert,
                "Selected Publisher 2002-and-later Contents/Quill generation",
                preservation: new[] { "Complete source stories", "Native bullet markers", "Physical page inventory", "Embedded image descriptors", "Source recovery diagnostics" },
                limitations: new[] { "Earlier native generations are rejected", "Page placement, tables and typography are omitted from semantic Reader output", "Native Publisher pixel fidelity is unqualified" },
                evidence: new[] { "OfficeIMO.Publisher/SUPPORT.md", "OfficeIMO.Reader.Tests/Reader.Publisher.cs" }) },
            DefaultMaxInputBytes = maximumBytes, MaxInputBytesCeiling = maximumBytes,
            SourceHashBehavior = ReaderSourceHashBehavior.HandlerManaged,
            ReadPath = (path, settings, token) => PublisherReaderAdapter.Read(path, settings, registered, token).Chunks,
            ReadStream = (stream, name, settings, token) => PublisherReaderAdapter.Read(stream, name, settings, registered, token).Chunks,
            ReadDocumentPath = (path, settings, token) => PublisherReaderAdapter.Read(path, settings, registered, token),
            ReadDocumentStream = (stream, name, settings, token) => PublisherReaderAdapter.Read(stream, name, settings, registered, token)
        }, replaceExisting);
    }
}
