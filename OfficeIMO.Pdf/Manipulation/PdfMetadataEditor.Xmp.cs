namespace OfficeIMO.Pdf;

internal static partial class PdfMetadataEditor {
    /// <summary>
    /// Creates a normalized full-rewrite PDF whose Info dictionary and XMP packet share the supplied common fields.
    /// Null values preserve the existing Info/XMP value and empty strings clear it. Existing custom XMP schemas are preserved.
    /// </summary>
    public static byte[] SynchronizeMetadata(
        byte[] pdf,
        string? title = null,
        string? author = null,
        string? subject = null,
        string? keywords = null,
        bool createXmpMetadata = true) {
        Guard.NotNull(pdf, nameof(pdf));
        _ = PdfMutationPlanner.RequireFullRewrite(pdf, PdfMutationOperation.SynchronizeMetadata);

        PdfDocumentInfo documentInfo = PdfInspector.Inspect(pdf);
        PdfMetadata existing = documentInfo.Metadata;
        PdfXmpMetadataInfo? existingXmp = documentInfo.XmpMetadata;
        var updated = new PdfMetadata {
            Title = title ?? existing.Title ?? existingXmp?.Title,
            Author = author ?? existing.Author ?? existingXmp?.Creator,
            Subject = subject ?? existing.Subject ?? existingXmp?.Description,
            Keywords = keywords ?? existing.Keywords ?? existingXmp?.Keywords
        };

        return PdfDocumentObjectGraphRewriter.Rewrite(
            pdf,
            sourceReadOptions: null,
            outputEncryption: null,
            (objects, security) => SynchronizeObjectGraph(objects, security, existingXmp, updated, createXmpMetadata));
    }

    /// <summary>Synchronizes Info and XMP metadata from the current position of a readable stream.</summary>
    public static byte[] SynchronizeMetadata(
        Stream stream,
        string? title = null,
        string? author = null,
        string? subject = null,
        string? keywords = null,
        bool createXmpMetadata = true) {
        return SynchronizeMetadata(ReadStream(stream, nameof(stream)), title, author, subject, keywords, createXmpMetadata);
    }

    /// <summary>Writes a PDF with synchronized Info and XMP metadata to a writable stream.</summary>
    public static void SynchronizeMetadata(
        byte[] pdf,
        Stream outputStream,
        string? title = null,
        string? author = null,
        string? subject = null,
        string? keywords = null,
        bool createXmpMetadata = true) {
        WriteOutput(outputStream, SynchronizeMetadata(pdf, title, author, subject, keywords, createXmpMetadata));
    }

    /// <summary>Writes a PDF with synchronized Info and XMP metadata from one stream to another.</summary>
    public static void SynchronizeMetadata(
        Stream inputStream,
        Stream outputStream,
        string? title = null,
        string? author = null,
        string? subject = null,
        string? keywords = null,
        bool createXmpMetadata = true) {
        WriteOutput(outputStream, SynchronizeMetadata(inputStream, title, author, subject, keywords, createXmpMetadata));
    }

    /// <summary>Writes a PDF with synchronized Info and XMP metadata to a file.</summary>
    public static void SynchronizeMetadata(
        string inputPath,
        string outputPath,
        string? title = null,
        string? author = null,
        string? subject = null,
        string? keywords = null,
        bool createXmpMetadata = true) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNull(outputPath, nameof(outputPath));
        WriteOutput(
            ValidateOutputPath(outputPath),
            SynchronizeMetadata(File.ReadAllBytes(inputPath), title, author, subject, keywords, createXmpMetadata));
    }

    /// <summary>Creates a PDF with synchronized Info and XMP metadata from a file.</summary>
    public static byte[] SynchronizeMetadataToBytes(
        string inputPath,
        string? title = null,
        string? author = null,
        string? subject = null,
        string? keywords = null,
        bool createXmpMetadata = true) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        return SynchronizeMetadata(File.ReadAllBytes(inputPath), title, author, subject, keywords, createXmpMetadata);
    }

    private static int? SynchronizeObjectGraph(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDocumentSecurityInfo security,
        PdfXmpMetadataInfo? existingXmp,
        PdfMetadata updated,
        bool createXmpMetadata) {
        int infoObjectNumber = ReplaceInfoDictionary(objects, security, updated);
        PdfDictionary catalog = RequireCatalog(objects, security);
        if (catalog.Items.TryGetValue("Metadata", out PdfObject? metadataObject)) {
            ReplaceXmpMetadata(objects, catalog, metadataObject, existingXmp, updated);
        } else if (createXmpMetadata) {
            int metadataObjectNumber = NextObjectNumber(objects);
            catalog.Items["Metadata"] = new PdfReference(metadataObjectNumber, 0);
            var dictionary = new PdfDictionary();
            dictionary.Items["Type"] = new PdfName("Metadata");
            dictionary.Items["Subtype"] = new PdfName("XML");
            objects[metadataObjectNumber] = new PdfIndirectObject(
                metadataObjectNumber,
                0,
                new PdfStream(dictionary, PdfXmpMetadataBuilder.Build(updated.Title, updated.Author, updated.Subject, updated.Keywords)));
        }

        return infoObjectNumber;
    }

    private static int ReplaceInfoDictionary(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDocumentSecurityInfo security,
        PdfMetadata updated) {
        int objectNumber = security.InfoObjectNumber.HasValue && objects.ContainsKey(security.InfoObjectNumber.Value)
            ? security.InfoObjectNumber.Value
            : NextObjectNumber(objects);
        int generation = objects.TryGetValue(objectNumber, out PdfIndirectObject? existingInfo)
            ? existingInfo.Generation
            : 0;
        objects[objectNumber] = new PdfIndirectObject(objectNumber, generation, PdfInfoDictionaryBuilder.BuildDictionary(updated));
        return objectNumber;
    }

    private static PdfDictionary RequireCatalog(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDocumentSecurityInfo security) {
        if (!security.RootObjectNumber.HasValue ||
            !objects.TryGetValue(security.RootObjectNumber.Value, out PdfIndirectObject? catalogObject) ||
            catalogObject.Value is not PdfDictionary catalog) {
            throw new InvalidOperationException("The active PDF trailer does not reference a readable catalog object.");
        }

        return catalog;
    }

    private static void ReplaceXmpMetadata(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary catalog,
        PdfObject metadataObject,
        PdfXmpMetadataInfo? existingXmp,
        PdfMetadata updated) {
        if (existingXmp is null || existingXmp.RawXml is null || !existingXmp.IsWellFormedXml || existingXmp.HasUnsupportedFilters) {
            throw new InvalidOperationException("The existing XMP metadata stream cannot be decoded and preserved safely.");
        }

        byte[] xml = PdfXmpMetadataSynchronizer.Synchronize(existingXmp.RawXml, updated);
        if (metadataObject is PdfReference reference) {
            if (!PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect) || indirect.Value is not PdfStream stream) {
                throw new InvalidOperationException("The catalog XMP metadata reference does not resolve to a readable stream.");
            }

            objects[indirect.ObjectNumber] = new PdfIndirectObject(
                indirect.ObjectNumber,
                indirect.Generation,
                new PdfStream(PdfXmpMetadataSynchronizer.CloneUnfilteredMetadataDictionary(stream.Dictionary), xml));
            return;
        }

        if (metadataObject is PdfStream directStream) {
            catalog.Items["Metadata"] = new PdfStream(
                PdfXmpMetadataSynchronizer.CloneUnfilteredMetadataDictionary(directStream.Dictionary),
                xml);
            return;
        }

        throw new InvalidOperationException("The catalog XMP metadata entry is not a readable stream.");
    }

    private static int NextObjectNumber(Dictionary<int, PdfIndirectObject> objects) {
        return objects.Count == 0 ? 1 : checked(objects.Keys.Max() + 1);
    }
}
