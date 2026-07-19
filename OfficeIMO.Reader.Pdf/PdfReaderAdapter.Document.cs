using OfficeIMO.Pdf;

namespace OfficeIMO.Reader.Pdf;

internal static partial class PdfReaderAdapter {
    /// <summary>
    /// Reads a PDF file and returns the shared OfficeIMO read result JSON envelope.
    /// </summary>
    public static string ReadDocumentJson(string pdfPath, ReaderOptions? readerOptions = null, ReaderPdfOptions? pdfOptions = null, bool indented = false, CancellationToken cancellationToken = default) {
        OfficeDocumentReadResult result = ReadDocument(pdfPath, readerOptions, pdfOptions, cancellationToken);
        return OfficeDocumentReadResultJson.Serialize(result, indented);
    }

    /// <summary>
    /// Reads a PDF stream and returns the shared OfficeIMO read result JSON envelope.
    /// </summary>
    public static string ReadDocumentJson(Stream pdfStream, string? sourceName = null, ReaderOptions? readerOptions = null, ReaderPdfOptions? pdfOptions = null, bool indented = false, CancellationToken cancellationToken = default) {
        OfficeDocumentReadResult result = ReadDocument(pdfStream, sourceName, readerOptions, pdfOptions, cancellationToken);
        return OfficeDocumentReadResultJson.Serialize(result, indented);
    }

    /// <summary>
    /// Reads PDF bytes and returns the shared OfficeIMO read result JSON envelope.
    /// </summary>
    public static string ReadDocumentJson(byte[] pdfBytes, string? sourceName = null, ReaderOptions? readerOptions = null, ReaderPdfOptions? pdfOptions = null, bool indented = false, CancellationToken cancellationToken = default) {
        OfficeDocumentReadResult result = ReadDocument(pdfBytes, sourceName, readerOptions, pdfOptions, cancellationToken);
        return OfficeDocumentReadResultJson.Serialize(result, indented);
    }

    /// <summary>
    /// Converts an already loaded logical PDF model into the shared OfficeIMO read result JSON envelope.
    /// </summary>
    public static string ReadDocumentJson(PdfLogicalDocument document, string sourceName = "document.pdf", ReaderOptions? readerOptions = null, ReaderPdfOptions? pdfOptions = null, bool indented = false, CancellationToken cancellationToken = default) {
        OfficeDocumentReadResult result = ReadDocument(document, sourceName, readerOptions, pdfOptions, cancellationToken);
        return OfficeDocumentReadResultJson.Serialize(result, indented);
    }

    /// <summary>
    /// Reads a PDF file and returns logical tables in source order.
    /// </summary>
    public static IReadOnlyList<ReaderTable> ReadTables(string pdfPath, ReaderOptions? readerOptions = null, ReaderPdfOptions? pdfOptions = null, CancellationToken cancellationToken = default) {
        return DocumentReaderEngine.ExtractTables(Read(pdfPath, readerOptions, pdfOptions, cancellationToken), cancellationToken);
    }

    /// <summary>
    /// Reads a PDF stream and returns logical tables in source order.
    /// </summary>
    public static IReadOnlyList<ReaderTable> ReadTables(Stream pdfStream, string? sourceName = null, ReaderOptions? readerOptions = null, ReaderPdfOptions? pdfOptions = null, CancellationToken cancellationToken = default) {
        return DocumentReaderEngine.ExtractTables(Read(pdfStream, sourceName, readerOptions, pdfOptions, cancellationToken), cancellationToken);
    }

    /// <summary>
    /// Reads PDF bytes and returns logical tables in source order.
    /// </summary>
    public static IReadOnlyList<ReaderTable> ReadTables(byte[] pdfBytes, string? sourceName = null, ReaderOptions? readerOptions = null, ReaderPdfOptions? pdfOptions = null, CancellationToken cancellationToken = default) {
        return DocumentReaderEngine.ExtractTables(Read(pdfBytes, sourceName, readerOptions, pdfOptions, cancellationToken), cancellationToken);
    }

    /// <summary>
    /// Converts an already loaded logical PDF model into logical tables in source order.
    /// </summary>
    public static IReadOnlyList<ReaderTable> ReadTables(PdfLogicalDocument document, string sourceName = "document.pdf", ReaderOptions? readerOptions = null, ReaderPdfOptions? pdfOptions = null, CancellationToken cancellationToken = default) {
        return DocumentReaderEngine.ExtractTables(Read(document, sourceName, readerOptions, pdfOptions, cancellationToken), cancellationToken);
    }

    /// <summary>
    /// Reads a PDF file and returns logical table export payloads in source order.
    /// </summary>
    public static IReadOnlyList<ReaderTableExportBundle> ReadTableExports(string pdfPath, ReaderOptions? readerOptions = null, ReaderPdfOptions? pdfOptions = null, bool indentedJson = false, CancellationToken cancellationToken = default) {
        return DocumentReaderEngine.ExportTables(ReadTables(pdfPath, readerOptions, pdfOptions, cancellationToken), indentedJson, cancellationToken);
    }

    /// <summary>
    /// Reads a PDF stream and returns logical table export payloads in source order.
    /// </summary>
    public static IReadOnlyList<ReaderTableExportBundle> ReadTableExports(Stream pdfStream, string? sourceName = null, ReaderOptions? readerOptions = null, ReaderPdfOptions? pdfOptions = null, bool indentedJson = false, CancellationToken cancellationToken = default) {
        return DocumentReaderEngine.ExportTables(ReadTables(pdfStream, sourceName, readerOptions, pdfOptions, cancellationToken), indentedJson, cancellationToken);
    }

    /// <summary>
    /// Reads PDF bytes and returns logical table export payloads in source order.
    /// </summary>
    public static IReadOnlyList<ReaderTableExportBundle> ReadTableExports(byte[] pdfBytes, string? sourceName = null, ReaderOptions? readerOptions = null, ReaderPdfOptions? pdfOptions = null, bool indentedJson = false, CancellationToken cancellationToken = default) {
        return DocumentReaderEngine.ExportTables(ReadTables(pdfBytes, sourceName, readerOptions, pdfOptions, cancellationToken), indentedJson, cancellationToken);
    }

    /// <summary>
    /// Converts an already loaded logical PDF model into logical table export payloads in source order.
    /// </summary>
    public static IReadOnlyList<ReaderTableExportBundle> ReadTableExports(PdfLogicalDocument document, string sourceName = "document.pdf", ReaderOptions? readerOptions = null, ReaderPdfOptions? pdfOptions = null, bool indentedJson = false, CancellationToken cancellationToken = default) {
        return DocumentReaderEngine.ExportTables(ReadTables(document, sourceName, readerOptions, pdfOptions, cancellationToken), indentedJson, cancellationToken);
    }

    /// <summary>
    /// Reads a PDF file and returns the shared OfficeIMO read result envelope.
    /// </summary>
    public static OfficeDocumentReadResult ReadDocument(string pdfPath, ReaderOptions? readerOptions = null, ReaderPdfOptions? pdfOptions = null, CancellationToken cancellationToken = default) {
        if (pdfPath == null) throw new ArgumentNullException(nameof(pdfPath));
        if (pdfPath.Length == 0) throw new ArgumentException("PDF path cannot be empty.", nameof(pdfPath));
        if (!File.Exists(pdfPath)) throw new FileNotFoundException($"PDF file '{pdfPath}' doesn't exist.", pdfPath);

        var effectiveReaderOptions = readerOptions ?? new ReaderOptions();
        var effectivePdfOptions = ReaderPdfOptionsCloner.CloneOrDefault(pdfOptions);
        ReaderInputLimits.EnforceFileSize(pdfPath, effectiveReaderOptions.MaxInputBytes);
        var source = BuildSourceMetadataFromPath(pdfPath, effectiveReaderOptions.ComputeHashes);
        PdfDocument pdf = PdfDocument.Open(pdfPath, CreatePdfReadOptions(effectiveReaderOptions));
        PdfDocumentPreflight preflight = pdf.Preflight();
        PdfLogicalDocument document = LoadDocument(pdf, effectivePdfOptions);
        return BuildDocumentResult(document, source, effectiveReaderOptions, effectivePdfOptions, preflight, applyPageRanges: false, cancellationToken);
    }

    /// <summary>
    /// Reads a PDF stream and returns the shared OfficeIMO read result envelope.
    /// </summary>
    public static OfficeDocumentReadResult ReadDocument(Stream pdfStream, string? sourceName = null, ReaderOptions? readerOptions = null, ReaderPdfOptions? pdfOptions = null, CancellationToken cancellationToken = default) {
        if (pdfStream == null) throw new ArgumentNullException(nameof(pdfStream));
        if (!pdfStream.CanRead) throw new ArgumentException("PDF stream must be readable.", nameof(pdfStream));

        var effectiveReaderOptions = readerOptions ?? new ReaderOptions();
        var effectivePdfOptions = ReaderPdfOptionsCloner.CloneOrDefault(pdfOptions);
        var logicalSourceName = NormalizeLogicalSourceName(sourceName, "document.pdf");
        var source = new SourceMetadata {
            Path = logicalSourceName,
            SourceId = BuildSourceId(logicalSourceName)
        };

        cancellationToken.ThrowIfCancellationRequested();
        PdfDocument pdf = OpenReaderPdf(pdfStream, effectiveReaderOptions);
        UpdateSourceMetadataFromPdfDocument(source, pdf, effectiveReaderOptions.ComputeHashes);
        PdfDocumentPreflight preflight = pdf.Preflight();
        PdfLogicalDocument document = LoadDocument(pdf, effectivePdfOptions);
        return BuildDocumentResult(document, source, effectiveReaderOptions, effectivePdfOptions, preflight, applyPageRanges: false, cancellationToken);
    }

    /// <summary>
    /// Reads PDF bytes and returns the shared OfficeIMO read result envelope.
    /// </summary>
    public static OfficeDocumentReadResult ReadDocument(byte[] pdfBytes, string? sourceName = null, ReaderOptions? readerOptions = null, ReaderPdfOptions? pdfOptions = null, CancellationToken cancellationToken = default) {
        if (pdfBytes == null) throw new ArgumentNullException(nameof(pdfBytes));

        using var stream = new MemoryStream(pdfBytes, writable: false);
        return ReadDocument(stream, sourceName, readerOptions, pdfOptions, cancellationToken);
    }

    /// <summary>
    /// Converts an already loaded logical PDF model into the shared OfficeIMO read result envelope.
    /// </summary>
    public static OfficeDocumentReadResult ReadDocument(PdfLogicalDocument document, string sourceName = "document.pdf", ReaderOptions? readerOptions = null, ReaderPdfOptions? pdfOptions = null, CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (sourceName == null) throw new ArgumentNullException(nameof(sourceName));

        var effectiveReaderOptions = readerOptions ?? new ReaderOptions();
        var effectivePdfOptions = ReaderPdfOptionsCloner.CloneOrDefault(pdfOptions);
        var logicalSourceName = NormalizeLogicalSourceName(sourceName, "document.pdf");
        var source = new SourceMetadata {
            Path = logicalSourceName,
            SourceId = BuildSourceId(logicalSourceName)
        };

        return BuildDocumentResult(document, source, effectiveReaderOptions, effectivePdfOptions, preflight: null, applyPageRanges: true, cancellationToken);
    }

    private static OfficeDocumentReadResult BuildDocumentResult(PdfLogicalDocument document, SourceMetadata source, ReaderOptions readerOptions, ReaderPdfOptions pdfOptions, PdfDocumentPreflight? preflight, bool applyPageRanges, CancellationToken cancellationToken) {
        var markdownOptions = ReaderPdfOptions.CloneMarkdownOptions(pdfOptions.MarkdownOptions) ?? ReaderPdfOptions.CreateOfficeIMOProfile().MarkdownOptions!;
        markdownOptions.IncludePageSeparators = false;
        IReadOnlyList<PdfLogicalPage> pages = applyPageRanges ? GetReaderPages(document, pdfOptions) : document.Pages;
        string markdown = BuildMarkdown(pages, markdownOptions);
        ReaderChunk[] chunks = Read(document, source, readerOptions, pdfOptions, applyPageRanges, cancellationToken).ToArray();
        ReaderTable[] tables = DocumentReaderEngine.ExtractTables(chunks, cancellationToken).ToArray();
        OfficeDocumentBlock[] blocks = BuildDocumentBlocks(pages, source).ToArray();
        OfficeDocumentAsset[] assets = BuildDocumentAssets(pages, source).ToArray();
        OfficeDocumentLink[] links = BuildDocumentLinks(pages, source).ToArray();
        OfficeDocumentFormField[] forms = BuildDocumentForms(pages, source).ToArray();
        OfficeDocumentOcrCandidate[] ocrCandidates = BuildDocumentOcrCandidates(pages, source, assets).ToArray();

        return new OfficeDocumentReadResult {
            Kind = ReaderInputKind.Pdf,
            Source = new OfficeDocumentSource {
                Path = source.Path,
                SourceId = source.SourceId,
                SourceHash = source.SourceHash,
                LastWriteUtc = source.LastWriteUtc,
                LengthBytes = source.LengthBytes,
                Title = document.Metadata.Title,
                Author = document.Metadata.Author,
                Subject = document.Metadata.Subject,
                Keywords = document.Metadata.Keywords
            },
            CapabilitiesUsed = new[] {
                "officeimo.reader.pdf",
                "officeimo.pdf.logical-document",
                "officeimo.pdf.logical-markdown",
                "officeimo.reader.pdf.pages.native"
            },
            Markdown = markdown,
            Chunks = chunks,
            Metadata = BuildDocumentMetadata(document, source, preflight, applyPageRanges ? pages : null, tables, ocrCandidates),
            Pages = BuildDocumentPages(pages, source, blocks, tables, assets, links, forms, ocrCandidates),
            Blocks = blocks,
            Tables = tables,
            Assets = assets,
            Links = links,
            Forms = forms,
            OcrCandidates = ocrCandidates,
            Visuals = chunks
                .Where(chunk => chunk.Visuals != null)
                .SelectMany(chunk => chunk.Visuals!)
                .ToArray(),
            Diagnostics = BuildDocumentDiagnostics(chunks, ocrCandidates, preflight)
        };
    }

    private static IReadOnlyList<OfficeDocumentPage> BuildDocumentPages(
        IReadOnlyList<PdfLogicalPage> pages,
        SourceMetadata source,
        IReadOnlyList<OfficeDocumentBlock> blocks,
        IReadOnlyList<ReaderTable> tables,
        IReadOnlyList<OfficeDocumentAsset> assets,
        IReadOnlyList<OfficeDocumentLink> links,
        IReadOnlyList<OfficeDocumentFormField> forms,
        IReadOnlyList<OfficeDocumentOcrCandidate> ocrCandidates) {
        if (pages.Count == 0) {
            return Array.Empty<OfficeDocumentPage>();
        }

        var result = new List<OfficeDocumentPage>(pages.Count);
        for (int i = 0; i < pages.Count; i++) {
            PdfLogicalPage page = pages[i];
            result.Add(new OfficeDocumentPage {
                Number = page.PageNumber,
                Width = page.Width,
                Height = page.Height,
                RotationDegrees = page.RotationDegrees,
                Location = BuildLocation(source, page.PageNumber, i, "page", "page-" + page.PageNumber.ToString(CultureInfo.InvariantCulture) + "-selection-" + i.ToString("D4", CultureInfo.InvariantCulture)),
                Blocks = blocks.Where(block => IsPageSelection(block.Location, page.PageNumber, i)).ToArray(),
                Tables = tables.Where(table => table.Location != null && IsPageSelection(table.Location, page.PageNumber, i)).ToArray(),
                Assets = assets.Where(asset => IsPageSelection(asset.Location, page.PageNumber, i)).ToArray(),
                Links = links.Where(link => IsPageSelection(link.Location, page.PageNumber, i)).ToArray(),
                Forms = forms.Where(form => IsPageSelection(form.Location, page.PageNumber, i)).ToArray(),
                OcrCandidates = ocrCandidates.Where(candidate => IsPageSelection(candidate.Location, page.PageNumber, i)).ToArray()
            });
        }

        return result.AsReadOnly();
    }

    private static IEnumerable<OfficeDocumentBlock> BuildDocumentBlocks(IReadOnlyList<PdfLogicalPage> pages, SourceMetadata source) {
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++) {
            PdfLogicalPage page = pages[pageIndex];
            int tableIndex = 0;
            for (int elementIndex = 0; elementIndex < page.Elements.Count; elementIndex++) {
                IPdfLogicalElement element = page.Elements[elementIndex];
                if (element is PdfLogicalTextBlock textBlock) {
                    PdfLogicalHeading? heading = FindHeading(page, textBlock);
                    PdfLogicalListItem? listItem = FindListItem(page, textBlock);
                    string kind = heading != null
                        ? "heading"
                        : listItem != null
                            ? "list-item"
                            : ToDocumentBlockKind(textBlock.Kind);
                    yield return new OfficeDocumentBlock {
                        Id = "pdf-page-" + page.PageNumber.ToString("D4", CultureInfo.InvariantCulture) + "-selection-" + pageIndex.ToString("D4", CultureInfo.InvariantCulture) + "-block-" + elementIndex.ToString("D4", CultureInfo.InvariantCulture),
                        Kind = kind,
                        Text = heading?.Text ?? listItem?.Text ?? textBlock.Text,
                        Level = heading?.Level ?? listItem?.Level,
                        Marker = listItem?.Marker,
                        Location = BuildLocation(source, page.PageNumber, pageIndex, kind, "page-" + page.PageNumber.ToString(CultureInfo.InvariantCulture) + "-selection-" + pageIndex.ToString("D4", CultureInfo.InvariantCulture) + "-block-" + elementIndex.ToString("D4", CultureInfo.InvariantCulture)),
                        Region = new OfficeDocumentRegion {
                            X = textBlock.XStart,
                            Y = textBlock.BaselineY,
                            Width = Math.Max(0D, textBlock.XEnd - textBlock.XStart),
                            Height = 0D
                        }
                    };
                } else if (element is PdfLogicalLeaderRow leaderRow) {
                    yield return new OfficeDocumentBlock {
                        Id = "pdf-page-" + page.PageNumber.ToString("D4", CultureInfo.InvariantCulture) + "-selection-" + pageIndex.ToString("D4", CultureInfo.InvariantCulture) + "-block-" + elementIndex.ToString("D4", CultureInfo.InvariantCulture),
                        Kind = "leader-row",
                        Text = leaderRow.Label + "\t" + leaderRow.Value,
                        Marker = leaderRow.Label,
                        Location = BuildLocation(source, page.PageNumber, pageIndex, "leader-row", "page-" + page.PageNumber.ToString(CultureInfo.InvariantCulture) + "-selection-" + pageIndex.ToString("D4", CultureInfo.InvariantCulture) + "-block-" + elementIndex.ToString("D4", CultureInfo.InvariantCulture))
                    };
                } else if (element is PdfLogicalTable table) {
                    yield return new OfficeDocumentBlock {
                        Id = "pdf-page-" + page.PageNumber.ToString("D4", CultureInfo.InvariantCulture) + "-selection-" + pageIndex.ToString("D4", CultureInfo.InvariantCulture) + "-table-" + tableIndex.ToString("D4", CultureInfo.InvariantCulture),
                        Kind = "table",
                        Text = "Detected PDF table with " + table.Rows.Count.ToString(CultureInfo.InvariantCulture) + " row(s).",
                        Location = BuildLocation(source, page.PageNumber, pageIndex, "table", "page-" + page.PageNumber.ToString(CultureInfo.InvariantCulture) + "-selection-" + pageIndex.ToString("D4", CultureInfo.InvariantCulture) + "-table-" + tableIndex.ToString(CultureInfo.InvariantCulture)),
                        Region = new OfficeDocumentRegion {
                            X = table.Columns.Count > 0 ? table.Columns[0].From : 0D,
                            Y = table.YBottom,
                            Width = table.Columns.Count > 0 ? Math.Max(0D, table.Columns[table.Columns.Count - 1].To - table.Columns[0].From) : 0D,
                            Height = Math.Max(0D, table.YTop - table.YBottom)
                        }
                    };
                    tableIndex++;
                }
            }
        }
    }

    private static IEnumerable<OfficeDocumentAsset> BuildDocumentAssets(IReadOnlyList<PdfLogicalPage> pages, SourceMetadata source) {
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++) {
            PdfLogicalPage page = pages[pageIndex];
            for (int imageIndex = 0; imageIndex < page.Images.Count; imageIndex++) {
                PdfLogicalImage image = page.Images[imageIndex];
                string imageObjectId = image.SourceImage.ObjectNumber > 0
                    ? image.SourceImage.ObjectNumber.ToString(CultureInfo.InvariantCulture)
                    : image.ResourceName;
                string assetId = "pdf-page-" + page.PageNumber.ToString("D4", CultureInfo.InvariantCulture) + "-selection-" + pageIndex.ToString("D4", CultureInfo.InvariantCulture) + "-image-" + imageIndex.ToString("D4", CultureInfo.InvariantCulture);
                yield return new OfficeDocumentAsset {
                    Id = assetId,
                    Kind = "image",
                    MediaType = image.MimeType,
                    Extension = image.SourceImage.FileExtension,
                    FileName = OfficeDocumentAssetNaming.BuildFileName(assetId, image.SourceImage.FileExtension),
                    Width = image.Width,
                    Height = image.Height,
                    LengthBytes = image.SourceImage.Bytes.LongLength,
                    PayloadHash = ComputeSha256Hex(image.SourceImage.Bytes),
                    PayloadBytes = image.SourceImage.Bytes,
                    SourceObjectId = imageObjectId,
                    Region = BuildImageAssetRegion(image),
                    Location = BuildLocation(source, page.PageNumber, pageIndex, "image", "page-" + page.PageNumber.ToString(CultureInfo.InvariantCulture) + "-selection-" + pageIndex.ToString("D4", CultureInfo.InvariantCulture) + "-image-" + imageIndex.ToString(CultureInfo.InvariantCulture))
                };
            }
        }
    }

    private static OfficeDocumentRegion? BuildImageAssetRegion(PdfLogicalImage image) {
        PdfImagePlacement? placement = image.PrimaryPlacement;
        if (placement == null) {
            return null;
        }

        return new OfficeDocumentRegion {
            X = placement.X,
            Y = placement.Y,
            Width = placement.Width,
            Height = placement.Height
        };
    }

    private static IEnumerable<OfficeDocumentOcrCandidate> BuildDocumentOcrCandidates(IReadOnlyList<PdfLogicalPage> pages, SourceMetadata source, IReadOnlyList<OfficeDocumentAsset> assets) {
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++) {
            PdfLogicalPage page = pages[pageIndex];
            if (page.Images.Count == 0 || HasMeaningfulNativeText(page)) {
                continue;
            }

            int pageNumber = page.PageNumber;
            OfficeDocumentAsset? asset = assets.FirstOrDefault(asset => IsPageSelection(asset.Location, pageNumber, pageIndex));
            PdfLogicalImage? image = page.Images.Count == 1 ? page.Images[0] : null;

            yield return new OfficeDocumentOcrCandidate {
                Id = "pdf-page-" + pageNumber.ToString("D4", CultureInfo.InvariantCulture) + "-selection-" + pageIndex.ToString("D4", CultureInfo.InvariantCulture) + "-ocr-0000",
                Kind = image == null ? "page" : "image",
                Reason = "PDF page contains image content but no meaningful native text blocks.",
                Confidence = 0.85D,
                AssetId = asset?.Id,
                ImageCount = page.Images.Count,
                TextBlockCount = page.TextBlocks.Count,
                Location = BuildLocation(source, pageNumber, pageIndex, "ocr-candidate", "page-" + pageNumber.ToString(CultureInfo.InvariantCulture) + "-selection-" + pageIndex.ToString("D4", CultureInfo.InvariantCulture) + "-ocr-0000"),
                Region = BuildOcrCandidateRegion(page, image)
            };
        }
    }

    private static bool HasMeaningfulNativeText(PdfLogicalPage page) {
        for (int i = 0; i < page.TextBlocks.Count; i++) {
            string text = page.TextBlocks[i].Text;
            if (!string.IsNullOrWhiteSpace(text)) {
                return true;
            }
        }

        return false;
    }

    private static OfficeDocumentRegion BuildOcrCandidateRegion(PdfLogicalPage page, PdfLogicalImage? image) {
        PdfImagePlacement? placement = image?.Placements.Count > 0 ? image.Placements[0] : null;
        if (placement != null) {
            return new OfficeDocumentRegion {
                X = placement.X,
                Y = placement.Y,
                Width = placement.Width,
                Height = placement.Height
            };
        }

        return new OfficeDocumentRegion {
            X = 0D,
            Y = 0D,
            Width = page.Width,
            Height = page.Height
        };
    }

    private static IEnumerable<OfficeDocumentLink> BuildDocumentLinks(IReadOnlyList<PdfLogicalPage> pages, SourceMetadata source) {
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++) {
            PdfLogicalPage page = pages[pageIndex];
            for (int linkIndex = 0; linkIndex < page.Links.Count; linkIndex++) {
                PdfLogicalLinkAnnotation link = page.Links[linkIndex];
                yield return new OfficeDocumentLink {
                    Id = "pdf-page-" + page.PageNumber.ToString("D4", CultureInfo.InvariantCulture) + "-selection-" + pageIndex.ToString("D4", CultureInfo.InvariantCulture) + "-link-" + linkIndex.ToString("D4", CultureInfo.InvariantCulture),
                    Kind = GetLinkKind(link),
                    Uri = link.Uri,
                    DestinationName = link.DestinationName,
                    DestinationPageNumber = link.DestinationPageNumber,
                    DestinationMode = link.DestinationMode?.ToString(),
                    DestinationTop = link.DestinationTop,
                    DestinationLeft = link.DestinationLeft,
                    DestinationBottom = link.DestinationBottom,
                    DestinationRight = link.DestinationRight,
                    NamedAction = link.NamedAction,
                    RemoteFile = link.RemoteFile,
                    RemoteDestinationName = link.RemoteDestinationName,
                    RemoteDestinationPageNumber = link.RemoteDestinationPageNumber,
                    RemoteDestinationMode = link.RemoteDestinationMode?.ToString(),
                    RemoteDestinationTop = link.RemoteDestinationTop,
                    RemoteDestinationLeft = link.RemoteDestinationLeft,
                    RemoteDestinationBottom = link.RemoteDestinationBottom,
                    RemoteDestinationRight = link.RemoteDestinationRight,
                    Text = link.Contents,
                    Location = BuildLocation(source, page.PageNumber, pageIndex, "link", "page-" + page.PageNumber.ToString(CultureInfo.InvariantCulture) + "-selection-" + pageIndex.ToString("D4", CultureInfo.InvariantCulture) + "-link-" + linkIndex.ToString(CultureInfo.InvariantCulture)),
                    Region = new OfficeDocumentRegion {
                        X = link.X1,
                        Y = link.Y1,
                        Width = link.Width,
                        Height = link.Height
                    }
                };
            }
        }
    }

    private static IEnumerable<OfficeDocumentFormField> BuildDocumentForms(IReadOnlyList<PdfLogicalPage> pages, SourceMetadata source) {
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++) {
            PdfLogicalPage page = pages[pageIndex];
            for (int formIndex = 0; formIndex < page.FormWidgets.Count; formIndex++) {
                PdfLogicalFormWidget widget = page.FormWidgets[formIndex];
                yield return new OfficeDocumentFormField {
                    Id = "pdf-page-" + page.PageNumber.ToString("D4", CultureInfo.InvariantCulture) + "-selection-" + pageIndex.ToString("D4", CultureInfo.InvariantCulture) + "-form-" + formIndex.ToString("D4", CultureInfo.InvariantCulture),
                    Name = widget.FieldName,
                    Kind = widget.Field.Kind.ToString(),
                    Value = widget.Value,
                    IsReadOnly = widget.Field.IsReadOnly || widget.IsReadOnly,
                    IsRequired = widget.Field.IsRequired,
                    Location = BuildLocation(source, page.PageNumber, pageIndex, "form-widget", "page-" + page.PageNumber.ToString(CultureInfo.InvariantCulture) + "-selection-" + pageIndex.ToString("D4", CultureInfo.InvariantCulture) + "-form-" + formIndex.ToString(CultureInfo.InvariantCulture)),
                    Region = new OfficeDocumentRegion {
                        X = widget.X1,
                        Y = widget.Y1,
                        Width = widget.Width,
                        Height = widget.Height
                    }
                };
            }
        }
    }

    private static IReadOnlyList<OfficeDocumentMetadataEntry> BuildDocumentMetadata(PdfLogicalDocument document, SourceMetadata source, PdfDocumentPreflight? preflight, IReadOnlyList<PdfLogicalPage>? selectedPages, IReadOnlyList<ReaderTable> tables, IReadOnlyList<OfficeDocumentOcrCandidate> ocrCandidates) {
        var entries = new List<OfficeDocumentMetadataEntry>();
        HashSet<int>? selectedPageNumbers = BuildSelectedPageNumberSet(selectedPages);
        AddMetadata(entries, "pdf-catalog-page-mode", "pdf.catalog", "PageMode", document.CatalogPageMode);
        AddMetadata(entries, "pdf-catalog-page-layout", "pdf.catalog", "PageLayout", document.CatalogPageLayout);
        AddMetadata(entries, "pdf-catalog-version", "pdf.catalog", "Version", document.CatalogVersion);
        AddMetadata(entries, "pdf-catalog-language", "pdf.catalog", "Language", document.CatalogLanguage);
        AddCountMetadata(entries, "pdf-outline-count", "pdf.outline", "Count", CountOutlines(document.Outlines, selectedPageNumbers));
        AddCountMetadata(entries, "pdf-named-destination-count", "pdf.destination", "Count", CountNamedDestinations(document.NamedDestinations, selectedPageNumbers));
        AddAttachmentMetadata(entries, document.Attachments);
        AddOutputIntentMetadata(entries, document.OutputIntents);
        AddOptionalContentMetadata(entries, document.OptionalContent);
        AddTaggedContentMetadata(entries, document.TaggedContent);
        AddXmpMetadata(entries, document.XmpMetadata);
        AddSecurityMetadata(entries, document.Security);
        AddCountMetadata(entries, "pdf-catalog-action-count", "pdf.catalog.action", "Count", document.CatalogActions.Count);
        AddCountMetadata(entries, "pdf-form-field-count", "pdf.form", "Count", CountFormFields(document.FormFields, selectedPageNumbers));
        AddImageGeometryMetadata(entries, selectedPages ?? document.Pages);
        AddTableQualityMetadata(entries, tables);
        AddOcrCandidateMetadata(entries, ocrCandidates);
        AddLinkMetadata(entries, selectedPages ?? document.Pages);
        AddAnnotationMetadata(entries, source, selectedPages ?? document.Pages);
        AddFormWidgetMetadata(entries, selectedPages ?? document.Pages);
        AddActionMetadata(entries, document, selectedPages ?? document.Pages);
        AddMetadata(entries, "pdf-acroform-need-appearances", "pdf.form", "NeedAppearances", ToMetadataText(document.AcroFormNeedAppearances), "boolean");
        AddMetadata(entries, "pdf-acroform-signature-flags", "pdf.form", "SignatureFlags", ToMetadataText(document.AcroFormSignatureFlags), "number");
        AddAcroFormXfaMetadata(entries, document.AcroFormXfa);
        entries.AddRange(BuildPdfPreflightMetadata(preflight));

        if (document.OpenAction != null && IsSelectedPage(document.OpenAction.PageNumber, selectedPageNumbers)) {
            entries.Add(new OfficeDocumentMetadataEntry {
                Id = "pdf-open-action",
                Category = "pdf.catalog.openAction",
                Name = document.OpenAction.ActionType,
                Value = document.OpenAction.PageNumber?.ToString(CultureInfo.InvariantCulture),
                ValueType = "object",
                Location = BuildMetadataLocation(source, document.OpenAction.PageNumber, "open-action", "pdf-open-action"),
                Attributes = BuildDestinationAttributes(
                    document.OpenAction.PageNumber,
                    document.OpenAction.DestinationTop,
                    document.OpenAction.DestinationMode,
                    document.OpenAction.DestinationLeft,
                    document.OpenAction.DestinationBottom,
                    document.OpenAction.DestinationRight)
            });
        }

        if (document.ViewerPreferences != null) {
            foreach (KeyValuePair<string, string> preference in document.ViewerPreferences.Values.OrderBy(item => item.Key, StringComparer.Ordinal)) {
                entries.Add(new OfficeDocumentMetadataEntry {
                    Id = "pdf-viewer-preference-" + entries.Count.ToString("D4", CultureInfo.InvariantCulture),
                    Category = "pdf.viewerPreference",
                    Name = preference.Key,
                    Value = preference.Value,
                    ValueType = "string"
                });
            }
        }

        AddOutlineMetadata(entries, source, document.Outlines, selectedPageNumbers);

        for (int i = 0; i < document.NamedDestinations.Count; i++) {
            PdfNamedDestination destination = document.NamedDestinations[i];
            if (!IsSelectedPage(destination.PageNumber, selectedPageNumbers)) {
                continue;
            }

            entries.Add(new OfficeDocumentMetadataEntry {
                Id = "pdf-named-destination-" + i.ToString("D4", CultureInfo.InvariantCulture),
                Category = "pdf.destination",
                Name = destination.Name,
                Value = destination.PageNumber?.ToString(CultureInfo.InvariantCulture),
                ValueType = "object",
                Location = BuildMetadataLocation(source, destination.PageNumber, "named-destination", "pdf-named-destination-" + i.ToString("D4", CultureInfo.InvariantCulture)),
                Attributes = BuildDestinationAttributes(
                    destination.PageNumber,
                    destination.DestinationTop,
                    destination.DestinationMode,
                    destination.DestinationLeft,
                    destination.DestinationBottom,
                    destination.DestinationRight)
            });
        }

        for (int i = 0; i < document.CatalogActions.Count; i++) {
            PdfCatalogAction action = document.CatalogActions[i];
            var attributes = new Dictionary<string, string>(StringComparer.Ordinal) {
                ["actionType"] = action.ActionType,
                ["source"] = action.Source
            };
            if (!string.IsNullOrWhiteSpace(action.TriggerName)) {
                attributes["triggerName"] = action.TriggerName!;
            }

            entries.Add(new OfficeDocumentMetadataEntry {
                Id = "pdf-catalog-action-" + i.ToString("D4", CultureInfo.InvariantCulture),
                Category = "pdf.catalog.action",
                Name = action.Name,
                Value = action.ActionType,
                ValueType = "object",
                Attributes = attributes
            });
        }

        return entries.Count == 0 ? Array.Empty<OfficeDocumentMetadataEntry>() : entries.AsReadOnly();
    }

    private static void AddOutlineMetadata(List<OfficeDocumentMetadataEntry> entries, SourceMetadata source, IReadOnlyList<PdfOutlineItem> outlines, HashSet<int>? selectedPageNumbers) {
        for (int i = 0; i < outlines.Count; i++) {
            PdfOutlineItem outline = outlines[i];
            if (IsSelectedPage(outline.PageNumber, selectedPageNumbers)) {
                string id = "pdf-outline-" + entries.Count.ToString("D4", CultureInfo.InvariantCulture);
                var attributes = BuildDestinationAttributes(
                    outline.PageNumber,
                    outline.DestinationTop,
                    outline.DestinationMode,
                    outline.DestinationLeft,
                    outline.DestinationBottom,
                    outline.DestinationRight);
                attributes["level"] = outline.Level.ToString(CultureInfo.InvariantCulture);
                attributes["isExpanded"] = ToMetadataText(outline.IsExpanded)!;
                attributes["childCount"] = CountOutlines(outline.Children, selectedPageNumbers).ToString(CultureInfo.InvariantCulture);

                entries.Add(new OfficeDocumentMetadataEntry {
                    Id = id,
                    Category = "pdf.outline",
                    Name = outline.Title,
                    Value = outline.Title,
                    ValueType = "object",
                    Location = BuildMetadataLocation(source, outline.PageNumber, "outline", id),
                    Attributes = attributes
                });
            }

            AddOutlineMetadata(entries, source, outline.Children, selectedPageNumbers);
        }
    }

    private static Dictionary<string, string> BuildDestinationAttributes(
        int? pageNumber,
        double? destinationTop,
        PdfOpenActionDestinationMode? destinationMode,
        double? destinationLeft,
        double? destinationBottom,
        double? destinationRight) {
        var attributes = new Dictionary<string, string>(StringComparer.Ordinal);
        AddAttribute(attributes, "pageNumber", pageNumber);
        AddAttribute(attributes, "destinationTop", destinationTop);
        AddAttribute(attributes, "destinationMode", destinationMode?.ToString());
        AddAttribute(attributes, "destinationLeft", destinationLeft);
        AddAttribute(attributes, "destinationBottom", destinationBottom);
        AddAttribute(attributes, "destinationRight", destinationRight);
        return attributes;
    }

    private static void AddAttribute(Dictionary<string, string> attributes, string name, int? value) {
        if (value.HasValue) {
            attributes[name] = value.Value.ToString(CultureInfo.InvariantCulture);
        }
    }

    private static void AddAttribute(Dictionary<string, string> attributes, string name, double? value) {
        if (value.HasValue) {
            attributes[name] = value.Value.ToString("R", CultureInfo.InvariantCulture);
        }
    }

    private static void AddAttribute(Dictionary<string, string> attributes, string name, string? value) {
        if (!string.IsNullOrWhiteSpace(value)) {
            attributes[name] = value!;
        }
    }

    private static HashSet<int>? BuildSelectedPageNumberSet(IReadOnlyList<PdfLogicalPage>? selectedPages) {
        if (selectedPages == null) {
            return null;
        }

        var pages = new HashSet<int>();
        for (int i = 0; i < selectedPages.Count; i++) {
            pages.Add(selectedPages[i].PageNumber);
        }

        return pages;
    }

    private static bool IsSelectedPage(int? pageNumber, HashSet<int>? selectedPageNumbers) {
        return selectedPageNumbers == null || !pageNumber.HasValue || selectedPageNumbers.Contains(pageNumber.Value);
    }

    private static int CountNamedDestinations(IReadOnlyList<PdfNamedDestination> destinations, HashSet<int>? selectedPageNumbers) {
        if (selectedPageNumbers == null) {
            return destinations.Count;
        }

        int count = 0;
        for (int i = 0; i < destinations.Count; i++) {
            if (IsSelectedPage(destinations[i].PageNumber, selectedPageNumbers)) {
                count++;
            }
        }

        return count;
    }

    private static int CountFormFields(IReadOnlyList<PdfFormField> fields, HashSet<int>? selectedPageNumbers) {
        if (selectedPageNumbers == null) {
            return fields.Count;
        }

        int count = 0;
        for (int i = 0; i < fields.Count; i++) {
            PdfFormField field = fields[i];
            for (int widgetIndex = 0; widgetIndex < field.Widgets.Count; widgetIndex++) {
                int? pageNumber = field.Widgets[widgetIndex].PageNumber;
                if (pageNumber.HasValue && selectedPageNumbers.Contains(pageNumber.Value)) {
                    count++;
                    break;
                }
            }
        }

        return count;
    }

    private static int CountOutlines(IReadOnlyList<PdfOutlineItem> outlines, HashSet<int>? selectedPageNumbers = null) {
        int count = 0;
        for (int i = 0; i < outlines.Count; i++) {
            if (IsSelectedPage(outlines[i].PageNumber, selectedPageNumbers)) {
                count++;
            }

            count += CountOutlines(outlines[i].Children, selectedPageNumbers);
        }

        return count;
    }

    private static string? ToMetadataText(bool? value) {
        return value.HasValue ? ToMetadataText(value.Value) : null;
    }

    private static string ToMetadataText(bool value) {
        return value ? "true" : "false";
    }

    private static string? ToMetadataText(int? value) {
        return value.HasValue ? value.Value.ToString(CultureInfo.InvariantCulture) : null;
    }

    private static ReaderLocation? BuildMetadataLocation(SourceMetadata source, int? pageNumber, string sourceBlockKind, string blockAnchor) {
        return pageNumber.HasValue ? BuildLocation(source, pageNumber.Value, 0, sourceBlockKind, blockAnchor) : null;
    }

    private static PdfLogicalHeading? FindHeading(PdfLogicalPage page, PdfLogicalTextBlock textBlock) {
        for (int i = 0; i < page.Headings.Count; i++) {
            if (ReferenceEquals(page.Headings[i].Line, textBlock)) {
                return page.Headings[i];
            }
        }

        return null;
    }

    private static PdfLogicalListItem? FindListItem(PdfLogicalPage page, PdfLogicalTextBlock textBlock) {
        for (int i = 0; i < page.ListItems.Count; i++) {
            if (ReferenceEquals(page.ListItems[i].Line, textBlock)) {
                return page.ListItems[i];
            }
        }

        return null;
    }

    private static string ToDocumentBlockKind(PdfLogicalElementKind kind) {
        switch (kind) {
            case PdfLogicalElementKind.Heading:
                return "heading";
            case PdfLogicalElementKind.ListItem:
                return "list-item";
            case PdfLogicalElementKind.LeaderRow:
                return "leader-row";
            case PdfLogicalElementKind.Table:
                return "table";
            default:
                return "text-block";
        }
    }

    private static string GetLinkKind(PdfLogicalLinkAnnotation link) {
        if (link.IsUriLink) return "uri";
        if (link.IsNamedDestinationLink || link.DestinationPageNumber.HasValue) return "destination";
        if (link.IsNamedActionLink) return "named-action";
        if (link.IsRemoteGoToLink) return "remote";
        return "link";
    }

    private static bool IsPageSelection(ReaderLocation location, int pageNumber, int selectionIndex) {
        return location.Page == pageNumber && location.SourceBlockIndex == selectionIndex;
    }

    private static ReaderLocation BuildLocation(SourceMetadata source, int pageNumber, int selectionIndex, string sourceBlockKind, string blockAnchor) {
        return new ReaderLocation {
            Path = source.Path,
            Page = pageNumber,
            SourceBlockIndex = selectionIndex,
            SourceBlockKind = sourceBlockKind,
            BlockAnchor = blockAnchor
        };
    }
}
