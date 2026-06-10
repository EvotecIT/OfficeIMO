using OfficeIMO.Visio;

namespace OfficeIMO.Reader.Visio;

public static partial class DocumentReaderVisioExtensions {
    /// <summary>
    /// Reads a Visio file and returns the shared OfficeIMO read result envelope.
    /// </summary>
    public static OfficeDocumentReadResult ReadVisioDocument(string visioPath, ReaderOptions? readerOptions = null, ReaderVisioOptions? visioOptions = null, CancellationToken cancellationToken = default) {
        if (visioPath == null) throw new ArgumentNullException(nameof(visioPath));
        if (visioPath.Length == 0) throw new ArgumentException("Visio path cannot be empty.", nameof(visioPath));
        if (!File.Exists(visioPath)) throw new FileNotFoundException($"Visio file '{visioPath}' doesn't exist.", visioPath);

        var effectiveReaderOptions = readerOptions ?? new ReaderOptions();
        var effectiveVisioOptions = ReaderVisioOptionsCloner.CloneOrDefault(visioOptions);
        ReaderInputLimits.EnforceFileSize(visioPath, effectiveReaderOptions.MaxInputBytes);
        SourceMetadata source = BuildSourceMetadataFromPath(visioPath, effectiveReaderOptions.ComputeHashes);
        VisioDocument document = VisioDocument.Load(visioPath);
        return BuildDocumentResult(document, source, effectiveReaderOptions, effectiveVisioOptions, cancellationToken);
    }

    /// <summary>
    /// Reads a Visio stream and returns the shared OfficeIMO read result envelope.
    /// </summary>
    public static OfficeDocumentReadResult ReadVisioDocument(Stream visioStream, string? sourceName = null, ReaderOptions? readerOptions = null, ReaderVisioOptions? visioOptions = null, CancellationToken cancellationToken = default) {
        if (visioStream == null) throw new ArgumentNullException(nameof(visioStream));
        if (!visioStream.CanRead) throw new ArgumentException("Visio stream must be readable.", nameof(visioStream));

        var effectiveReaderOptions = readerOptions ?? new ReaderOptions();
        var effectiveVisioOptions = ReaderVisioOptionsCloner.CloneOrDefault(visioOptions);
        var logicalSourceName = NormalizeLogicalSourceName(sourceName, "document.vsdx");
        var source = new SourceMetadata {
            Path = logicalSourceName,
            SourceId = BuildSourceId(logicalSourceName)
        };

        Stream parseStream = EnsureParseStream(visioStream, effectiveReaderOptions.MaxInputBytes, cancellationToken, out bool ownsParseStream);
        try {
            UpdateSourceMetadataFromSeekableStream(source, parseStream, effectiveReaderOptions.ComputeHashes);
            if (parseStream.CanSeek) {
                parseStream.Position = 0;
            }

            VisioDocument document = VisioDocument.Load(parseStream);
            return BuildDocumentResult(document, source, effectiveReaderOptions, effectiveVisioOptions, cancellationToken);
        } finally {
            if (ownsParseStream) {
                parseStream.Dispose();
            }
        }
    }

    /// <summary>
    /// Converts an already loaded Visio document into the shared OfficeIMO read result envelope.
    /// </summary>
    public static OfficeDocumentReadResult ReadVisioDocument(VisioDocument document, string sourceName = "document.vsdx", ReaderOptions? readerOptions = null, ReaderVisioOptions? visioOptions = null, CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (sourceName == null) throw new ArgumentNullException(nameof(sourceName));

        var effectiveReaderOptions = readerOptions ?? new ReaderOptions();
        var effectiveVisioOptions = ReaderVisioOptionsCloner.CloneOrDefault(visioOptions);
        var logicalSourceName = NormalizeLogicalSourceName(sourceName, "document.vsdx");
        var source = new SourceMetadata {
            Path = logicalSourceName,
            SourceId = BuildSourceId(logicalSourceName)
        };

        return BuildDocumentResult(document, source, effectiveReaderOptions, effectiveVisioOptions, cancellationToken);
    }

    /// <summary>
    /// Reads a Visio file and returns shape-data tables in source order.
    /// </summary>
    public static IReadOnlyList<ReaderTable> ReadVisioTables(string visioPath, ReaderOptions? readerOptions = null, ReaderVisioOptions? visioOptions = null, CancellationToken cancellationToken = default) {
        return DocumentReader.ExtractTables(ReadVisioFile(visioPath, readerOptions, visioOptions, cancellationToken), cancellationToken);
    }

    /// <summary>
    /// Reads a Visio stream and returns shape-data tables in source order.
    /// </summary>
    public static IReadOnlyList<ReaderTable> ReadVisioTables(Stream visioStream, string? sourceName = null, ReaderOptions? readerOptions = null, ReaderVisioOptions? visioOptions = null, CancellationToken cancellationToken = default) {
        return DocumentReader.ExtractTables(ReadVisio(visioStream, sourceName, readerOptions, visioOptions, cancellationToken), cancellationToken);
    }

    /// <summary>
    /// Converts an already loaded Visio document into shape-data tables in source order.
    /// </summary>
    public static IReadOnlyList<ReaderTable> ReadVisioTables(VisioDocument document, string sourceName = "document.vsdx", ReaderOptions? readerOptions = null, ReaderVisioOptions? visioOptions = null, CancellationToken cancellationToken = default) {
        return DocumentReader.ExtractTables(ReadVisio(document, sourceName, readerOptions, visioOptions, cancellationToken), cancellationToken);
    }

    /// <summary>
    /// Reads a Visio file and returns shape-data table export payloads in source order.
    /// </summary>
    public static IReadOnlyList<ReaderTableExportBundle> ReadVisioTableExports(string visioPath, ReaderOptions? readerOptions = null, ReaderVisioOptions? visioOptions = null, bool indentedJson = false, CancellationToken cancellationToken = default) {
        return DocumentReader.ExportTables(ReadVisioTables(visioPath, readerOptions, visioOptions, cancellationToken), indentedJson, cancellationToken);
    }

    /// <summary>
    /// Reads a Visio stream and returns shape-data table export payloads in source order.
    /// </summary>
    public static IReadOnlyList<ReaderTableExportBundle> ReadVisioTableExports(Stream visioStream, string? sourceName = null, ReaderOptions? readerOptions = null, ReaderVisioOptions? visioOptions = null, bool indentedJson = false, CancellationToken cancellationToken = default) {
        return DocumentReader.ExportTables(ReadVisioTables(visioStream, sourceName, readerOptions, visioOptions, cancellationToken), indentedJson, cancellationToken);
    }

    /// <summary>
    /// Converts an already loaded Visio document into shape-data table export payloads in source order.
    /// </summary>
    public static IReadOnlyList<ReaderTableExportBundle> ReadVisioTableExports(VisioDocument document, string sourceName = "document.vsdx", ReaderOptions? readerOptions = null, ReaderVisioOptions? visioOptions = null, bool indentedJson = false, CancellationToken cancellationToken = default) {
        return DocumentReader.ExportTables(ReadVisioTables(document, sourceName, readerOptions, visioOptions, cancellationToken), indentedJson, cancellationToken);
    }

    /// <summary>
    /// Reads a Visio file and returns the shared OfficeIMO read result JSON envelope.
    /// </summary>
    public static string ReadVisioDocumentJson(string visioPath, ReaderOptions? readerOptions = null, ReaderVisioOptions? visioOptions = null, bool indented = false, CancellationToken cancellationToken = default) {
        return OfficeDocumentReadResultJson.Serialize(ReadVisioDocument(visioPath, readerOptions, visioOptions, cancellationToken), indented);
    }

    /// <summary>
    /// Reads a Visio stream and returns the shared OfficeIMO read result JSON envelope.
    /// </summary>
    public static string ReadVisioDocumentJson(Stream visioStream, string? sourceName = null, ReaderOptions? readerOptions = null, ReaderVisioOptions? visioOptions = null, bool indented = false, CancellationToken cancellationToken = default) {
        return OfficeDocumentReadResultJson.Serialize(ReadVisioDocument(visioStream, sourceName, readerOptions, visioOptions, cancellationToken), indented);
    }

    /// <summary>
    /// Converts an already loaded Visio document into the shared OfficeIMO read result JSON envelope.
    /// </summary>
    public static string ReadVisioDocumentJson(VisioDocument document, string sourceName = "document.vsdx", ReaderOptions? readerOptions = null, ReaderVisioOptions? visioOptions = null, bool indented = false, CancellationToken cancellationToken = default) {
        return OfficeDocumentReadResultJson.Serialize(ReadVisioDocument(document, sourceName, readerOptions, visioOptions, cancellationToken), indented);
    }

    private static OfficeDocumentReadResult BuildDocumentResult(VisioDocument document, SourceMetadata source, ReaderOptions readerOptions, ReaderVisioOptions visioOptions, CancellationToken cancellationToken) {
        VisioInspectionSnapshot snapshot = document.CreateInspectionSnapshot();
        ReaderChunk[] chunks = ReadVisio(document, source, readerOptions, cancellationToken).ToArray();
        ReaderTable[] tables = chunks
            .Where(static chunk => chunk.Tables != null)
            .SelectMany(static chunk => chunk.Tables!)
            .ToArray();
        OfficeDocumentBlock[] blocks = BuildDocumentBlocks(snapshot, source).ToArray();
        VisioPage[] snapshotOrderedPages = GetSnapshotOrderedPages(document, snapshot).ToArray();
        OfficeDocumentLink[] links = BuildDocumentLinks(snapshotOrderedPages, source).ToArray();
        OfficeDocumentAsset[] assets = BuildDocumentAssets(snapshotOrderedPages, source, visioOptions, cancellationToken).ToArray();

        return new OfficeDocumentReadResult {
            Kind = ReaderInputKind.Visio,
            Source = new OfficeDocumentSource {
                Path = source.Path,
                SourceId = source.SourceId,
                SourceHash = source.SourceHash,
                LastWriteUtc = source.LastWriteUtc,
                LengthBytes = source.LengthBytes,
                Title = document.Title,
                Author = document.Author
            },
            CapabilitiesUsed = BuildDocumentCapabilities(visioOptions),
            Markdown = chunks.Length == 0 ? null : string.Join(Environment.NewLine + Environment.NewLine, chunks.Select(static chunk => chunk.Markdown ?? chunk.Text)),
            Chunks = chunks,
            Pages = BuildDocumentPages(snapshot, source, blocks, tables, links, assets),
            Blocks = blocks,
            Tables = tables,
            Assets = assets,
            Links = links,
            Forms = Array.Empty<OfficeDocumentFormField>(),
            Visuals = Array.Empty<ReaderVisual>(),
            Diagnostics = Array.Empty<OfficeDocumentDiagnostic>()
        };
    }

    private static IReadOnlyList<string> BuildDocumentCapabilities(ReaderVisioOptions options) {
        var capabilities = new List<string> {
            "officeimo.reader.visio",
            "officeimo.visio.inspection-snapshot"
        };
        if (options.IncludeSvgPreviewAssets) {
            capabilities.Add("officeimo.visio.svg-preview");
        }
        if (options.IncludePngPreviewAssets) {
            capabilities.Add("officeimo.visio.png-preview");
        }

        return capabilities;
    }

    private static IEnumerable<OfficeDocumentBlock> BuildDocumentBlocks(VisioInspectionSnapshot snapshot, SourceMetadata source) {
        for (int pageIndex = 0; pageIndex < snapshot.Pages.Count; pageIndex++) {
            VisioInspectionPageSnapshot page = snapshot.Pages[pageIndex];
            foreach (VisioInspectionShapeSnapshot shape in page.Shapes) {
                yield return new OfficeDocumentBlock {
                    Id = "visio-page-" + (pageIndex + 1).ToString("D4", CultureInfo.InvariantCulture) + "-shape-" + shape.Id,
                    Kind = ResolveShapeKind(shape),
                    Text = BuildShapeBlockText(shape),
                    Location = BuildLocation(source, pageIndex, "shape", "page-" + (pageIndex + 1).ToString(CultureInfo.InvariantCulture) + "-shape-" + shape.Id),
                    Region = new OfficeDocumentRegion {
                        X = shape.PinX - (shape.Width / 2D),
                        Y = shape.PinY - (shape.Height / 2D),
                        Width = shape.Width,
                        Height = shape.Height
                    }
                };
            }

            foreach (VisioInspectionConnectorSnapshot connector in page.Connectors) {
                yield return new OfficeDocumentBlock {
                    Id = "visio-page-" + (pageIndex + 1).ToString("D4", CultureInfo.InvariantCulture) + "-connector-" + connector.Id,
                    Kind = "connector",
                    Text = BuildConnectorBlockText(connector),
                    Location = BuildLocation(source, pageIndex, "connector", "page-" + (pageIndex + 1).ToString(CultureInfo.InvariantCulture) + "-connector-" + connector.Id)
                };
            }
        }
    }

    private static IReadOnlyList<OfficeDocumentPage> BuildDocumentPages(
        VisioInspectionSnapshot snapshot,
        SourceMetadata source,
        IReadOnlyList<OfficeDocumentBlock> blocks,
        IReadOnlyList<ReaderTable> tables,
        IReadOnlyList<OfficeDocumentLink> links,
        IReadOnlyList<OfficeDocumentAsset> assets) {
        var pages = new List<OfficeDocumentPage>(snapshot.Pages.Count);
        for (int pageIndex = 0; pageIndex < snapshot.Pages.Count; pageIndex++) {
            VisioInspectionPageSnapshot page = snapshot.Pages[pageIndex];
            int pageNumber = pageIndex + 1;
            pages.Add(new OfficeDocumentPage {
                Number = pageNumber,
                Name = page.Name,
                Width = page.Width,
                Height = page.Height,
                Location = BuildLocation(source, pageIndex, "page", "page-" + pageNumber.ToString(CultureInfo.InvariantCulture)),
                Blocks = blocks.Where(block => block.Location.Page == pageNumber).ToArray(),
                Tables = tables.Where(table => table.Location?.Page == pageNumber).ToArray(),
                Assets = assets.Where(asset => asset.Location.Page == pageNumber).ToArray(),
                Links = links.Where(link => link.Location.Page == pageNumber).ToArray(),
                Forms = Array.Empty<OfficeDocumentFormField>()
            });
        }

        return pages;
    }

    private static IEnumerable<VisioPage> GetSnapshotOrderedPages(VisioDocument document, VisioInspectionSnapshot snapshot) {
        for (int pageIndex = 0; pageIndex < snapshot.Pages.Count; pageIndex++) {
            int pageId = snapshot.Pages[pageIndex].Id;
            VisioPage? page = document.Pages.FirstOrDefault(candidate => candidate.Id == pageId);
            if (page != null) {
                yield return page;
            }
        }
    }

    private static IEnumerable<OfficeDocumentLink> BuildDocumentLinks(IReadOnlyList<VisioPage> pages, SourceMetadata source) {
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++) {
            VisioPage page = pages[pageIndex];
            foreach (VisioShape shape in page.AllShapes()) {
                for (int linkIndex = 0; linkIndex < shape.Hyperlinks.Count; linkIndex++) {
                    VisioHyperlink link = shape.Hyperlinks[linkIndex];
                    yield return BuildLink(
                        id: "visio-page-" + (pageIndex + 1).ToString("D4", CultureInfo.InvariantCulture) + "-shape-" + shape.Id + "-link-" + linkIndex.ToString("D4", CultureInfo.InvariantCulture),
                        link,
                        source,
                        pageIndex,
                        ownerKind: "shape",
                        ownerId: shape.Id,
                        region: new OfficeDocumentRegion {
                            X = shape.PinX - (shape.Width / 2D),
                            Y = shape.PinY - (shape.Height / 2D),
                            Width = shape.Width,
                            Height = shape.Height
                        });
                }
            }

            foreach (VisioConnector connector in page.Connectors) {
                for (int linkIndex = 0; linkIndex < connector.Hyperlinks.Count; linkIndex++) {
                    VisioHyperlink link = connector.Hyperlinks[linkIndex];
                    yield return BuildLink(
                        id: "visio-page-" + (pageIndex + 1).ToString("D4", CultureInfo.InvariantCulture) + "-connector-" + connector.Id + "-link-" + linkIndex.ToString("D4", CultureInfo.InvariantCulture),
                        link,
                        source,
                        pageIndex,
                        ownerKind: "connector",
                        ownerId: connector.Id,
                        region: null);
                }
            }
        }
    }

    private static OfficeDocumentLink BuildLink(string id, VisioHyperlink link, SourceMetadata source, int pageIndex, string ownerKind, string ownerId, OfficeDocumentRegion? region) {
        return new OfficeDocumentLink {
            Id = id,
            Kind = string.IsNullOrWhiteSpace(link.Address) ? "internal" : "uri",
            Uri = link.Address,
            DestinationName = link.SubAddress,
            Text = link.Description,
            Location = BuildLocation(source, pageIndex, ownerKind + "-hyperlink", "page-" + (pageIndex + 1).ToString(CultureInfo.InvariantCulture) + "-" + ownerKind + "-" + ownerId + "-link"),
            Region = region
        };
    }

    private static IEnumerable<OfficeDocumentAsset> BuildDocumentAssets(IReadOnlyList<VisioPage> pages, SourceMetadata source, ReaderVisioOptions options, CancellationToken cancellationToken) {
        if (!options.IncludeSvgPreviewAssets && !options.IncludePngPreviewAssets) {
            yield break;
        }

        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            VisioPage page = pages[pageIndex];
            if (options.IncludeSvgPreviewAssets) {
                byte[] svgBytes = Encoding.UTF8.GetBytes(page.ToSvg(options.SvgOptions));
                yield return BuildPreviewAsset(source, pageIndex, "preview-svg", "image/svg+xml", ".svg", svgBytes);
            }

            if (options.IncludePngPreviewAssets) {
                byte[] pngBytes = page.ToPng(options.PngOptions);
                yield return BuildPreviewAsset(source, pageIndex, "preview-png", "image/png", ".png", pngBytes);
            }
        }
    }

    private static OfficeDocumentAsset BuildPreviewAsset(SourceMetadata source, int pageIndex, string kind, string mediaType, string extension, byte[] payload) {
        string assetId = "visio-page-" + (pageIndex + 1).ToString("D4", CultureInfo.InvariantCulture) + "-" + kind;
        return new OfficeDocumentAsset {
            Id = assetId,
            Kind = kind,
            MediaType = mediaType,
            Extension = extension,
            FileName = OfficeDocumentAssetNaming.BuildFileName(assetId, extension),
            LengthBytes = payload.LongLength,
            PayloadHash = ComputeSha256Hex(payload),
            PayloadBytes = payload,
            Location = BuildLocation(source, pageIndex, kind, "page-" + (pageIndex + 1).ToString(CultureInfo.InvariantCulture) + "-" + kind)
        };
    }

    private static string ResolveShapeKind(VisioInspectionShapeSnapshot shape) {
        if (shape.IsContainer) return "container";
        if (shape.IsCallout) return "callout";
        if (shape.IsBackgroundSurface) return "background";
        if (shape.IsDiagramAdornment) return "adornment";
        if (string.Equals(shape.Type, "Group", StringComparison.OrdinalIgnoreCase)) return "group";
        return "shape";
    }

    private static string BuildShapeBlockText(VisioInspectionShapeSnapshot shape) {
        var builder = new StringBuilder();
        builder.Append(string.IsNullOrWhiteSpace(shape.Text) ? shape.Id : shape.Text);
        if (!string.IsNullOrWhiteSpace(shape.MasterNameU)) {
            builder.Append(" [");
            builder.Append(shape.MasterNameU);
            builder.Append(']');
        }
        if (shape.ShapeData.Count > 0) {
            builder.Append(" ");
            builder.Append(string.Join("; ", shape.ShapeData.Select(FormatShapeData)));
        }
        return builder.ToString();
    }

    private static string BuildConnectorBlockText(VisioInspectionConnectorSnapshot connector) {
        var builder = new StringBuilder();
        builder.Append(connector.FromId);
        builder.Append(" -> ");
        builder.Append(connector.ToId);
        if (!string.IsNullOrWhiteSpace(connector.Label)) {
            builder.Append(": ");
            builder.Append(connector.Label);
        }
        if (connector.ShapeData.Count > 0) {
            builder.Append(" ");
            builder.Append(string.Join("; ", connector.ShapeData.Select(FormatShapeData)));
        }
        return builder.ToString();
    }
}
