using OfficeIMO.Html;
using OfficeIMO.Markdown.Html;
using OfficeIMO.Reader;
using System.Linq;
using System.Security.Cryptography;

namespace OfficeIMO.Reader.Html;

internal static partial class HtmlReaderAdapter {
    /// <summary>Reads an HTML file into the shared rich document envelope.</summary>
    public static OfficeDocumentReadResult ReadDocument(string htmlPath, ReaderOptions? readerOptions = null, ReaderHtmlOptions? htmlOptions = null, CancellationToken cancellationToken = default) {
        if (htmlPath == null) throw new ArgumentNullException(nameof(htmlPath));
        if (htmlPath.Length == 0) throw new ArgumentException("HTML path cannot be empty.", nameof(htmlPath));
        if (!File.Exists(htmlPath)) throw new FileNotFoundException($"HTML file '{htmlPath}' doesn't exist.", htmlPath);
        ReaderOptions effective = readerOptions ?? new ReaderOptions();
        ReaderInputLimits.EnforceFileSize(htmlPath, effective.MaxInputBytes);
        SourceMetadata source = BuildSourceMetadataFromPath(htmlPath, effective.ComputeHashes);
        if (IsMhtmlSource(htmlPath)) {
            MhtmlDocument archive = MhtmlDocument.Load(htmlPath, CreateMhtmlReaderOptions(effective),
                cancellationToken: cancellationToken);
            return BuildHtmlDocumentResult(archive.Html, source, effective,
                PrepareMhtmlHtmlOptions(htmlOptions, archive), cancellationToken, archive);
        }
        using var stream = new FileStream(htmlPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        string html = ReadAllText(stream, cancellationToken);
        return BuildHtmlDocumentResult(html, source, effective, htmlOptions, cancellationToken);
    }

    /// <summary>Reads an HTML stream into the shared rich document envelope.</summary>
    public static OfficeDocumentReadResult ReadDocument(Stream htmlStream, string? sourceName = null, ReaderOptions? readerOptions = null, ReaderHtmlOptions? htmlOptions = null, CancellationToken cancellationToken = default) {
        if (htmlStream == null) throw new ArgumentNullException(nameof(htmlStream));
        if (!htmlStream.CanRead) throw new ArgumentException("HTML stream must be readable.", nameof(htmlStream));
        ReaderOptions effective = readerOptions ?? new ReaderOptions();
        string logicalSourceName = string.IsNullOrWhiteSpace(sourceName) ? "document.html" : sourceName!.Trim();
        var source = new SourceMetadata { Path = logicalSourceName, SourceId = BuildSourceId(logicalSourceName) };
        Stream parseStream = ReaderInputLimits.EnsureSeekableReadStream(htmlStream, effective.MaxInputBytes, cancellationToken, out bool ownsParseStream);
        try {
            UpdateSourceMetadataFromSeekableStream(source, parseStream, effective.ComputeHashes);
            MhtmlDocument? archive = IsMhtmlSource(logicalSourceName)
                ? LoadMhtml(parseStream, effective, cancellationToken)
                : null;
            string html = archive?.Html ?? ReadAllText(parseStream, cancellationToken);
            return BuildHtmlDocumentResult(html, source, effective,
                archive == null ? htmlOptions : PrepareMhtmlHtmlOptions(htmlOptions, archive),
                cancellationToken, archive);
        } finally {
            if (ownsParseStream) parseStream.Dispose();
        }
    }

    /// <summary>Reads an HTML string into the shared rich document envelope.</summary>
    public static OfficeDocumentReadResult ReadContentDocument(string html, string sourceName = "document.html", ReaderOptions? readerOptions = null, ReaderHtmlOptions? htmlOptions = null, CancellationToken cancellationToken = default) {
        if (html == null) throw new ArgumentNullException(nameof(html));
        if (sourceName == null) throw new ArgumentNullException(nameof(sourceName));
        ReaderOptions effective = readerOptions ?? new ReaderOptions();
        string logicalSourceName = string.IsNullOrWhiteSpace(sourceName) ? "document.html" : sourceName.Trim();
        SourceMetadata source = BuildSourceMetadataFromHtmlString(logicalSourceName, html, effective.ComputeHashes);
        return BuildHtmlDocumentResult(html, source, effective, htmlOptions, cancellationToken);
    }

    /// <summary>Reads an HTML file into the shared rich document JSON envelope.</summary>
    public static string ReadDocumentJson(string htmlPath, ReaderOptions? readerOptions = null, ReaderHtmlOptions? htmlOptions = null, bool indented = false, CancellationToken cancellationToken = default) {
        return OfficeDocumentReadResultJson.Serialize(ReadDocument(htmlPath, readerOptions, htmlOptions, cancellationToken), indented);
    }

    /// <summary>Reads an HTML stream into the shared rich document JSON envelope.</summary>
    public static string ReadDocumentJson(Stream htmlStream, string? sourceName = null, ReaderOptions? readerOptions = null, ReaderHtmlOptions? htmlOptions = null, bool indented = false, CancellationToken cancellationToken = default) {
        return OfficeDocumentReadResultJson.Serialize(ReadDocument(htmlStream, sourceName, readerOptions, htmlOptions, cancellationToken), indented);
    }

    /// <summary>Reads an HTML string into the shared rich document JSON envelope.</summary>
    public static string ReadContentDocumentJson(string html, string sourceName = "document.html", ReaderOptions? readerOptions = null, ReaderHtmlOptions? htmlOptions = null, bool indented = false, CancellationToken cancellationToken = default) {
        return OfficeDocumentReadResultJson.Serialize(ReadContentDocument(html, sourceName, readerOptions, htmlOptions, cancellationToken), indented);
    }

    private static OfficeDocumentReadResult BuildHtmlDocumentResult(string html, SourceMetadata source,
        ReaderOptions readerOptions, ReaderHtmlOptions? htmlOptions, CancellationToken cancellationToken,
        MhtmlDocument? archive = null) {
        ReaderHtmlOptions effectiveHtmlOptions = ReaderHtmlOptionsCloner.CloneOrDefault(htmlOptions);
        HtmlToMarkdownOptions projectionOptions = effectiveHtmlOptions.HtmlToMarkdownOptions ?? HtmlToMarkdownOptions.CreateOfficeIMOProfile();
        bool hasProjectionFilters = projectionOptions.ExcludeSelectors.Count > 0 || projectionOptions.ElementFilters.Count > 0;
        string projectedHtml = html;
        ReaderHtmlOptions chunkHtmlOptions = effectiveHtmlOptions;
        HtmlConversionDocument conversionDocument = HtmlConversionDocument.Parse(
            html,
            new HtmlConversionDocumentOptions {
                BaseUri = projectionOptions.BaseUri,
                UseBodyContentsOnly = false,
                IncludeNormalizedHtml = false
            });
        var filtered = HtmlToMarkdownConverter.PrepareDocument(
            conversionDocument.CreateDocumentForConversion(HtmlCssMediaContext.Screen),
            projectionOptions);
        projectionOptions.BaseUri = HtmlDocumentParser.ResolveEffectiveBaseUri(filtered, projectionOptions.BaseUri);
        HtmlLogicalDocument logical = HtmlLogicalDocumentBuilder.FromDocument(filtered, useBodyContentsOnly: false);
        if (hasProjectionFilters) {
            projectedHtml = filtered.DocumentElement?.OuterHtml ?? html;
            chunkHtmlOptions = effectiveHtmlOptions.Clone();
            chunkHtmlOptions.HtmlToMarkdownOptions?.ExcludeSelectors.Clear();
            chunkHtmlOptions.HtmlToMarkdownOptions?.ElementFilters.Clear();
        }
        ReaderChunk[] chunks = ReadContent(projectedHtml, source, readerOptions, chunkHtmlOptions, cancellationToken).ToArray();
        HtmlProjection projection = ProjectHtml(logical, source.Path, readerOptions.MaxTableRows, projectionOptions, cancellationToken);
        if (archive != null) MergeMhtmlResources(projection, archive, source.Path);
        var documentSource = new OfficeDocumentSource {
            Path = source.Path,
            SourceId = source.SourceId,
            SourceHash = source.SourceHash,
            LastWriteUtc = source.LastWriteUtc,
            LengthBytes = source.LengthBytes,
            Title = FindHtmlMetadata(logical.Root, "title", null) ?? archive?.Subject,
            Author = FindHtmlMetadata(logical.Root, "meta", "author"),
            Subject = FindHtmlMetadata(logical.Root, "meta", "description"),
            Keywords = FindHtmlMetadata(logical.Root, "meta", "keywords")
        };
        OfficeDocumentReadResult result = DocumentReaderEngine.CreateDocumentResult(
            chunks,
            ReaderInputKind.Html,
            documentSource,
            new[] { "officeimo.reader.html.rich-v5", "officeimo.html.logical-document" }
                .Concat(archive == null ? Array.Empty<string>() : new[] { "officeimo.html.mhtml" })
                .Concat(logical.Capabilities.Select(static capability => "officeimo.html." + capability)),
            projection.Assets);
        result.Html = projectedHtml;
        result.Blocks = projection.Blocks;
        result.Tables = projection.Tables;
        result.Links = projection.Links;
        result.Forms = projection.Forms;
        result.Visuals = projection.Visuals;
        result.Metadata = BuildHtmlMetadata(logical, projection);
        if (archive != null) result.Diagnostics = result.Diagnostics.Concat(MapMhtmlDiagnostics(archive, source.Path)).ToArray();
        return result;
    }

    private static HtmlProjection ProjectHtml(
        HtmlLogicalDocument document,
        string? path,
        int maxTableRows,
        HtmlToMarkdownOptions htmlOptions,
        CancellationToken cancellationToken) {
        var projection = new HtmlProjection();
        int blockIndex = 0;
        int tableIndex = 0;
        int linkIndex = 0;
        int assetIndex = 0;
        int formIndex = 0;
        TraverseHtml(document.Root, null, 0, path, maxTableRows, htmlOptions, projection, ref blockIndex, ref tableIndex, ref linkIndex, ref assetIndex, ref formIndex, cancellationToken);
        return projection;
    }

    private static void TraverseHtml(
        HtmlLogicalNode node,
        string? listName,
        int listLevel,
        string? path,
        int maxTableRows,
        HtmlToMarkdownOptions htmlOptions,
        HtmlProjection projection,
        ref int blockIndex,
        ref int tableIndex,
        ref int linkIndex,
        ref int assetIndex,
        ref int formIndex,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        string? nextListName = node.Kind == HtmlLogicalNodeKind.List ? node.Name : listName;
        int nextListLevel = node.Kind == HtmlLogicalNodeKind.List ? listLevel + 1 : listLevel;
        ReaderTable? mappedTable = node.Kind == HtmlLogicalNodeKind.Table
            ? MapHtmlTable(node, path, tableIndex++, maxTableRows)
            : null;
        if (IsHtmlBlock(node.Kind)) {
            string kind = NormalizeHtmlBlockKind(node.Kind);
            string anchor = "html-" + kind + "-" + blockIndex.ToString("D4", CultureInfo.InvariantCulture);
            projection.Blocks.Add(new OfficeDocumentBlock {
                Id = anchor,
                Kind = kind,
                Text = mappedTable == null ? GetHtmlNodeText(node) : BuildHtmlTableBlockText(mappedTable),
                Level = node.Kind == HtmlLogicalNodeKind.Heading ? ParseHtmlHeadingLevel(node.Name) : node.Kind == HtmlLogicalNodeKind.ListItem ? nextListLevel : null,
                Marker = node.Kind == HtmlLogicalNodeKind.ListItem ? (string.Equals(nextListName, "ol", StringComparison.OrdinalIgnoreCase) ? "1." : "-") : null,
                Location = BuildHtmlLocation(path, blockIndex, kind, anchor)
            });
            blockIndex++;
        }
        if (mappedTable != null) projection.Tables.Add(mappedTable);
        if (node.Kind == HtmlLogicalNodeKind.Link && node.Attributes.TryGetValue("href", out string? href)) {
            string resolvedHref = HtmlUrlPolicyEvaluator.ResolveUrl(href, htmlOptions.BaseUri, htmlOptions.UrlPolicy);
            if (!string.IsNullOrWhiteSpace(resolvedHref)) {
                projection.Links.Add(new OfficeDocumentLink {
                    Id = "html-link-" + linkIndex.ToString("D4", CultureInfo.InvariantCulture),
                    Kind = resolvedHref.StartsWith("#", StringComparison.Ordinal) ? "internal" : "uri",
                    Uri = resolvedHref.StartsWith("#", StringComparison.Ordinal) ? null : resolvedHref,
                    DestinationName = resolvedHref.StartsWith("#", StringComparison.Ordinal) ? resolvedHref.Substring(1) : null,
                    Text = GetHtmlNodeText(node),
                    Location = BuildHtmlLocation(path, null, "hyperlink", "html-link-" + linkIndex.ToString("D4", CultureInfo.InvariantCulture))
                });
                linkIndex++;
            }
        }
        if (node.Kind == HtmlLogicalNodeKind.Image) {
            OfficeDocumentAsset? asset = MapHtmlImage(node, path, assetIndex, htmlOptions);
            if (asset != null) {
                assetIndex++;
                projection.Assets.Add(asset);
                projection.Visuals.Add(MapHtmlVisual(node, asset.Location, asset.PayloadHash, asset.MediaType, asset.SourceObjectId));
            }
        } else if (node.Kind == HtmlLogicalNodeKind.Media &&
            !string.Equals(node.Name, "source", StringComparison.OrdinalIgnoreCase)) {
            string anchor = "html-media-" + projection.Visuals.Count.ToString("D4", CultureInfo.InvariantCulture);
            projection.Visuals.Add(MapHtmlVisual(node, BuildHtmlLocation(path, null, "media", anchor), null, null));
        }
        if (node.Kind == HtmlLogicalNodeKind.FormControl &&
            !string.Equals(node.Name, "option", StringComparison.OrdinalIgnoreCase)) {
            projection.Forms.Add(MapHtmlFormControl(node, path, formIndex++));
        }
        foreach (HtmlLogicalNode child in node.Children) {
            TraverseHtml(child, nextListName, nextListLevel, path, maxTableRows, htmlOptions, projection, ref blockIndex, ref tableIndex, ref linkIndex, ref assetIndex, ref formIndex, cancellationToken);
        }
    }

    private static ReaderTable MapHtmlTable(HtmlLogicalNode table, string? path, int tableIndex, int maxRows) {
        List<HtmlLogicalNode> rows = GetHtmlTableRows(table).ToList();
        int columnCount = rows.Count == 0 ? 0 : rows.Max(row => row.Children.Count(child => child.Kind == HtmlLogicalNodeKind.TableCell));
        bool hasHeaderRow = rows.Count > 0 && rows[0].Children.Any(child =>
            child.Kind == HtmlLogicalNodeKind.TableCell && string.Equals(child.Name, "th", StringComparison.OrdinalIgnoreCase));
        IReadOnlyList<string> columns = hasHeaderRow
            ? BuildHtmlTableRow(rows[0], columnCount, true)
            : Enumerable.Range(1, columnCount)
                .Select(index => "Column " + index.ToString(CultureInfo.InvariantCulture))
                .ToArray();
        int dataStart = hasHeaderRow ? 1 : 0;
        int totalRows = Math.Max(0, rows.Count - dataStart);
        int emittedRows = maxRows > 0 ? Math.Min(totalRows, maxRows) : totalRows;
        IReadOnlyList<IReadOnlyList<string>> values = rows.Skip(dataStart).Take(emittedRows).Select(row => BuildHtmlTableRow(row, columnCount, false)).ToArray();
        return new ReaderTable {
            Title = table.Children.FirstOrDefault(child => child.Kind == HtmlLogicalNodeKind.TableCaption)?.Text ?? "HTML table " + (tableIndex + 1).ToString(CultureInfo.InvariantCulture),
            Kind = "html-table",
            Location = BuildHtmlLocation(path, null, "table", "html-table-" + tableIndex.ToString("D4", CultureInfo.InvariantCulture), tableIndex),
            Columns = columns,
            ColumnProfiles = ReaderTableProfiler.CreateProfiles(columns, values),
            Rows = values,
            TotalRowCount = totalRows,
            Truncated = emittedRows < totalRows
        };
    }

    private static IReadOnlyList<string> BuildHtmlTableRow(HtmlLogicalNode row, int columnCount, bool fallbacks) {
        HtmlLogicalNode[] cells = row.Children.Where(child => child.Kind == HtmlLogicalNodeKind.TableCell).ToArray();
        return Enumerable.Range(0, columnCount).Select(index => {
            string value = index < cells.Length ? GetHtmlNodeText(cells[index]) : string.Empty;
            return string.IsNullOrWhiteSpace(value) && fallbacks ? "Column " + (index + 1).ToString(CultureInfo.InvariantCulture) : value;
        }).ToArray();
    }

    private static OfficeDocumentAsset? MapHtmlImage(HtmlLogicalNode node, string? path, int index, HtmlToMarkdownOptions htmlOptions) {
        node.Attributes.TryGetValue("alt", out string? altText);
        node.Attributes.TryGetValue("title", out string? title);
        string resolvedSource = ResolveHtmlImageSource(node, htmlOptions);
        if (string.IsNullOrWhiteSpace(resolvedSource)) return null;
        byte[]? payload = null;
        string? mediaType = null;
        string? extension = null;
        if (HtmlImageDataUri.TryParse(resolvedSource, out HtmlImageDataUri dataUri)) {
            if (dataUri.IsBase64 && htmlOptions.Base64Images != HtmlBase64ImageHandling.Include) return null;
            if (!dataUri.TryDecodeBytes(out byte[] bytes)) return null;
            payload = bytes;
            mediaType = dataUri.MediaType;
            extension = dataUri.FileExtension;
        }
        string id = "html-image-" + index.ToString("D4", CultureInfo.InvariantCulture);
        return new OfficeDocumentAsset {
            Id = id,
            Kind = "image",
            MediaType = mediaType,
            Extension = extension,
            FileName = extension == null ? null : OfficeDocumentAssetNaming.BuildFileName(id, extension),
            AltText = altText,
            Title = title,
            LengthBytes = payload?.LongLength,
            PayloadHash = payload == null ? null : ComputeHtmlHash(payload),
            PayloadBytes = payload,
            SourceObjectId = payload == null ? resolvedSource : "data-uri",
            Location = BuildHtmlLocation(path, null, "image", id)
        };
    }

    private static OfficeDocumentFormField MapHtmlFormControl(HtmlLogicalNode node, string? path, int index) {
        node.Attributes.TryGetValue("name", out string? name);
        bool hasValue = node.Attributes.TryGetValue("value", out string? value);
        string kind = node.Attributes.TryGetValue("type", out string? type) && !string.IsNullOrWhiteSpace(type)
            ? type
            : node.Name;
        if (!hasValue && string.Equals(node.Name, "textarea", StringComparison.OrdinalIgnoreCase)) {
            value = GetHtmlNodeText(node);
        } else if (!hasValue && string.Equals(node.Name, "select", StringComparison.OrdinalIgnoreCase)) {
            HtmlLogicalNode[] options = EnumerateHtmlOptions(node).ToArray();
            HtmlLogicalNode[] selected = options.Where(option => option.Attributes.ContainsKey("selected")).ToArray();
            if (selected.Length == 0 && options.Length > 0) selected = new[] { options[0] };
            value = string.Join("\n", selected.Select(option =>
                option.Attributes.TryGetValue("value", out string? optionValue) && !string.IsNullOrWhiteSpace(optionValue)
                    ? optionValue
                    : GetHtmlNodeText(option)));
        }
        if (string.Equals(kind, "checkbox", StringComparison.OrdinalIgnoreCase)
            || string.Equals(kind, "radio", StringComparison.OrdinalIgnoreCase)) {
            value = node.Attributes.ContainsKey("checked")
                ? hasValue ? value : "on"
                : null;
        }
        string id = "html-form-" + index.ToString("D4", CultureInfo.InvariantCulture);
        return new OfficeDocumentFormField {
            Id = id,
            Name = name,
            Kind = kind,
            Value = value,
            IsReadOnly = node.Attributes.ContainsKey("readonly") || node.Attributes.ContainsKey("disabled"),
            IsRequired = node.Attributes.ContainsKey("required"),
            Location = BuildHtmlLocation(path, null, "form-control", id)
        };
    }

    private static IEnumerable<HtmlLogicalNode> EnumerateHtmlOptions(HtmlLogicalNode node) {
        foreach (HtmlLogicalNode child in node.Children) {
            if (child.Kind == HtmlLogicalNodeKind.FormControl &&
                string.Equals(child.Name, "option", StringComparison.OrdinalIgnoreCase)) {
                yield return child;
            }
            foreach (HtmlLogicalNode descendant in EnumerateHtmlOptions(child)) yield return descendant;
        }
    }

    private static string BuildHtmlTableBlockText(ReaderTable table) {
        IEnumerable<IReadOnlyList<string>> rows = table.Columns.Count == 0
            ? table.Rows
            : new[] { table.Columns }.Concat(table.Rows);
        return string.Join(Environment.NewLine, rows.Select(static row => string.Join(" | ", row)));
    }

    private static string ResolveHtmlImageSource(HtmlLogicalNode node, HtmlToMarkdownOptions options) {
        foreach (string attribute in new[] { "data-src", "data-original", "data-original-src", "data-lazy-src" }) {
            if (!node.Attributes.TryGetValue(attribute, out string? value)) continue;
            string resolved = HtmlUrlPolicyEvaluator.ResolveUrl(value, options.BaseUri, options.UrlPolicy);
            if (!string.IsNullOrWhiteSpace(resolved)) return resolved;
        }
        foreach (string attribute in new[] { "srcset", "data-srcset", "data-original-srcset", "data-lazy-srcset" }) {
            if (!node.Attributes.TryGetValue(attribute, out string? value)) continue;
            string resolved = HtmlImageSourceResolver.ResolveUrlFromSrcSet(value, options.BaseUri, options.UrlPolicy);
            if (!string.IsNullOrWhiteSpace(resolved)) return resolved;
        }
        return node.Attributes.TryGetValue("src", out string? source)
            ? HtmlUrlPolicyEvaluator.ResolveUrl(source, options.BaseUri, options.UrlPolicy)
            : string.Empty;
    }

    private static ReaderVisual MapHtmlVisual(
        HtmlLogicalNode node,
        ReaderLocation location,
        string? payloadHash,
        string? mediaType,
        string? sourceOverride = null) {
        node.Attributes.TryGetValue("src", out string? source);
        HtmlLogicalNode? mediaSource = null;
        if (string.IsNullOrWhiteSpace(source) && node.Kind == HtmlLogicalNodeKind.Media) {
            mediaSource = FindHtmlMediaSource(node);
            mediaSource?.Attributes.TryGetValue("src", out source);
        }
        if (!string.IsNullOrWhiteSpace(sourceOverride)) source = sourceOverride;
        node.Attributes.TryGetValue("alt", out string? altText);
        node.Attributes.TryGetValue("title", out string? title);
        if (mediaType == null && node.Attributes.TryGetValue("type", out string? declaredType)) mediaType = declaredType;
        if (mediaType == null && mediaSource != null && mediaSource.Attributes.TryGetValue("type", out string? sourceType)) {
            mediaType = sourceType;
        }
        string content = altText ?? title ?? GetHtmlNodeText(node);
        if (string.IsNullOrWhiteSpace(content)) content = source ?? node.Name;
        return new ReaderVisual {
            Kind = node.Kind == HtmlLogicalNodeKind.Image ? "image" : "media",
            Language = node.Name,
            Content = content,
            PayloadHash = payloadHash,
            SourceName = source,
            MimeType = mediaType,
            PlacementCount = 1,
            Location = new ReaderLocation {
                Path = location.Path,
                SourceBlockIndex = location.SourceBlockIndex,
                SourceBlockKind = location.SourceBlockKind,
                BlockAnchor = location.BlockAnchor
            }
        };
    }

    private static HtmlLogicalNode? FindHtmlMediaSource(HtmlLogicalNode node) {
        foreach (HtmlLogicalNode child in node.Children) {
            if (child.Kind == HtmlLogicalNodeKind.Media
                && string.Equals(child.Name, "source", StringComparison.OrdinalIgnoreCase)
                && child.Attributes.TryGetValue("src", out string? source)
                && !string.IsNullOrWhiteSpace(source)) {
                return child;
            }
            HtmlLogicalNode? descendant = FindHtmlMediaSource(child);
            if (descendant != null) return descendant;
        }
        return null;
    }

    private static IReadOnlyList<OfficeDocumentMetadataEntry> BuildHtmlMetadata(HtmlLogicalDocument logical, HtmlProjection projection) {
        return new[] {
            CountHtmlMetadata("html-block-count", "BlockCount", projection.Blocks.Count),
            CountHtmlMetadata("html-table-count", "TableCount", projection.Tables.Count),
            CountHtmlMetadata("html-link-count", "LinkCount", projection.Links.Count),
            CountHtmlMetadata("html-image-count", "ImageCount", projection.Assets.Count),
            CountHtmlMetadata("html-form-count", "FormFieldCount", projection.Forms.Count),
            CountHtmlMetadata("html-visual-count", "VisualCount", projection.Visuals.Count),
            CountHtmlMetadata("html-heading-count", "HeadingCount", logical.Count(HtmlLogicalNodeKind.Heading))
        };
    }

    private static OfficeDocumentMetadataEntry CountHtmlMetadata(string id, string name, int count) => new OfficeDocumentMetadataEntry {
        Id = id, Category = "html.summary", Name = name, Value = count.ToString(CultureInfo.InvariantCulture), ValueType = "count"
    };

    private static string? FindHtmlMetadata(HtmlLogicalNode node, string elementName, string? metaName) {
        if (node.Kind == HtmlLogicalNodeKind.Metadata && string.Equals(node.Name, elementName, StringComparison.OrdinalIgnoreCase)) {
            if (metaName == null) return node.Text;
            if (node.Attributes.TryGetValue("name", out string? name) && string.Equals(name, metaName, StringComparison.OrdinalIgnoreCase)
                && node.Attributes.TryGetValue("content", out string? content)) return content;
        }
        foreach (HtmlLogicalNode child in node.Children) {
            string? value = FindHtmlMetadata(child, elementName, metaName);
            if (!string.IsNullOrWhiteSpace(value)) return value;
        }
        return null;
    }

    private static IEnumerable<HtmlLogicalNode> Descendants(HtmlLogicalNode node, HtmlLogicalNodeKind kind) {
        foreach (HtmlLogicalNode child in node.Children) {
            if (child.Kind == kind) yield return child;
            foreach (HtmlLogicalNode descendant in Descendants(child, kind)) yield return descendant;
        }
    }

    private static IEnumerable<HtmlLogicalNode> GetHtmlTableRows(HtmlLogicalNode node) {
        foreach (HtmlLogicalNode child in node.Children) {
            if (child.Kind == HtmlLogicalNodeKind.Table) continue;
            if (child.Kind == HtmlLogicalNodeKind.TableRow) yield return child;
            foreach (HtmlLogicalNode row in GetHtmlTableRows(child)) yield return row;
        }
    }

    private static string GetHtmlNodeText(HtmlLogicalNode node) {
        if (!string.IsNullOrWhiteSpace(node.Text)) return node.Text;
        return string.Join(" ", Descendants(node, HtmlLogicalNodeKind.Text).Select(static child => child.Text).Where(static text => !string.IsNullOrWhiteSpace(text)));
    }

    private static bool IsHtmlBlock(HtmlLogicalNodeKind kind) => kind == HtmlLogicalNodeKind.Heading || kind == HtmlLogicalNodeKind.Paragraph
        || kind == HtmlLogicalNodeKind.ListItem || kind == HtmlLogicalNodeKind.Table || kind == HtmlLogicalNodeKind.Figure
        || kind == HtmlLogicalNodeKind.Image || kind == HtmlLogicalNodeKind.Media || kind == HtmlLogicalNodeKind.Form;

    private static string NormalizeHtmlBlockKind(HtmlLogicalNodeKind kind) => kind switch {
        HtmlLogicalNodeKind.ListItem => "list-item",
        _ => kind.ToString().ToLowerInvariant()
    };

    private static int? ParseHtmlHeadingLevel(string name) => name.Length == 2 && name[0] == 'h' && name[1] >= '1' && name[1] <= '6' ? name[1] - '0' : null;

    private static ReaderLocation BuildHtmlLocation(string? path, int? blockIndex, string kind, string anchor, int? tableIndex = null) => new ReaderLocation {
        Path = path, SourceBlockIndex = blockIndex, SourceBlockKind = kind, BlockAnchor = anchor, TableIndex = tableIndex
    };

    private static string ComputeHtmlHash(byte[] bytes) {
        using var sha = SHA256.Create();
        return string.Concat(sha.ComputeHash(bytes).Select(static value => value.ToString("x2", CultureInfo.InvariantCulture)));
    }

    private sealed class HtmlProjection {
        internal List<OfficeDocumentBlock> Blocks { get; } = new List<OfficeDocumentBlock>();
        internal List<ReaderTable> Tables { get; } = new List<ReaderTable>();
        internal List<OfficeDocumentLink> Links { get; } = new List<OfficeDocumentLink>();
        internal List<OfficeDocumentAsset> Assets { get; } = new List<OfficeDocumentAsset>();
        internal List<OfficeDocumentFormField> Forms { get; } = new List<OfficeDocumentFormField>();
        internal List<ReaderVisual> Visuals { get; } = new List<ReaderVisual>();
    }
}
