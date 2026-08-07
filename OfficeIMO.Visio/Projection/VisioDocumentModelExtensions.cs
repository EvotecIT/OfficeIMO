using OfficeIMO;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Threading;

namespace OfficeIMO.Visio;

/// <summary>Projects Visio documents into the dependency-free OfficeIMO document model.</summary>
public static class VisioDocumentModelExtensions {
    /// <summary>Projects an already loaded Visio document into the neutral document model.</summary>
    public static OfficeDocumentModel ToOfficeDocumentModel(
        this VisioDocument document,
        string? sourceName = null,
        VisioDocumentProjectionOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (sourceName != null && string.IsNullOrWhiteSpace(sourceName)) {
            throw new ArgumentException("Source name cannot be empty.", nameof(sourceName));
        }

        VisioDocumentProjectionOptions operation = (options ?? new VisioDocumentProjectionOptions()).Snapshot();
        cancellationToken.ThrowIfCancellationRequested();

        string logicalSourceName = sourceName ?? document.FilePath ?? "document.vsdx";
        var source = new OfficeDocumentModelSource {
            Path = logicalSourceName,
            SourceId = BuildSourceId(logicalSourceName),
            Title = document.Title,
            Author = document.Author
        };
        VisioInspectionSnapshot snapshot = document.CreateInspectionSnapshot();
        OfficeDocumentModelBlock[] blocks = BuildBlocks(snapshot, logicalSourceName).ToArray();
        OfficeDocumentModelTable[] tables = BuildTables(snapshot, logicalSourceName, operation.MaxTableRows).ToArray();
        VisioPage[] orderedPages = GetSnapshotOrderedPages(document, snapshot).ToArray();
        OfficeDocumentModelLink[] links = BuildLinks(orderedPages, logicalSourceName).ToArray();
        OfficeDocumentModelAsset[] assets = BuildAssets(orderedPages, logicalSourceName, operation, cancellationToken).ToArray();
        OfficeDocumentModelVisual[] visuals = BuildVisuals(snapshot, logicalSourceName).ToArray();
        OfficeDocumentModelPage[] pages = BuildPages(snapshot, logicalSourceName, blocks, tables, links, assets).ToArray();

        return new OfficeDocumentModel {
            Format = OfficeDocumentFormat.Visio,
            Source = source,
            CapabilitiesUsed = BuildCapabilities(operation),
            Markdown = pages.Length == 0
                ? null
                : string.Join(Environment.NewLine + Environment.NewLine, pages.Select(static page => page.Markdown)),
            Metadata = BuildMetadata(snapshot, tables, links, assets, visuals),
            Pages = pages,
            Blocks = blocks,
            Tables = tables,
            Assets = assets,
            Links = links,
            Forms = Array.Empty<OfficeDocumentModelFormField>(),
            Visuals = visuals,
            Diagnostics = Array.Empty<OfficeDocumentModelDiagnostic>()
        };
    }

    private static IReadOnlyList<string> BuildCapabilities(VisioDocumentProjectionOptions options) {
        var capabilities = new List<string> {
            "officeimo.visio.document-model",
            "officeimo.visio.inspection-snapshot",
            "officeimo.visio.topology-visual"
        };
        if (options.IncludeSvgPreviewAssets) capabilities.Add("officeimo.visio.svg-preview");
        if (options.IncludePngPreviewAssets) capabilities.Add("officeimo.visio.png-preview");
        return capabilities;
    }

    private static IEnumerable<OfficeDocumentModelBlock> BuildBlocks(
        VisioInspectionSnapshot snapshot,
        string sourceName) {
        for (int pageIndex = 0; pageIndex < snapshot.Pages.Count; pageIndex++) {
            VisioInspectionPageSnapshot page = snapshot.Pages[pageIndex];
            foreach (VisioInspectionShapeSnapshot shape in page.Shapes) {
                yield return new OfficeDocumentModelBlock {
                    Id = "visio-page-" + (pageIndex + 1).ToString("D4", CultureInfo.InvariantCulture) + "-shape-" + shape.Id,
                    Kind = ResolveShapeKind(shape),
                    Text = BuildShapeText(shape),
                    Location = BuildLocation(sourceName, pageIndex, "shape", "page-" + (pageIndex + 1).ToString(CultureInfo.InvariantCulture) + "-shape-" + shape.Id),
                    Region = new OfficeDocumentModelRegion {
                        X = InchesToPoints(shape.PinX - (shape.Width / 2D)),
                        Y = InchesToPoints(shape.PinY - (shape.Height / 2D)),
                        Width = InchesToPoints(shape.Width),
                        Height = InchesToPoints(shape.Height)
                    }
                };
            }

            foreach (VisioInspectionConnectorSnapshot connector in page.Connectors) {
                yield return new OfficeDocumentModelBlock {
                    Id = "visio-page-" + (pageIndex + 1).ToString("D4", CultureInfo.InvariantCulture) + "-connector-" + connector.Id,
                    Kind = "connector",
                    Text = BuildConnectorText(connector),
                    Location = BuildLocation(sourceName, pageIndex, "connector", "page-" + (pageIndex + 1).ToString(CultureInfo.InvariantCulture) + "-connector-" + connector.Id)
                };
            }
        }
    }

    private static IEnumerable<OfficeDocumentModelTable> BuildTables(
        VisioInspectionSnapshot snapshot,
        string sourceName,
        int maxTableRows) {
        for (int pageIndex = 0; pageIndex < snapshot.Pages.Count; pageIndex++) {
            VisioInspectionPageSnapshot page = snapshot.Pages[pageIndex];
            var rows = new List<IReadOnlyList<string>>();
            foreach (VisioInspectionShapeSnapshot shape in page.Shapes) {
                AddShapeDataRows(rows, "shape", shape.Id, shape.Text, shape.ShapeData);
            }
            foreach (VisioInspectionConnectorSnapshot connector in page.Connectors) {
                AddShapeDataRows(rows, "connector", connector.Id, connector.Label, connector.ShapeData);
            }
            if (rows.Count == 0) continue;

            int totalRowCount = rows.Count;
            IReadOnlyList<IReadOnlyList<string>> visibleRows = rows.Count > maxTableRows
                ? rows.Take(maxTableRows).ToArray()
                : rows;
            yield return new OfficeDocumentModelTable {
                Title = page.Name + " Shape Data",
                Kind = "visio-shape-data",
                Location = BuildLocation(sourceName, pageIndex, "shape-data", "page-" + (pageIndex + 1).ToString(CultureInfo.InvariantCulture) + "-shape-data"),
                Columns = new[] { "OwnerType", "OwnerId", "OwnerText", "Name", "Label", "Value", "Type", "Prompt" },
                Rows = visibleRows,
                TotalRowCount = totalRowCount,
                Truncated = totalRowCount > visibleRows.Count
            };
        }
    }

    private static IEnumerable<OfficeDocumentModelPage> BuildPages(
        VisioInspectionSnapshot snapshot,
        string sourceName,
        IReadOnlyList<OfficeDocumentModelBlock> blocks,
        IReadOnlyList<OfficeDocumentModelTable> tables,
        IReadOnlyList<OfficeDocumentModelLink> links,
        IReadOnlyList<OfficeDocumentModelAsset> assets) {
        for (int pageIndex = 0; pageIndex < snapshot.Pages.Count; pageIndex++) {
            VisioInspectionPageSnapshot page = snapshot.Pages[pageIndex];
            int pageNumber = pageIndex + 1;
            yield return new OfficeDocumentModelPage {
                Number = pageNumber,
                Name = page.Name,
                Text = BuildPageText(page),
                Markdown = BuildPageMarkdown(snapshot, page),
                Width = InchesToPoints(page.Width),
                Height = InchesToPoints(page.Height),
                Location = BuildLocation(sourceName, pageIndex, "page", "page-" + pageNumber.ToString(CultureInfo.InvariantCulture)),
                Blocks = blocks.Where(block => block.Location.Page == pageNumber).ToArray(),
                Tables = tables.Where(table => table.Location?.Page == pageNumber).ToArray(),
                Assets = assets.Where(asset => asset.Location.Page == pageNumber).ToArray(),
                Links = links.Where(link => link.Location.Page == pageNumber).ToArray(),
                Forms = Array.Empty<OfficeDocumentModelFormField>()
            };
        }
    }

    private static IEnumerable<VisioPage> GetSnapshotOrderedPages(
        VisioDocument document,
        VisioInspectionSnapshot snapshot) {
        foreach (VisioInspectionPageSnapshot snapshotPage in snapshot.Pages) {
            VisioPage? page = document.Pages.FirstOrDefault(candidate => candidate.Id == snapshotPage.Id);
            if (page != null) yield return page;
        }
    }

    private static IEnumerable<OfficeDocumentModelLink> BuildLinks(
        IReadOnlyList<VisioPage> pages,
        string sourceName) {
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++) {
            VisioPage page = pages[pageIndex];
            foreach (VisioShape shape in page.AllShapes()) {
                for (int linkIndex = 0; linkIndex < shape.Hyperlinks.Count; linkIndex++) {
                    VisioHyperlink link = shape.Hyperlinks[linkIndex];
                    yield return BuildLink(
                        "visio-page-" + (pageIndex + 1).ToString("D4", CultureInfo.InvariantCulture) + "-shape-" + shape.Id + "-link-" + linkIndex.ToString("D4", CultureInfo.InvariantCulture),
                        link,
                        sourceName,
                        pageIndex,
                        "shape",
                        shape.Id,
                        new OfficeDocumentModelRegion {
                            X = InchesToPoints(shape.PinX - (shape.Width / 2D)),
                            Y = InchesToPoints(shape.PinY - (shape.Height / 2D)),
                            Width = InchesToPoints(shape.Width),
                            Height = InchesToPoints(shape.Height)
                        });
                }
            }
            foreach (VisioConnector connector in page.Connectors) {
                for (int linkIndex = 0; linkIndex < connector.Hyperlinks.Count; linkIndex++) {
                    yield return BuildLink(
                        "visio-page-" + (pageIndex + 1).ToString("D4", CultureInfo.InvariantCulture) + "-connector-" + connector.Id + "-link-" + linkIndex.ToString("D4", CultureInfo.InvariantCulture),
                        connector.Hyperlinks[linkIndex],
                        sourceName,
                        pageIndex,
                        "connector",
                        connector.Id,
                        null);
                }
            }
        }
    }

    private static OfficeDocumentModelLink BuildLink(
        string id,
        VisioHyperlink link,
        string sourceName,
        int pageIndex,
        string ownerKind,
        string ownerId,
        OfficeDocumentModelRegion? region) => new OfficeDocumentModelLink {
            Id = id,
            Kind = string.IsNullOrWhiteSpace(link.Address) ? "internal" : "uri",
            Uri = link.Address,
            DestinationName = link.SubAddress,
            Text = link.Description,
            Location = BuildLocation(sourceName, pageIndex, ownerKind + "-hyperlink", "page-" + (pageIndex + 1).ToString(CultureInfo.InvariantCulture) + "-" + ownerKind + "-" + ownerId + "-link"),
            Region = region
        };

    private static IEnumerable<OfficeDocumentModelAsset> BuildAssets(
        IReadOnlyList<VisioPage> pages,
        string sourceName,
        VisioDocumentProjectionOptions options,
        CancellationToken cancellationToken) {
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            if (options.IncludeSvgPreviewAssets) {
                yield return BuildPreviewAsset(sourceName, pageIndex, "preview-svg", "image/svg+xml", ".svg", Encoding.UTF8.GetBytes(pages[pageIndex].ToSvg(options.SvgOptions)));
            }
            if (options.IncludePngPreviewAssets) {
                yield return BuildPreviewAsset(
                    sourceName,
                    pageIndex,
                    "preview-png",
                    "image/png",
                    ".png",
                    VisioPngExportExtensions.ToPng(pages[pageIndex], options.PngOptions, cancellationToken));
            }
        }
    }

    private static OfficeDocumentModelAsset BuildPreviewAsset(
        string sourceName,
        int pageIndex,
        string kind,
        string mediaType,
        string extension,
        byte[] payload) {
        string id = "visio-page-" + (pageIndex + 1).ToString("D4", CultureInfo.InvariantCulture) + "-" + kind;
        return new OfficeDocumentModelAsset {
            Id = id,
            Kind = kind,
            MediaType = mediaType,
            Extension = extension,
            FileName = id + extension,
            LengthBytes = payload.LongLength,
            PayloadHash = ComputeSha256Hex(payload),
            PayloadBytes = payload,
            Location = BuildLocation(sourceName, pageIndex, kind, "page-" + (pageIndex + 1).ToString(CultureInfo.InvariantCulture) + "-" + kind)
        };
    }

    private static IEnumerable<OfficeDocumentModelVisual> BuildVisuals(
        VisioInspectionSnapshot snapshot,
        string sourceName) {
        for (int pageIndex = 0; pageIndex < snapshot.Pages.Count; pageIndex++) {
            VisioInspectionPageSnapshot page = snapshot.Pages[pageIndex];
            string content = BuildTopologyPayload(page);
            yield return new OfficeDocumentModelVisual {
                Kind = "network",
                Language = "officeimo-visio-topology-v1",
                Content = content,
                PayloadHash = ComputeSha256Hex(Encoding.UTF8.GetBytes(content)),
                SourceName = page.Name,
                Width = InchesToPoints(page.Width),
                Height = InchesToPoints(page.Height),
                PlacedWidth = InchesToPoints(page.Width),
                PlacedHeight = InchesToPoints(page.Height),
                PlacementCount = 1,
                HasGeometry = true,
                IsAxisAligned = true,
                Location = BuildLocation(sourceName, pageIndex, "diagram", "page-" + (pageIndex + 1).ToString(CultureInfo.InvariantCulture) + "-topology")
            };
        }
    }

    private static string BuildTopologyPayload(VisioInspectionPageSnapshot page) {
        var builder = new StringBuilder();
        builder.Append("page\t").Append(EscapeTopologyValue(page.Id.ToString(CultureInfo.InvariantCulture))).Append('\t')
            .Append(EscapeTopologyValue(page.Name)).Append('\t')
            .Append(page.Width.ToString(CultureInfo.InvariantCulture)).Append('\t')
            .Append(page.Height.ToString(CultureInfo.InvariantCulture)).AppendLine();
        foreach (VisioInspectionShapeSnapshot shape in page.Shapes) {
            builder.Append("node\t").Append(EscapeTopologyValue(shape.Id)).Append('\t')
                .Append(EscapeTopologyValue(shape.Text)).Append('\t')
                .Append(EscapeTopologyValue(shape.MasterNameU)).Append('\t')
                .Append(shape.PinX.ToString(CultureInfo.InvariantCulture)).Append('\t')
                .Append(shape.PinY.ToString(CultureInfo.InvariantCulture)).AppendLine();
        }
        foreach (VisioInspectionConnectorSnapshot connector in page.Connectors) {
            builder.Append("edge\t").Append(EscapeTopologyValue(connector.Id)).Append('\t')
                .Append(EscapeTopologyValue(connector.FromId)).Append('\t')
                .Append(EscapeTopologyValue(connector.ToId)).Append('\t')
                .Append(EscapeTopologyValue(connector.Label)).AppendLine();
        }
        return builder.ToString().TrimEnd();
    }

    private static string EscapeTopologyValue(string? value) =>
        (value ?? string.Empty).Replace("\\", "\\\\").Replace("\t", "\\t").Replace("\r", "\\r").Replace("\n", "\\n");

    private static IReadOnlyList<OfficeDocumentModelMetadataEntry> BuildMetadata(
        VisioInspectionSnapshot snapshot,
        IReadOnlyList<OfficeDocumentModelTable> tables,
        IReadOnlyList<OfficeDocumentModelLink> links,
        IReadOnlyList<OfficeDocumentModelAsset> assets,
        IReadOnlyList<OfficeDocumentModelVisual> visuals) {
        var metadata = new List<OfficeDocumentModelMetadataEntry> {
            BuildCountMetadata("visio-page-count", "PageCount", snapshot.Pages.Count),
            BuildCountMetadata("visio-shape-count", "ShapeCount", snapshot.ShapeCount),
            BuildCountMetadata("visio-connector-count", "ConnectorCount", snapshot.ConnectorCount),
            BuildCountMetadata("visio-table-count", "TableCount", tables.Count),
            BuildCountMetadata("visio-link-count", "LinkCount", links.Count),
            BuildCountMetadata("visio-asset-count", "AssetCount", assets.Count),
            BuildCountMetadata("visio-visual-count", "VisualCount", visuals.Count)
        };
        if (!string.IsNullOrWhiteSpace(snapshot.ThemeType)) {
            metadata.Add(new OfficeDocumentModelMetadataEntry {
                Id = "visio-theme",
                Category = "visio.document",
                Name = "Theme",
                Value = snapshot.ThemeType,
                ValueType = "string"
            });
        }
        return metadata;
    }

    private static OfficeDocumentModelMetadataEntry BuildCountMetadata(string id, string name, int count) =>
        new OfficeDocumentModelMetadataEntry {
            Id = id,
            Category = "visio.summary",
            Name = name,
            Value = count.ToString(CultureInfo.InvariantCulture),
            ValueType = "count"
        };

    private static string BuildPageMarkdown(
        VisioInspectionSnapshot snapshot,
        VisioInspectionPageSnapshot page) {
        var builder = new StringBuilder();
        builder.Append("# ").AppendLine(string.IsNullOrWhiteSpace(page.Name) ? "Page " + page.Id.ToString(CultureInfo.InvariantCulture) : page.Name)
            .AppendLine()
            .Append("Document: ").AppendLine(string.IsNullOrWhiteSpace(snapshot.Title) ? "Untitled Visio document" : snapshot.Title)
            .Append("- Shapes: ").AppendLine(page.Shapes.Count.ToString(CultureInfo.InvariantCulture))
            .Append("- Connectors: ").AppendLine(page.Connectors.Count.ToString(CultureInfo.InvariantCulture));
        if (page.Layers.Count > 0) builder.Append("- Layers: ").AppendLine(string.Join(", ", page.Layers));
        if (page.Shapes.Count > 0) {
            builder.AppendLine().AppendLine("## Shapes");
            foreach (VisioInspectionShapeSnapshot shape in page.Shapes) {
                builder.Append("- ").Append(string.IsNullOrWhiteSpace(shape.Text) ? shape.Id : shape.Text)
                    .Append(" (`").Append(shape.Id).Append('`');
                if (!string.IsNullOrWhiteSpace(shape.MasterNameU)) builder.Append(", master `").Append(shape.MasterNameU).Append('`');
                builder.Append(')');
                if (shape.ShapeData.Count > 0) builder.Append(": ").Append(string.Join("; ", shape.ShapeData.Select(FormatShapeData)));
                builder.AppendLine();
            }
        }
        if (page.Connectors.Count > 0) {
            builder.AppendLine().AppendLine("## Connectors");
            foreach (VisioInspectionConnectorSnapshot connector in page.Connectors) {
                builder.Append("- ").Append(connector.FromId).Append(" -> ").Append(connector.ToId);
                if (!string.IsNullOrWhiteSpace(connector.Label)) builder.Append(": ").Append(connector.Label);
                if (connector.ShapeData.Count > 0) builder.Append(" (").Append(string.Join("; ", connector.ShapeData.Select(FormatShapeData))).Append(')');
                builder.AppendLine();
            }
        }
        return builder.ToString().TrimEnd();
    }

    private static string BuildPageText(VisioInspectionPageSnapshot page) {
        var parts = new List<string> {
            "Visio page " + page.Name + ": " + page.Shapes.Count.ToString(CultureInfo.InvariantCulture) + " shape(s), " + page.Connectors.Count.ToString(CultureInfo.InvariantCulture) + " connector(s)."
        };
        parts.AddRange(page.Shapes.Select(shape => string.IsNullOrWhiteSpace(shape.Text) ? shape.Id : shape.Text!));
        parts.AddRange(page.Connectors.Select(connector => string.IsNullOrWhiteSpace(connector.Label) ? connector.FromId + " -> " + connector.ToId : connector.Label!));
        return string.Join(Environment.NewLine, parts);
    }

    private static void AddShapeDataRows(
        List<IReadOnlyList<string>> rows,
        string ownerType,
        string ownerId,
        string? ownerText,
        IReadOnlyList<VisioInspectionShapeDataSnapshot> shapeDataRows) {
        foreach (VisioInspectionShapeDataSnapshot row in shapeDataRows) {
            rows.Add(new[] {
                ownerType,
                ownerId,
                ownerText ?? string.Empty,
                row.Name,
                row.Label ?? string.Empty,
                row.Value ?? string.Empty,
                row.Type ?? string.Empty,
                row.Prompt ?? string.Empty
            });
        }
    }

    private static OfficeDocumentModelLocation BuildLocation(
        string sourceName,
        int pageIndex,
        string sourceBlockKind,
        string blockAnchor) => new OfficeDocumentModelLocation {
            Path = sourceName,
            Page = pageIndex + 1,
            SourceBlockIndex = pageIndex,
            SourceBlockKind = sourceBlockKind,
            BlockAnchor = blockAnchor
        };

    private static string ResolveShapeKind(VisioInspectionShapeSnapshot shape) {
        if (shape.IsContainer) return "container";
        if (shape.IsCallout) return "callout";
        if (shape.IsBackgroundSurface) return "background";
        if (shape.IsDiagramAdornment) return "adornment";
        if (string.Equals(shape.Type, "Group", StringComparison.OrdinalIgnoreCase)) return "group";
        return "shape";
    }

    private static string BuildShapeText(VisioInspectionShapeSnapshot shape) {
        var builder = new StringBuilder(string.IsNullOrWhiteSpace(shape.Text) ? shape.Id : shape.Text);
        if (!string.IsNullOrWhiteSpace(shape.MasterNameU)) builder.Append(" [").Append(shape.MasterNameU).Append(']');
        if (shape.ShapeData.Count > 0) builder.Append(' ').Append(string.Join("; ", shape.ShapeData.Select(FormatShapeData)));
        return builder.ToString();
    }

    private static string BuildConnectorText(VisioInspectionConnectorSnapshot connector) {
        var builder = new StringBuilder().Append(connector.FromId).Append(" -> ").Append(connector.ToId);
        if (!string.IsNullOrWhiteSpace(connector.Label)) builder.Append(": ").Append(connector.Label);
        if (connector.ShapeData.Count > 0) builder.Append(' ').Append(string.Join("; ", connector.ShapeData.Select(FormatShapeData)));
        return builder.ToString();
    }

    private static string FormatShapeData(VisioInspectionShapeDataSnapshot row) =>
        (string.IsNullOrWhiteSpace(row.Label) ? row.Name : row.Label!) + "=" + (row.Value ?? string.Empty);

    private static string BuildSourceId(string sourceName) =>
        "visio:" + ComputeSha256Hex(Encoding.UTF8.GetBytes(sourceName.Replace('\\', '/').ToUpperInvariant()));

    private static string ComputeSha256Hex(byte[] value) {
        using SHA256 algorithm = SHA256.Create();
        byte[] hash = algorithm.ComputeHash(value);
        var builder = new StringBuilder(hash.Length * 2);
        foreach (byte item in hash) builder.Append(item.ToString("x2", CultureInfo.InvariantCulture));
        return builder.ToString();
    }

    private static double InchesToPoints(double value) => value * 72D;
}
