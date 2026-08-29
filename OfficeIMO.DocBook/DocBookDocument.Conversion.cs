using OfficeIMO;
using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Xml.Linq;

namespace OfficeIMO.DocBook;

public sealed partial class DocBookDocument {
    /// <summary>Converts the typed common structure to the shared recursive document model.</summary>
    public DocBookConversionResult<OfficeDocumentModel> ToOfficeDocumentModel(
        string? sourcePath = null,
        DocBookConversionOptions? options = null,
        CancellationToken cancellationToken = default) {
        options ??= new DocBookConversionOptions();
        options.Validate();
        cancellationToken.ThrowIfCancellationRequested();
        var diagnostics = new DocBookDiagnosticCollector(options.MaxDetailedDiagnosticsPerCode);
        var textBudget = new DocBookTextProjectionBudget(options.MaxTotalTextCharacters, diagnostics, Namespace, cancellationToken);
        var blocks = new List<OfficeDocumentModelBlock>();
        var tables = new List<OfficeDocumentModelTable>();
        var assets = new List<OfficeDocumentModelAsset>();
        var links = new List<OfficeDocumentModelLink>();
        var tableIndexes = new Dictionary<XElement, int>();
        int remainingTableCells = options.MaxTableCells;
        int index = 0;
        bool hasPreservedNonElementNodes = false;
        foreach (XNode node in _xml.DescendantNodes()) {
            cancellationToken.ThrowIfCancellationRequested();
            if (node is XComment || node is XProcessingInstruction) {
                hasPreservedNonElementNodes = true;
                break;
            }
        }
        if (hasPreservedNonElementNodes) {
            diagnostics.Add(new DocBookDiagnostic("DB105", DocBookDiagnosticSeverity.Warning,
                "Comments and processing instructions remain native but are not represented by the shared document model."));
        }
        bool hasRootExtensionAttributes = false;
        foreach (XAttribute attribute in RootElement.Attributes()) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!attribute.IsNamespaceDeclaration && attribute.Name != XName.Get("version")) {
                hasRootExtensionAttributes = true;
                break;
            }
        }
        if (hasRootExtensionAttributes) {
            diagnostics.Add(new DocBookDiagnostic("DB106", DocBookDiagnosticSeverity.Warning,
                "Root extension attributes remain native but are not represented by the shared document model.", "/" + RootElement.Name.LocalName));
        }
        bool hasSignificantRootText = false;
        foreach (XNode node in RootElement.Nodes()) {
            cancellationToken.ThrowIfCancellationRequested();
            if (node is XText text && !string.IsNullOrWhiteSpace(text.Value)) {
                hasSignificantRootText = true;
                break;
            }
        }
        if (hasSignificantRootText) {
            diagnostics.Add(new DocBookDiagnostic("DB110", DocBookDiagnosticSeverity.Warning,
                "Significant root text remains native but is not represented by the shared document model.", "/" + RootElement.Name.LocalName));
        }
        if (Profile == DocBookProfile.DocBook52 && (string?)RootElement.Attribute("version") != "5.2") {
            diagnostics.Add(new DocBookDiagnostic("DB111", DocBookDiagnosticSeverity.Warning,
                $"The source declares DocBook '{(string?)RootElement.Attribute("version") ?? "unspecified"}'; shared-model reconstruction normalizes it to the exact 5.2 writer profile.",
                "/" + RootElement.Name.LocalName + "/@version"));
        }
        XDocumentType? documentType = _xml.DocumentType;
        bool documentTypeIsNormalized = Profile == DocBookProfile.DocBook45
            ? documentType == null || documentType.Name != RootElement.Name.LocalName ||
              documentType.PublicId != DocBookSchemaProfiles.DocBook45.DtdPublicId ||
              documentType.SystemId != DocBookSchemaProfiles.DocBook45.DtdSystemId ||
              !string.IsNullOrWhiteSpace(documentType.InternalSubset)
            : documentType != null;
        if (documentTypeIsNormalized) {
            diagnostics.Add(new DocBookDiagnostic("DB107", DocBookDiagnosticSeverity.Warning,
                "The source document type differs from the exact writer profile; shared-model reconstruction normalizes it."));
        }

        OfficeDocumentModelNode Convert(XElement element, int level, string parentPath) {
            cancellationToken.ThrowIfCancellationRequested();
            DocBookNodeKind kind = DocBookNames.GetKind(element.Name, Namespace);
            string normalizedKind = kind == DocBookNodeKind.Unknown
                ? "extension:" + element.Name
                : kind == DocBookNodeKind.Table && element.Name.LocalName == "informaltable"
                    ? "informal-table"
                    : ToModelKind(kind);
            string text = textBudget.GetPrimaryText(element, kind, Namespace, parentPath);
            string path = kind == DocBookNodeKind.Section
                ? OfficeDocumentHeadingPath.Append(parentPath, text, " / ") : parentPath;
            var attributes = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (XAttribute attribute in element.Attributes()) {
                cancellationToken.ThrowIfCancellationRequested();
                attributes.Add(attribute.Name.ToString(), attribute.Value);
            }
            int nodeIndex = index++;
            int? tableIndex = kind == DocBookNodeKind.Table && tableIndexes.TryGetValue(element, out int projectedTableIndex)
                ? projectedTableIndex : (int?)null;
            if (kind == DocBookNodeKind.Unknown) {
                diagnostics.Add(new DocBookDiagnostic("DB100", DocBookDiagnosticSeverity.Info,
                    $"Extension element '{element.Name}' was represented as a generic shared-model node.", path));
            }
            if (element.Name.Namespace == Namespace &&
                (element.Name.LocalName == "simpara" || element.Name.LocalName == "sect1" || element.Name.LocalName == "sect2" ||
                 element.Name.LocalName == "sect3" || element.Name.LocalName == "sect4" || element.Name.LocalName == "sect5" ||
                 (Profile == DocBookProfile.DocBook52 &&
                    (element.Name.LocalName == "ulink" || element.Name.LocalName == "articleinfo" || element.Name.LocalName == "bookinfo" ||
                     element.Name.LocalName == "sectioninfo")) ||
                 (Profile == DocBookProfile.DocBook45 && element.Name.LocalName == "info"))) {
                diagnostics.Add(new DocBookDiagnostic("DB115", DocBookDiagnosticSeverity.Warning,
                    $"Native element '{element.Name.LocalName}' is canonicalized by shared-model reconstruction.", path));
            }
            if (kind == DocBookNodeKind.Link || kind == DocBookNodeKind.CrossReference) {
                links.Add(new OfficeDocumentModelLink {
                    Id = "docbook-link-" + nodeIndex,
                    Kind = kind == DocBookNodeKind.CrossReference ? "cross-reference" : "link",
                    Uri = (string?)element.Attribute("url") ?? (string?)element.Attribute(XName.Get("href", "http://www.w3.org/1999/xlink")),
                    DestinationName = (string?)element.Attribute("linkend"),
                    Text = text,
                    Location = new OfficeDocumentModelLocation { Path = sourcePath, HeadingPath = path }
                });
            }
            if (kind == DocBookNodeKind.ImageData) {
                string? fileReference = (string?)element.Attribute("fileref");
                if (!string.IsNullOrWhiteSpace(fileReference)) {
                    XElement? mediaObject = element.Ancestors().FirstOrDefault(ancestor =>
                        DocBookNames.GetKind(ancestor.Name, Namespace) == DocBookNodeKind.MediaObject);
                    XElement? captionElement = mediaObject?.Elements().FirstOrDefault(child =>
                        DocBookNames.GetKind(child.Name, Namespace) == DocBookNodeKind.Caption);
                    string? caption = captionElement == null ? null : textBudget.GetElementValue(captionElement, path);
                    XElement? alternateTextElement = mediaObject?.Elements(Namespace + "textobject")
                        .SelectMany(textObject => textObject.Descendants(Namespace + "phrase").Where(phrase =>
                            ReferenceEquals(phrase.Ancestors(Namespace + "textobject").FirstOrDefault(), textObject)))
                        .FirstOrDefault();
                    string? alternateText = alternateTextElement == null ? null : textBudget.GetElementValue(alternateTextElement, path);
                    string? fileName = GetReferenceFileName(fileReference!);
                    string? extension = GetReferenceExtension(fileName);
                    string mediaType = OfficeImageInfo.GetMimeTypeFromExtension(extension);
                    assets.Add(new OfficeDocumentModelAsset {
                        Id = "docbook-image-" + nodeIndex,
                        Kind = "image",
                        MediaType = mediaType == "application/octet-stream" ? null : mediaType,
                        Extension = extension,
                        FileName = fileName,
                        AltText = string.IsNullOrWhiteSpace(alternateText) ? caption : alternateText,
                        Title = caption,
                        SourceObjectId = fileReference,
                        Location = new OfficeDocumentModelLocation { Path = sourcePath, HeadingPath = path, SourceBlockKind = "image" }
                    });
                }
            }
            if (kind == DocBookNodeKind.Section || kind == DocBookNodeKind.Paragraph || kind == DocBookNodeKind.ProgramListing ||
                kind == DocBookNodeKind.Screen || kind == DocBookNodeKind.ListItem || kind == DocBookNodeKind.Note ||
                kind == DocBookNodeKind.Warning || kind == DocBookNodeKind.Tip || kind == DocBookNodeKind.Important || kind == DocBookNodeKind.Caution) {
                blocks.Add(new OfficeDocumentModelBlock {
                    Id = "docbook-" + nodeIndex,
                    Kind = normalizedKind,
                    Text = text,
                    Level = level,
                    Location = new OfficeDocumentModelLocation { Path = sourcePath, BlockIndex = blocks.Count, HeadingPath = path }
                });
            }
            IReadOnlyList<OfficeDocumentModelNode> children = BuildChildren(element, kind, level, path);
            return new OfficeDocumentModelNode {
                Id = "docbook-" + nodeIndex,
                Kind = normalizedKind,
                Text = text,
                Level = level,
                Attributes = attributes,
                Children = children,
                Location = new OfficeDocumentModelLocation { Path = sourcePath, HeadingPath = path, TableIndex = tableIndex }
            };
        }

        IReadOnlyList<OfficeDocumentModelNode> BuildChildren(XElement element, DocBookNodeKind kind, int level, string path) {
            bool mixedContent = kind == DocBookNodeKind.Unknown || kind == DocBookNodeKind.Paragraph || kind == DocBookNodeKind.Title || kind == DocBookNodeKind.Subtitle ||
                kind == DocBookNodeKind.Link || kind == DocBookNodeKind.Entry || kind == DocBookNodeKind.Caption || kind == DocBookNodeKind.Author ||
                kind == DocBookNodeKind.ProgramListing || kind == DocBookNodeKind.Screen;
            var children = new List<OfficeDocumentModelNode>();
            int childLevel = kind == DocBookNodeKind.Info ? level : level + 1;
            foreach (XNode node in element.Nodes()) {
                cancellationToken.ThrowIfCancellationRequested();
                if (node is XElement child) {
                    children.Add(Convert(child, childLevel, path));
                } else if (node is XText textNode && textNode.Value.Length > 0 && (mixedContent || !string.IsNullOrWhiteSpace(textNode.Value))) {
                    children.Add(new OfficeDocumentModelNode {
                        Id = "docbook-" + index++,
                        Kind = "text",
                        Text = textBudget.GetTextValue(textNode, path),
                        Level = childLevel,
                        Location = new OfficeDocumentModelLocation { Path = sourcePath, HeadingPath = path }
                    });
                }
            }
            return children;
        }

        foreach (XElement tableElement in RootElement.Descendants()) {
            cancellationToken.ThrowIfCancellationRequested();
            if (DocBookNames.GetKind(tableElement.Name, Namespace) != DocBookNodeKind.Table) continue;
            int tableCellBudget = Math.Min(options.MaxTableCells, remainingTableCells);
            int discoveryCapacity = options.MaxTableRows > int.MaxValue / 2
                ? int.MaxValue : options.MaxTableRows * 2;
            var rowElements = new List<XElement>(Math.Min(discoveryCapacity, 4_096));
            bool rowDiscoveryTruncated = false;
            bool footerRowsEncountered = false;
            bool nestedEntryTableEncountered = false;
            int totalBodyRows = 0;
            int headerRowsRetained = 0;
            int bodyRowsRetained = 0;
            foreach (XElement element in tableElement.Descendants()) {
                cancellationToken.ThrowIfCancellationRequested();
                if (DocBookNames.GetKind(element.Name, Namespace) != DocBookNodeKind.Row ||
                    !ReferenceEquals(element.Ancestors().FirstOrDefault(ancestor =>
                        DocBookNames.GetKind(ancestor.Name, Namespace) == DocBookNodeKind.Table), tableElement)) continue;
                if (element.Ancestors().TakeWhile(ancestor => !ReferenceEquals(ancestor, tableElement))
                    .Any(ancestor => ancestor.Name == Namespace + "entrytbl")) {
                    nestedEntryTableEncountered = true;
                    continue;
                }
                bool isHeader = element.Ancestors().TakeWhile(ancestor => !ReferenceEquals(ancestor, tableElement)).Any(ancestor =>
                    DocBookNames.GetKind(ancestor.Name, Namespace) == DocBookNodeKind.TableHead);
                if (element.Ancestors().TakeWhile(ancestor => !ReferenceEquals(ancestor, tableElement)).Any(ancestor =>
                        ancestor.Name == Namespace + "tfoot")) footerRowsEncountered = true;
                if (!isHeader) totalBodyRows++;
                if (isHeader ? headerRowsRetained >= options.MaxTableRows : bodyRowsRetained >= options.MaxTableRows) {
                    rowDiscoveryTruncated = true;
                    continue;
                }
                rowElements.Add(element);
                if (isHeader) headerRowsRetained++;
                else bodyRowsRetained++;
            }
            var projectedRows = new List<KeyValuePair<bool, IReadOnlyList<string>>>();
            var activeRowSpans = new Dictionary<XElement, Dictionary<int, int>>();
            var groupLayouts = new Dictionary<XElement, CalsProjectionLayout>();
            int columnCount = 0;
            bool flattenedCalsLayout = tableElement.Elements(Namespace + "tgroup").Skip(1).Any();
            bool flattenedFooterRows = footerRowsEncountered;
            bool projectionTruncated = rowDiscoveryTruncated;

            foreach (XElement rowElement in rowElements) {
                cancellationToken.ThrowIfCancellationRequested();
                int maxWidthForAdditionalRow = tableCellBudget / (projectedRows.Count + 2);
                if (maxWidthForAdditionalRow < Math.Max(1, columnCount)) {
                    projectionTruncated = true;
                    break;
                }
                int rowColumnLimit = Math.Min(options.MaxTableColumns, maxWidthForAdditionalRow);
                bool isHeader = rowElement.Ancestors().TakeWhile(ancestor => !ReferenceEquals(ancestor, tableElement)).Any(ancestor =>
                    DocBookNames.GetKind(ancestor.Name, Namespace) == DocBookNodeKind.TableHead);
                XElement? tableGroup = rowElement.Ancestors().FirstOrDefault(ancestor =>
                    DocBookNames.GetKind(ancestor.Name, Namespace) == DocBookNodeKind.TableGroup);
                XElement? rowGroup = rowElement.Ancestors().FirstOrDefault(ancestor =>
                    DocBookNames.GetKind(ancestor.Name, Namespace) == DocBookNodeKind.TableHead ||
                    DocBookNames.GetKind(ancestor.Name, Namespace) == DocBookNodeKind.TableBody ||
                    ancestor.Name == Namespace + "tfoot");
                XElement spanOwner = rowGroup ?? tableGroup ?? tableElement;
                CalsProjectionLayout layout;
                if (tableGroup == null) {
                    layout = CalsProjectionLayout.Empty;
                } else if (!groupLayouts.TryGetValue(tableGroup, out layout!)) {
                    int declaredColumns = 0;
                    if (int.TryParse((string?)tableGroup.Attribute("cols"), out int parsedColumns) && parsedColumns > 0) {
                        declaredColumns = Math.Min(parsedColumns, options.MaxTableColumns);
                        if (parsedColumns > options.MaxTableColumns) projectionTruncated = true;
                    }
                    var namedColumns = new Dictionary<string, int>(StringComparer.Ordinal);
                    var namedSpans = new Dictionary<string, KeyValuePair<int, int>>(StringComparer.Ordinal);
                    int nextColumn = 0;
                    foreach (XElement columnSpec in tableGroup.Elements(Namespace + "colspec")) {
                        cancellationToken.ThrowIfCancellationRequested();
                        int column = nextColumn;
                        if (int.TryParse((string?)columnSpec.Attribute("colnum"), out int columnNumber) && columnNumber > 0) {
                            if (columnNumber > options.MaxTableColumns) {
                                column = options.MaxTableColumns - 1;
                                projectionTruncated = true;
                            } else {
                                column = columnNumber - 1;
                            }
                        }
                        string? columnName = (string?)columnSpec.Attribute("colname");
                        if (!string.IsNullOrEmpty(columnName)) namedColumns[columnName!] = column;
                        nextColumn = Math.Max(nextColumn, column + 1);
                    }
                    foreach (XElement spanSpec in tableGroup.Elements(Namespace + "spanspec")) {
                        cancellationToken.ThrowIfCancellationRequested();
                        string? spanName = (string?)spanSpec.Attribute("spanname");
                        string? startName = (string?)spanSpec.Attribute("namest");
                        string? endName = (string?)spanSpec.Attribute("nameend");
                        if (!string.IsNullOrEmpty(spanName) && !string.IsNullOrEmpty(startName) && !string.IsNullOrEmpty(endName) &&
                            namedColumns.TryGetValue(startName!, out int spanStart) && namedColumns.TryGetValue(endName!, out int spanEnd)) {
                            namedSpans[spanName!] = new KeyValuePair<int, int>(spanStart, spanEnd);
                        }
                    }
                    layout = new CalsProjectionLayout(declaredColumns, nextColumn, namedColumns, namedSpans);
                    groupLayouts.Add(tableGroup, layout);
                }

                if (!activeRowSpans.TryGetValue(spanOwner, out Dictionary<int, int>? activeSpans)) {
                    activeSpans = new Dictionary<int, int>();
                    activeRowSpans.Add(spanOwner, activeSpans);
                }
                int initialWidth = Math.Max(layout.DeclaredColumns, layout.NextColumn);
                if (initialWidth > rowColumnLimit) projectionTruncated = true;
                initialWidth = Math.Min(initialWidth, rowColumnLimit);
                var cells = Enumerable.Repeat<string?>(null, initialWidth).ToList();
                int nextEntryColumn = 0;
                foreach (KeyValuePair<int, int> activeSpan in activeSpans.ToArray()) {
                    if (activeSpan.Key >= rowColumnLimit) {
                        activeSpans.Remove(activeSpan.Key);
                        projectionTruncated = true;
                        continue;
                    }
                    while (cells.Count <= activeSpan.Key) cells.Add(null);
                    cells[activeSpan.Key] = string.Empty;
                    if (activeSpan.Value <= 1) activeSpans.Remove(activeSpan.Key);
                    else activeSpans[activeSpan.Key] = activeSpan.Value - 1;
                }

                foreach (XElement entry in rowElement.Elements().Where(element =>
                             DocBookNames.GetKind(element.Name, Namespace) == DocBookNodeKind.Entry)) {
                    cancellationToken.ThrowIfCancellationRequested();
                    int start = -1;
                    int end = -1;
                    string? spanName = (string?)entry.Attribute("spanname");
                    if (!string.IsNullOrEmpty(spanName)) {
                        if (layout.NamedSpans.TryGetValue(spanName!, out KeyValuePair<int, int> namedSpan)) {
                            start = namedSpan.Key;
                            end = namedSpan.Value;
                        } else {
                            flattenedCalsLayout = true;
                        }
                    }
                    string? startName = (string?)entry.Attribute("namest") ?? (string?)entry.Attribute("colname");
                    if (start < 0 && !string.IsNullOrEmpty(startName)) {
                        if (layout.NamedColumns.TryGetValue(startName!, out int namedStart)) start = namedStart;
                        else flattenedCalsLayout = true;
                    }
                    string? endName = (string?)entry.Attribute("nameend");
                    if (end < 0 && !string.IsNullOrEmpty(endName)) {
                        if (layout.NamedColumns.TryGetValue(endName!, out int namedEnd)) end = namedEnd;
                        else flattenedCalsLayout = true;
                    }
                    if (start < 0) {
                        start = nextEntryColumn;
                        while (start < cells.Count && cells[start] != null) start++;
                    }
                    if (end < start) end = start;
                    if (start >= rowColumnLimit) {
                        projectionTruncated = true;
                        continue;
                    }
                    if (end >= rowColumnLimit) {
                        end = rowColumnLimit - 1;
                        projectionTruncated = true;
                    }
                    while (cells.Count <= end) cells.Add(null);
                    if (cells[start] != null) {
                        flattenedCalsLayout = true;
                        while (start < cells.Count && cells[start] != null) start++;
                        if (start >= rowColumnLimit) {
                            projectionTruncated = true;
                            continue;
                        }
                        end = start;
                        while (cells.Count <= end) cells.Add(null);
                    }
                    cells[start] = textBudget.GetPrimaryText(entry, DocBookNodeKind.Entry, Namespace, string.Empty);
                    for (int column = start + 1; column <= end; column++) cells[column] = string.Empty;
                    nextEntryColumn = Math.Max(nextEntryColumn, end + 1);

                    int moreRows = 0;
                    if (int.TryParse((string?)entry.Attribute("morerows"), out int parsedMoreRows) && parsedMoreRows > 0) moreRows = parsedMoreRows;
                    for (int column = start; column <= end && moreRows > 0; column++) {
                        if (!activeSpans.TryGetValue(column, out int existingRows) || existingRows < moreRows) activeSpans[column] = moreRows;
                    }
                    if (end > start || moreRows > 0 || !string.IsNullOrEmpty(spanName)) flattenedCalsLayout = true;
                }

                while (cells.Count > 0 && cells[cells.Count - 1] == null && cells.Count > layout.DeclaredColumns) cells.RemoveAt(cells.Count - 1);
                columnCount = Math.Max(columnCount, cells.Count);
                projectedRows.Add(new KeyValuePair<bool, IReadOnlyList<string>>(isHeader,
                    cells.Select(cell => cell ?? string.Empty).ToArray()));
            }

            var headerRows = new List<IReadOnlyList<string>>();
            var bodyRows = new List<IReadOnlyList<string>>();
            foreach (KeyValuePair<bool, IReadOnlyList<string>> projectedRow in projectedRows) {
                cancellationToken.ThrowIfCancellationRequested();
                (projectedRow.Key ? headerRows : bodyRows).Add(projectedRow.Value);
            }
            if (headerRows.Count > 1) flattenedCalsLayout = true;
            var columns = new List<string>(columnCount);
            if (headerRows.Count > 0) {
                for (int column = 0; column < columnCount; column++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    var labels = new List<string>();
                    foreach (IReadOnlyList<string> headerRow in headerRows) {
                        cancellationToken.ThrowIfCancellationRequested();
                        string label = column < headerRow.Count ? headerRow[column] : string.Empty;
                        if (!string.IsNullOrWhiteSpace(label)) labels.Add(label);
                    }
                    columns.Add(string.Join(" / ", labels));
                }
            }
            while (columns.Count < columnCount) columns.Add("Column " + (columns.Count + 1));
            for (int column = 0; column < columns.Count; column++) {
                if (string.IsNullOrWhiteSpace(columns[column])) columns[column] = "Column " + (column + 1);
            }
            var rows = new List<IReadOnlyList<string>>();
            foreach (IReadOnlyList<string> sourceRow in bodyRows) {
                cancellationToken.ThrowIfCancellationRequested();
                rows.Add(sourceRow.Count == columnCount
                    ? sourceRow
                    : sourceRow.Concat(Enumerable.Repeat(string.Empty, columnCount - sourceRow.Count)).ToArray());
            }
            XElement? tableTitleElement = tableElement.Element(Namespace + "title");
            string? title = tableTitleElement == null ? null : textBudget.GetElementValue(tableTitleElement, string.Empty);
            string tableHeadingPath = BuildTableHeadingPath(tableElement, title);
            if (flattenedCalsLayout) {
                diagnostics.Add(new DocBookDiagnostic("DB112", DocBookDiagnosticSeverity.Warning,
                    "CALS groups, spans, or multi-row headers were flattened in the shared table projection; the recursive structure retains the native markup.", tableHeadingPath));
            }
            if (flattenedFooterRows) {
                diagnostics.Add(new DocBookDiagnostic("DB119", DocBookDiagnosticSeverity.Warning,
                    "CALS footer rows were flattened into shared body rows; the recursive structure retains the native tfoot markup.", tableHeadingPath));
            }
            if (nestedEntryTableEncountered) {
                diagnostics.Add(new DocBookDiagnostic("DB121", DocBookDiagnosticSeverity.Warning,
                    "Nested CALS entrytbl rows were omitted from the outer shared table projection; the recursive structure retains the nested table markup.", tableHeadingPath));
            }
            if (projectionTruncated) {
                diagnostics.Add(new DocBookDiagnostic("DB113", DocBookDiagnosticSeverity.Warning,
                    "CALS geometry exceeded the configured shared table projection limits; the recursive structure retains the native markup.", tableHeadingPath));
            }
            int tableIndex = tables.Count;
            tableIndexes[tableElement] = tableIndex;
            tables.Add(new OfficeDocumentModelTable {
                Title = title,
                Kind = tableElement.Name.LocalName,
                Location = new OfficeDocumentModelLocation {
                    Path = sourcePath, HeadingPath = tableHeadingPath, SourceBlockKind = "table", TableIndex = tableIndex
                },
                Columns = columns,
                Rows = rows,
                TotalRowCount = totalBodyRows,
                Truncated = projectionTruncated || rows.Count < totalBodyRows
            });
            int retainedCellSlots = columns.Count;
            foreach (IReadOnlyList<string> row in rows) {
                cancellationToken.ThrowIfCancellationRequested();
                retainedCellSlots += row.Count;
            }
            remainingTableCells -= retainedCellSlots;
        }

        XElement? firstAuthor = null;
        XElement? documentInfo = FindInfo();
        if (documentInfo != null) {
            foreach (XElement element in documentInfo.Descendants()) {
                cancellationToken.ThrowIfCancellationRequested();
                if (DocBookNames.GetKind(element.Name, Namespace) == DocBookNodeKind.Author) {
                    firstAuthor = element;
                    break;
                }
            }
        }
        string? author = firstAuthor == null ? null : textBudget.GetAuthorName(firstAuthor, string.Empty);
        var metadata = new List<OfficeDocumentModelMetadataEntry> {
            new OfficeDocumentModelMetadataEntry {
                Id = "docbook-profile", Category = "docbook", Name = "profile",
                Value = Profile == DocBookProfile.DocBook45 ? "4.5" : "5.2", ValueType = "string"
            },
            new OfficeDocumentModelMetadataEntry {
                Id = "docbook-kind", Category = "docbook", Name = "kind",
                Value = Kind == DocBookDocumentKind.Book ? "book" : "article", ValueType = "string"
            }
        };
        int authorIndex = 0;
        foreach (XElement authorElement in RootElement.Descendants()) {
            cancellationToken.ThrowIfCancellationRequested();
            if (DocBookNames.GetKind(authorElement.Name, Namespace) != DocBookNodeKind.Author) continue;
            metadata.Add(new OfficeDocumentModelMetadataEntry {
                Id = "docbook-author-" + authorIndex++, Category = "docbook", Name = "author",
                Value = textBudget.GetAuthorName(authorElement, string.Empty), ValueType = "string"
            });
        }

        var structureItems = new List<OfficeDocumentModelNode>();
        foreach (XElement child in RootElement.Elements()) {
            cancellationToken.ThrowIfCancellationRequested();
            structureItems.Add(Convert(child, 1, string.Empty));
        }
        OfficeDocumentModelNode[] structure = structureItems.ToArray();
        var capabilities = new List<string> {
            "docbook.common-structure", "docbook.extensions", Profile == DocBookProfile.DocBook45 ? "docbook.4.5" : "docbook.5.2"
        };
        if (assets.Count > 0) capabilities.Add("docbook.media-references");
        var model = new OfficeDocumentModel {
            Format = OfficeDocumentFormat.DocBook,
            Source = new OfficeDocumentModelSource { Path = sourcePath, Title = GetProjectedDocumentTitle(), Author = author },
            CapabilitiesUsed = capabilities,
            Metadata = metadata,
            Structure = structure,
            Blocks = blocks,
            Tables = tables,
            Assets = assets,
            Links = links
        };
        return new DocBookConversionResult<OfficeDocumentModel>(model, diagnostics.ToArray());

        string BuildTableHeadingPath(XElement tableElement, string? tableTitle) {
            string path = string.Empty;
            foreach (XElement section in tableElement.Ancestors().Reverse().Where(ancestor =>
                         DocBookNames.GetKind(ancestor.Name, Namespace) == DocBookNodeKind.Section)) {
                cancellationToken.ThrowIfCancellationRequested();
                path = OfficeDocumentHeadingPath.Append(path,
                    textBudget.GetPrimaryText(section, DocBookNodeKind.Section, Namespace, path), " / ");
            }
            return OfficeDocumentHeadingPath.Append(path, tableTitle, " / ");
        }

        string? GetProjectedDocumentTitle() {
            XElement? info = FindInfo();
            XElement? titleElement = info?.Element(Namespace + "title") ?? RootElement.Element(Namespace + "title");
            return titleElement == null ? null : textBudget.GetElementValue(titleElement, string.Empty);
        }

        static string? GetReferenceFileName(string fileReference) {
            int delimiter = fileReference.IndexOfAny(new[] { '?', '#' });
            string clean = delimiter < 0 ? fileReference : fileReference.Substring(0, delimiter);
            int separator = Math.Max(clean.LastIndexOf('/'), clean.LastIndexOf('\\'));
            string fileName = separator < 0 ? clean : clean.Substring(separator + 1);
            return string.IsNullOrWhiteSpace(fileName) ? null : fileName;
        }

        static string? GetReferenceExtension(string? fileName) {
            if (string.IsNullOrWhiteSpace(fileName)) return null;
            int dot = fileName!.LastIndexOf('.');
            return dot < 0 || dot == fileName.Length - 1 ? null : fileName.Substring(dot);
        }
    }

    /// <summary>Creates an article or book from the shared recursive common structure.</summary>
    public static DocBookConversionResult<DocBookDocument> FromOfficeDocumentModel(
        OfficeDocumentModel model,
        DocBookDocumentKind? kind = null,
        DocBookProfile? profile = null,
        DocBookConversionOptions? options = null) {
        if (model == null) throw new ArgumentNullException(nameof(model));
        options ??= new DocBookConversionOptions();
        options.Validate();
        IReadOnlyList<OfficeDocumentModelNode> structureNodes = OfficeDocumentModelStructureTraversal.ValidateAndFlatten(
            model.Structure, options.MaxStructureDepth, options.MaxStructureNodes);
        var diagnostics = new DocBookDiagnosticCollector(options.MaxDetailedDiagnosticsPerCode);
        string? sourceKind = model.Metadata.FirstOrDefault(entry => entry.Category == "docbook" && entry.Name == "kind")?.Value;
        string? sourceProfile = model.Metadata.FirstOrDefault(entry => entry.Category == "docbook" && entry.Name == "profile")?.Value;
        bool sourceKindIsSupported = sourceKind == null || sourceKind == "article" || sourceKind == "book";
        bool sourceProfileIsSupported = sourceProfile == null || sourceProfile == "4.5" || sourceProfile == "5.2";
        DocBookDocumentKind inferredKind = sourceKind == "book" ? DocBookDocumentKind.Book : DocBookDocumentKind.Article;
        DocBookProfile inferredProfile = sourceProfile == "4.5" ? DocBookProfile.DocBook45 : DocBookProfile.DocBook52;
        DocBookDocumentKind selectedKind = kind ?? inferredKind;
        DocBookProfile selectedProfile = profile ?? inferredProfile;
        XNamespace sourceDocBookNamespace = inferredProfile == DocBookProfile.DocBook52
            ? DocBookSchemaProfiles.DocBook52.NamespaceUri : XNamespace.None;
        XNamespace targetDocBookNamespace = selectedProfile == DocBookProfile.DocBook52
            ? DocBookSchemaProfiles.DocBook52.NamespaceUri : XNamespace.None;
        bool requalifySourceVocabulary = sourceProfile != null && sourceProfileIsSupported && selectedProfile != inferredProfile;
        DocBookDocument document = selectedKind == DocBookDocumentKind.Article ? CreateArticle(selectedProfile) : CreateBook(selectedProfile);
        if (!sourceKindIsSupported) {
            diagnostics.Add(new DocBookDiagnostic("DB114", DocBookDiagnosticSeverity.Warning,
                $"The unsupported source DocBook kind '{sourceKind}' was normalized to '{selectedKind.ToString().ToLowerInvariant()}'."));
        }
        if (!sourceProfileIsSupported) {
            diagnostics.Add(new DocBookDiagnostic("DB114", DocBookDiagnosticSeverity.Warning,
                $"The unsupported source DocBook profile '{sourceProfile}' was normalized to '{(selectedProfile == DocBookProfile.DocBook45 ? "4.5" : "5.2")}'."));
        }
        if (kind.HasValue && sourceKind != null && selectedKind != inferredKind) {
            diagnostics.Add(new DocBookDiagnostic("DB108", DocBookDiagnosticSeverity.Warning,
                $"The source root kind '{sourceKind}' was changed to '{selectedKind.ToString().ToLowerInvariant()}' by the requested conversion."));
        }
        if (profile.HasValue && sourceProfile != null && selectedProfile != inferredProfile) {
            diagnostics.Add(new DocBookDiagnostic("DB109", DocBookDiagnosticSeverity.Warning,
                $"The source profile '{sourceProfile}' was changed to '{(selectedProfile == DocBookProfile.DocBook45 ? "4.5" : "5.2")}' by the requested conversion."));
        }

        void Add(OfficeDocumentModelNode source, DocBookNode parent) {
            DocBookNode target;
            if (string.Equals(source.Kind, "text", StringComparison.OrdinalIgnoreCase)) {
                parent.AddText(source.Text);
                return;
            }
            if (source.Kind.StartsWith("extension:", StringComparison.Ordinal)) {
                string expandedName = source.Kind.Substring("extension:".Length);
                try {
                    XName extensionName = XName.Get(expandedName);
                    if (requalifySourceVocabulary && extensionName.Namespace == sourceDocBookNamespace) {
                        extensionName = targetDocBookNamespace + extensionName.LocalName;
                    }
                    target = parent.AddExtension(extensionName, source.Children.Count == 0 ? source.Text : null);
                } catch (Exception) {
                    target = parent.Add(DocBookNodeKind.Paragraph, source.Text);
                    diagnostics.Add(new DocBookDiagnostic("DB104", DocBookDiagnosticSeverity.Warning,
                        $"Extension node name '{expandedName}' could not be reconstructed and was represented as a paragraph.", source.Location.HeadingPath));
                }
            } else if (string.Equals(source.Kind, "informal-table", StringComparison.OrdinalIgnoreCase)) {
                target = parent.AddRaw("informaltable");
            } else if (TryMapKind(source.Kind, out DocBookNodeKind nodeKind)) {
                string? directText = NodeAcceptsDirectText(nodeKind) && source.Children.Count == 0 ? source.Text : null;
                bool externalLink = nodeKind == DocBookNodeKind.Link &&
                    (source.Attributes.ContainsKey("url") || source.Attributes.ContainsKey("{http://www.w3.org/1999/xlink}href"));
                target = nodeKind == DocBookNodeKind.Link && selectedProfile == DocBookProfile.DocBook45 && externalLink
                    ? parent.AddRaw("ulink", directText) : parent.Add(nodeKind, directText);
                if (!string.IsNullOrEmpty(source.Text) && !NodeAcceptsDirectText(nodeKind) &&
                    !SourceChildrenRepresentText(source, nodeKind)) {
                    if (NodeUsesTitleText(nodeKind)) {
                        target.Add(DocBookNodeKind.Title, source.Text);
                    } else if (NodeUsesParagraphText(nodeKind)) {
                        target.AddParagraph(source.Text);
                    } else if (nodeKind == DocBookNodeKind.IndexTerm) {
                        target.AddRaw("primary", source.Text);
                    } else {
                        diagnostics.Add(new DocBookDiagnostic("DB116", DocBookDiagnosticSeverity.Warning,
                            $"Text on shared container node '{source.Kind}' could not be represented by the selected DocBook profile.",
                            source.Location.HeadingPath));
                    }
                }
            } else {
                target = parent.Add(DocBookNodeKind.Paragraph, source.Text);
                target.SetAttribute("role", "officeimo-" + SanitizeRole(source.Kind));
                diagnostics.Add(new DocBookDiagnostic("DB101", DocBookDiagnosticSeverity.Warning,
                    $"Shared node kind '{source.Kind}' was represented as a role-qualified paragraph.", source.Location.HeadingPath));
            }
            foreach (KeyValuePair<string, string> attribute in source.Attributes) {
                try {
                    XName attributeName = XName.Get(attribute.Key);
                    if (target.Kind == DocBookNodeKind.Link &&
                        (attributeName == XName.Get("url") || attributeName == XName.Get("href", "http://www.w3.org/1999/xlink"))) {
                        attributeName = selectedProfile == DocBookProfile.DocBook45
                            ? XName.Get("url") : XName.Get("href", "http://www.w3.org/1999/xlink");
                    }
                    target.SetAttribute(attributeName, attribute.Value);
                } catch (Exception exception) when (exception is ArgumentException || exception is System.Xml.XmlException) {
                    diagnostics.Add(new DocBookDiagnostic("DB102", DocBookDiagnosticSeverity.Warning,
                        $"Attribute name '{attribute.Key}' could not be represented.", source.Location.HeadingPath));
                }
            }
            foreach (OfficeDocumentModelNode child in source.Children) Add(child, target);
        }

        void AddFlatTable(OfficeDocumentModelTable source) {
            bool formal = !string.IsNullOrWhiteSpace(source.Title);
            DocBookNode table = formal ? document.Root.Add(DocBookNodeKind.Table) : document.Root.AddRaw("informaltable");
            if (formal) table.Add(DocBookNodeKind.Title, source.Title);
            DocBookNode group = table.Add(DocBookNodeKind.TableGroup);
            int rowWidth = source.Rows.Count == 0 ? 0 : source.Rows.Max(row => row.Count);
            int columnCount = Math.Max(1, Math.Max(source.Columns.Count, rowWidth));
            group.SetAttribute("cols", columnCount.ToString(System.Globalization.CultureInfo.InvariantCulture));
            if (source.Columns.Count > 0) {
                DocBookNode headerRow = group.Add(DocBookNodeKind.TableHead).Add(DocBookNodeKind.Row);
                foreach (string column in source.Columns) headerRow.Add(DocBookNodeKind.Entry, column);
            }
            DocBookNode body = group.Add(DocBookNodeKind.TableBody);
            foreach (IReadOnlyList<string> sourceRow in source.Rows) {
                DocBookNode row = body.Add(DocBookNodeKind.Row);
                foreach (string cell in sourceRow) row.Add(DocBookNodeKind.Entry, cell);
            }
            if (source.Truncated || source.TotalRowCount > source.Rows.Count) {
                diagnostics.Add(new DocBookDiagnostic("DB117", DocBookDiagnosticSeverity.Warning,
                    "The shared flat table was already truncated; only its available rows were emitted.", source.Title));
            }
        }

        bool AddFlatAsset(OfficeDocumentModelAsset source) {
            string? reference = model.Format == OfficeDocumentFormat.DocBook && !string.IsNullOrWhiteSpace(source.SourceObjectId)
                ? source.SourceObjectId
                : source.FileName;
            if (!string.Equals(source.Kind, "image", StringComparison.OrdinalIgnoreCase) || string.IsNullOrWhiteSpace(reference)) {
                diagnostics.Add(new DocBookDiagnostic("DB118", DocBookDiagnosticSeverity.Warning,
                    $"Shared asset '{source.Id}' could not be represented as a DocBook image reference.", source.Location?.HeadingPath));
                return false;
            }
            document.Root.AddImage(reference!, source.Title, source.AltText);
            return true;
        }

        bool AddFlatLink(OfficeDocumentModelLink source) {
            string text = source.Text ?? source.Uri ?? source.DestinationName ?? source.Id;
            DocBookNode paragraph = document.Root.Add(DocBookNodeKind.Paragraph);
            bool represented = false;
            if (!string.IsNullOrWhiteSpace(source.Uri)) {
                DocBookNode link = selectedProfile == DocBookProfile.DocBook45
                    ? paragraph.AddRaw("ulink", text)
                    : paragraph.Add(DocBookNodeKind.Link, text);
                link.SetAttribute(selectedProfile == DocBookProfile.DocBook45
                    ? XName.Get("url") : XName.Get("href", "http://www.w3.org/1999/xlink"), source.Uri);
                represented = true;
            } else if (!string.IsNullOrWhiteSpace(source.DestinationName)) {
                DocBookNode link = paragraph.Add(DocBookNodeKind.Link, text);
                link.SetAttribute("linkend", source.DestinationName);
                represented = true;
            }
            bool hasUnsupportedTarget = source.DestinationPageNumber.HasValue || !string.IsNullOrWhiteSpace(source.DestinationMode) ||
                !string.IsNullOrWhiteSpace(source.NamedAction) || !string.IsNullOrWhiteSpace(source.RemoteFile) ||
                !string.IsNullOrWhiteSpace(source.RemoteDestinationName) || source.RemoteDestinationPageNumber.HasValue ||
                (!string.IsNullOrWhiteSpace(source.Uri) && !string.IsNullOrWhiteSpace(source.DestinationName));
            if (!represented || hasUnsupportedTarget) {
                if (!represented) paragraph.Remove();
                diagnostics.Add(new DocBookDiagnostic("DB120", DocBookDiagnosticSeverity.Warning,
                    represented
                        ? $"Shared link '{source.Id}' was emitted, but one or more additional target fields could not be represented in DocBook."
                        : $"Shared link '{source.Id}' had no DocBook-representable URI or named destination.",
                    source.Location?.HeadingPath));
            }
            return represented;
        }

        if (model.Structure.Count > 0) {
            bool hasTitle = model.Structure.Any(ContainsDocumentTitle);
            foreach (OfficeDocumentModelNode node in model.Structure) Add(node, document.Root);
            if (!hasTitle) document.Title = model.Source.Title;
            var consumedTableNodes = new HashSet<OfficeDocumentModelNode>();
            foreach (OfficeDocumentModelBlock block in model.Blocks.Where(block => !IsDerivedBlock(block, structureNodes))) {
                document.AddParagraph(block.Text);
                AddSupplementaryDiagnostic("block", block.Id, block.Location?.HeadingPath);
            }
            foreach (OfficeDocumentModelTable table in model.Tables.Where(table => !IsDerivedTable(table, structureNodes, consumedTableNodes))) {
                AddFlatTable(table);
                AddSupplementaryDiagnostic("table", table.Title ?? table.Kind ?? "unnamed", table.Location?.HeadingPath);
            }
            foreach (OfficeDocumentModelAsset asset in model.Assets.Where(asset => !IsDerivedAsset(asset, structureNodes))) {
                if (AddFlatAsset(asset)) AddSupplementaryDiagnostic("asset", asset.Id, asset.Location?.HeadingPath);
            }
            foreach (OfficeDocumentModelLink link in model.Links.Where(link => !IsDerivedLink(link, structureNodes))) {
                if (AddFlatLink(link)) AddSupplementaryDiagnostic("link", link.Id, link.Location?.HeadingPath);
            }
        } else {
            document.Title = model.Source.Title;
            diagnostics.Add(new DocBookDiagnostic("DB103", DocBookDiagnosticSeverity.Warning,
                "The shared model had no recursive Structure; flat Blocks, Tables, image Assets, and Links were emitted as common DocBook structures."));
            foreach (OfficeDocumentModelBlock block in model.Blocks) document.AddParagraph(block.Text);
            foreach (OfficeDocumentModelTable table in model.Tables) AddFlatTable(table);
            foreach (OfficeDocumentModelAsset asset in model.Assets) AddFlatAsset(asset);
            foreach (OfficeDocumentModelLink link in model.Links) AddFlatLink(link);
        }
        foreach (OfficeDocumentModelPage page in model.Pages) {
            string identity = page.Name ?? (page.Number.HasValue
                ? page.Number.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)
                : "unnamed");
            AddUnsupportedChannelDiagnostic("page", identity, page.Location?.HeadingPath);
        }
        foreach (OfficeDocumentModelFormField form in model.Forms) {
            AddUnsupportedChannelDiagnostic("form", form.Name ?? form.Id, form.Location?.HeadingPath);
        }
        foreach (OfficeDocumentModelVisual visual in model.Visuals) {
            AddUnsupportedChannelDiagnostic("visual", visual.SourceName ?? visual.Kind, visual.Location?.HeadingPath);
        }
        bool hasAuthor = model.Structure.Any(ContainsDocumentAuthor);
        if (!hasAuthor && !string.IsNullOrWhiteSpace(model.Source.Author)) {
            new DocBookNode(document, document.EnsureInfo()).Add(DocBookNodeKind.Author, model.Source.Author);
        }
        return new DocBookConversionResult<DocBookDocument>(document, diagnostics.ToArray());

        void AddSupplementaryDiagnostic(string channel, string identity, string? path) =>
            diagnostics.Add(new DocBookDiagnostic("DB122", DocBookDiagnosticSeverity.Warning,
                $"Supplementary shared {channel} '{identity}' was appended at the document root because it was not represented by recursive Structure.", path));

        void AddUnsupportedChannelDiagnostic(string channel, string identity, string? path) =>
            diagnostics.Add(new DocBookDiagnostic("DB124", DocBookDiagnosticSeverity.Warning,
                $"Shared {channel} '{identity}' could not be represented by the bounded DocBook common-structure profile.", path));

        static bool ContainsDocumentTitle(OfficeDocumentModelNode node) =>
            string.Equals(node.Kind, "title", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(node.Kind, "metadata", StringComparison.OrdinalIgnoreCase) &&
            node.Children.Any(ContainsTitleNode);

        static bool ContainsTitleNode(OfficeDocumentModelNode node) =>
            string.Equals(node.Kind, "title", StringComparison.OrdinalIgnoreCase) || node.Children.Any(ContainsTitleNode);

        static bool ContainsDocumentAuthor(OfficeDocumentModelNode node) =>
            string.Equals(node.Kind, "author", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(node.Kind, "metadata", StringComparison.OrdinalIgnoreCase) &&
            node.Children.Any(ContainsAuthorNode);

        static bool ContainsAuthorNode(OfficeDocumentModelNode node) =>
            string.Equals(node.Kind, "author", StringComparison.OrdinalIgnoreCase) || node.Children.Any(ContainsAuthorNode);
    }

    private static string ToModelKind(DocBookNodeKind kind) {
        switch (kind) {
            case DocBookNodeKind.Info: return "metadata";
            case DocBookNodeKind.Title: return "title";
            case DocBookNodeKind.Subtitle: return "subtitle";
            case DocBookNodeKind.Author: return "author";
            case DocBookNodeKind.Section: return "section";
            case DocBookNodeKind.Paragraph: return "paragraph";
            case DocBookNodeKind.ItemizedList: return "itemized-list";
            case DocBookNodeKind.OrderedList: return "ordered-list";
            case DocBookNodeKind.VariableList: return "variable-list";
            case DocBookNodeKind.ListItem: return "list-item";
            case DocBookNodeKind.Table: return "table";
            case DocBookNodeKind.TableGroup: return "table-group";
            case DocBookNodeKind.TableHead: return "table-head";
            case DocBookNodeKind.TableBody: return "table-body";
            case DocBookNodeKind.Row: return "table-row";
            case DocBookNodeKind.Entry: return "table-cell";
            case DocBookNodeKind.ProgramListing: return "code";
            case DocBookNodeKind.Screen: return "screen";
            case DocBookNodeKind.Link: return "link";
            case DocBookNodeKind.CrossReference: return "cross-reference";
            case DocBookNodeKind.Note: return "note";
            case DocBookNodeKind.Tip: return "tip";
            case DocBookNodeKind.Important: return "important";
            case DocBookNodeKind.Caution: return "caution";
            case DocBookNodeKind.Warning: return "warning";
            case DocBookNodeKind.Figure: return "figure";
            case DocBookNodeKind.MediaObject: return "media";
            case DocBookNodeKind.ImageObject: return "image-object";
            case DocBookNodeKind.ImageData: return "image";
            case DocBookNodeKind.Caption: return "caption";
            case DocBookNodeKind.Index: return "index";
            case DocBookNodeKind.IndexTerm: return "index-term";
            default: return "unknown";
        }
    }

    private static bool TryMapKind(string kind, out DocBookNodeKind nodeKind) {
        var mappings = new Dictionary<string, DocBookNodeKind>(StringComparer.OrdinalIgnoreCase) {
            ["metadata"] = DocBookNodeKind.Info,
            ["title"] = DocBookNodeKind.Title,
            ["subtitle"] = DocBookNodeKind.Subtitle,
            ["author"] = DocBookNodeKind.Author,
            ["section"] = DocBookNodeKind.Section,
            ["paragraph"] = DocBookNodeKind.Paragraph,
            ["itemized-list"] = DocBookNodeKind.ItemizedList,
            ["ordered-list"] = DocBookNodeKind.OrderedList,
            ["variable-list"] = DocBookNodeKind.VariableList,
            ["list-item"] = DocBookNodeKind.ListItem,
            ["table"] = DocBookNodeKind.Table,
            ["table-group"] = DocBookNodeKind.TableGroup,
            ["table-head"] = DocBookNodeKind.TableHead,
            ["table-body"] = DocBookNodeKind.TableBody,
            ["table-row"] = DocBookNodeKind.Row,
            ["table-cell"] = DocBookNodeKind.Entry,
            ["code"] = DocBookNodeKind.ProgramListing,
            ["screen"] = DocBookNodeKind.Screen,
            ["link"] = DocBookNodeKind.Link,
            ["cross-reference"] = DocBookNodeKind.CrossReference,
            ["note"] = DocBookNodeKind.Note,
            ["tip"] = DocBookNodeKind.Tip,
            ["important"] = DocBookNodeKind.Important,
            ["caution"] = DocBookNodeKind.Caution,
            ["warning"] = DocBookNodeKind.Warning,
            ["figure"] = DocBookNodeKind.Figure,
            ["media"] = DocBookNodeKind.MediaObject,
            ["image-object"] = DocBookNodeKind.ImageObject,
            ["image"] = DocBookNodeKind.ImageData,
            ["caption"] = DocBookNodeKind.Caption,
            ["index"] = DocBookNodeKind.Index,
            ["index-term"] = DocBookNodeKind.IndexTerm
        };
        return mappings.TryGetValue(kind ?? string.Empty, out nodeKind);
    }

    private static bool NodeAcceptsDirectText(DocBookNodeKind kind) =>
        kind == DocBookNodeKind.Title || kind == DocBookNodeKind.Subtitle || kind == DocBookNodeKind.Paragraph ||
        kind == DocBookNodeKind.ProgramListing || kind == DocBookNodeKind.Screen || kind == DocBookNodeKind.Entry ||
        kind == DocBookNodeKind.Link || kind == DocBookNodeKind.Author;

    private static bool NodeUsesTitleText(DocBookNodeKind kind) =>
        kind == DocBookNodeKind.Info || kind == DocBookNodeKind.Section || kind == DocBookNodeKind.Table ||
        kind == DocBookNodeKind.Figure;

    private static bool NodeUsesParagraphText(DocBookNodeKind kind) =>
        kind == DocBookNodeKind.ListItem || kind == DocBookNodeKind.Note || kind == DocBookNodeKind.Tip ||
        kind == DocBookNodeKind.Important || kind == DocBookNodeKind.Caution || kind == DocBookNodeKind.Warning;

    private static bool SourceChildrenRepresentText(OfficeDocumentModelNode source, DocBookNodeKind kind) {
        string representedKind = NodeUsesTitleText(kind) ? "title" : NodeUsesParagraphText(kind) ? "paragraph" : string.Empty;
        return source.Children.Any(child =>
            (child.Kind == "text" || representedKind.Length > 0 && string.Equals(child.Kind, representedKind, StringComparison.OrdinalIgnoreCase)) &&
            string.Equals(child.Text, source.Text, StringComparison.Ordinal));
    }

    private static string SanitizeRole(string value) => new string((value ?? "unknown").Select(c => char.IsLetterOrDigit(c) || c == '-' ? c : '-').ToArray());

    private sealed class CalsProjectionLayout {
        internal static readonly CalsProjectionLayout Empty = new CalsProjectionLayout(
            0, 0,
            new Dictionary<string, int>(StringComparer.Ordinal),
            new Dictionary<string, KeyValuePair<int, int>>(StringComparer.Ordinal));

        internal CalsProjectionLayout(
            int declaredColumns,
            int nextColumn,
            Dictionary<string, int> namedColumns,
            Dictionary<string, KeyValuePair<int, int>> namedSpans) {
            DeclaredColumns = declaredColumns;
            NextColumn = nextColumn;
            NamedColumns = namedColumns;
            NamedSpans = namedSpans;
        }

        internal int DeclaredColumns { get; }
        internal int NextColumn { get; }
        internal Dictionary<string, int> NamedColumns { get; }
        internal Dictionary<string, KeyValuePair<int, int>> NamedSpans { get; }
    }
}
