using OfficeIMO;
using OfficeIMO.Core.Internal;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.DocBook;

public sealed partial class DocBookDocument {
    /// <summary>Converts the typed common structure to the shared recursive document model.</summary>
    public DocBookConversionResult<OfficeDocumentModel> ToOfficeDocumentModel(
        string? sourcePath = null,
        DocBookConversionOptions? options = null) {
        options ??= new DocBookConversionOptions();
        options.Validate();
        var diagnostics = new List<DocBookDiagnostic>();
        var blocks = new List<OfficeDocumentModelBlock>();
        var tables = new List<OfficeDocumentModelTable>();
        var links = new List<OfficeDocumentModelLink>();
        int index = 0;
        if (_xml.DescendantNodes().Any(node => node is XComment || node is XProcessingInstruction)) {
            diagnostics.Add(new DocBookDiagnostic("DB105", DocBookDiagnosticSeverity.Warning,
                "Comments and processing instructions remain native but are not represented by the shared document model."));
        }
        if (RootElement.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration && attribute.Name.LocalName != "version")) {
            diagnostics.Add(new DocBookDiagnostic("DB106", DocBookDiagnosticSeverity.Warning,
                "Root extension attributes remain native but are not represented by the shared document model.", "/" + RootElement.Name.LocalName));
        }
        if (RootElement.Nodes().OfType<XText>().Any(text => !string.IsNullOrWhiteSpace(text.Value))) {
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
            DocBookNodeKind kind = DocBookNames.GetKind(element.Name, Namespace);
            string normalizedKind = kind == DocBookNodeKind.Unknown
                ? "extension:" + element.Name
                : kind == DocBookNodeKind.Table && element.Name.LocalName == "informaltable"
                    ? "informal-table"
                    : ToModelKind(kind);
            string text = GetPrimaryText(element, kind);
            string path = kind == DocBookNodeKind.Section
                ? OfficeDocumentHeadingPath.Append(parentPath, text, " / ") : parentPath;
            var attributes = element.Attributes().ToDictionary(a => a.Name.ToString(), a => a.Value, StringComparer.Ordinal);
            int nodeIndex = index++;
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
                Location = new OfficeDocumentModelLocation { Path = sourcePath, HeadingPath = path }
            };
        }

        IReadOnlyList<OfficeDocumentModelNode> BuildChildren(XElement element, DocBookNodeKind kind, int level, string path) {
            bool mixedContent = kind == DocBookNodeKind.Unknown || kind == DocBookNodeKind.Paragraph || kind == DocBookNodeKind.Title || kind == DocBookNodeKind.Subtitle ||
                kind == DocBookNodeKind.Link || kind == DocBookNodeKind.Entry || kind == DocBookNodeKind.Caption || kind == DocBookNodeKind.Author ||
                kind == DocBookNodeKind.ProgramListing || kind == DocBookNodeKind.Screen;
            var children = new List<OfficeDocumentModelNode>();
            foreach (XNode node in element.Nodes()) {
                if (node is XElement child) {
                    children.Add(Convert(child, level + 1, path));
                } else if (node is XText textNode && textNode.Value.Length > 0 && (mixedContent || !string.IsNullOrWhiteSpace(textNode.Value))) {
                    children.Add(new OfficeDocumentModelNode {
                        Id = "docbook-" + index++,
                        Kind = "text",
                        Text = textNode.Value,
                        Level = level + 1,
                        Location = new OfficeDocumentModelLocation { Path = sourcePath, HeadingPath = path }
                    });
                }
            }
            return children;
        }

        foreach (XElement tableElement in RootElement.Descendants().Where(element =>
                     DocBookNames.GetKind(element.Name, Namespace) == DocBookNodeKind.Table)) {
            int discoveryCapacity = options.MaxTableRows > int.MaxValue / 2
                ? int.MaxValue : options.MaxTableRows * 2;
            var rowElements = new List<XElement>(Math.Min(discoveryCapacity, 4_096));
            bool rowDiscoveryTruncated = false;
            int totalBodyRows = 0;
            int headerRowsRetained = 0;
            int bodyRowsRetained = 0;
            foreach (XElement element in tableElement.Descendants()) {
                if (DocBookNames.GetKind(element.Name, Namespace) != DocBookNodeKind.Row ||
                    !ReferenceEquals(element.Ancestors().FirstOrDefault(ancestor =>
                        DocBookNames.GetKind(ancestor.Name, Namespace) == DocBookNodeKind.Table), tableElement)) continue;
                bool isHeader = element.Ancestors().TakeWhile(ancestor => !ReferenceEquals(ancestor, tableElement)).Any(ancestor =>
                    DocBookNames.GetKind(ancestor.Name, Namespace) == DocBookNodeKind.TableHead);
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
            bool flattenedCalsLayout = false;
            bool projectionTruncated = rowDiscoveryTruncated;

            foreach (XElement rowElement in rowElements) {
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
                initialWidth = Math.Min(initialWidth, options.MaxTableColumns);
                var cells = Enumerable.Repeat<string?>(null, initialWidth).ToList();
                foreach (KeyValuePair<int, int> activeSpan in activeSpans.ToArray()) {
                    if (activeSpan.Key >= options.MaxTableColumns) {
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
                        start = 0;
                        while (start < cells.Count && cells[start] != null) start++;
                    }
                    if (end < start) end = start;
                    if (start >= options.MaxTableColumns) {
                        projectionTruncated = true;
                        continue;
                    }
                    if (end >= options.MaxTableColumns) {
                        end = options.MaxTableColumns - 1;
                        projectionTruncated = true;
                    }
                    while (cells.Count <= end) cells.Add(null);
                    if (cells[start] != null) {
                        flattenedCalsLayout = true;
                        while (start < cells.Count && cells[start] != null) start++;
                        if (start >= options.MaxTableColumns) {
                            projectionTruncated = true;
                            continue;
                        }
                        end = start;
                        while (cells.Count <= end) cells.Add(null);
                    }
                    cells[start] = entry.Value;
                    for (int column = start + 1; column <= end; column++) cells[column] = string.Empty;

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

            List<IReadOnlyList<string>> headerRows = projectedRows.Where(row => row.Key).Select(row => row.Value).ToList();
            if (headerRows.Count > 1) flattenedCalsLayout = true;
            var columns = new List<string>(columnCount);
            if (headerRows.Count > 0) {
                for (int column = 0; column < columnCount; column++) {
                    columns.Add(string.Join(" / ", headerRows.Select(row => column < row.Count ? row[column] : string.Empty)
                        .Where(value => !string.IsNullOrWhiteSpace(value))));
                }
            }
            while (columns.Count < columnCount) columns.Add("Column " + (columns.Count + 1));
            for (int column = 0; column < columns.Count; column++) {
                if (string.IsNullOrWhiteSpace(columns[column])) columns[column] = "Column " + (column + 1);
            }
            var rows = new List<IReadOnlyList<string>>();
            foreach (IReadOnlyList<string> sourceRow in projectedRows.Where(row => !row.Key).Select(row => row.Value)) {
                rows.Add(sourceRow.Count == columnCount
                    ? sourceRow
                    : sourceRow.Concat(Enumerable.Repeat(string.Empty, columnCount - sourceRow.Count)).ToArray());
            }
            string? title = tableElement.Element(Namespace + "title")?.Value;
            if (flattenedCalsLayout) {
                diagnostics.Add(new DocBookDiagnostic("DB112", DocBookDiagnosticSeverity.Warning,
                    "CALS spans or multi-row headers were flattened in the shared table projection; the recursive structure retains the native markup.", title));
            }
            if (projectionTruncated) {
                diagnostics.Add(new DocBookDiagnostic("DB113", DocBookDiagnosticSeverity.Warning,
                    "CALS geometry exceeded the configured shared table projection limits; the recursive structure retains the native markup.", title));
            }
            tables.Add(new OfficeDocumentModelTable {
                Title = title,
                Kind = tableElement.Name.LocalName,
                Location = new OfficeDocumentModelLocation { Path = sourcePath, HeadingPath = title },
                Columns = columns,
                Rows = rows,
                TotalRowCount = totalBodyRows,
                Truncated = projectionTruncated || rows.Count < totalBodyRows
            });
        }

        XElement? firstAuthor = FindInfo()?.Descendants().FirstOrDefault(element =>
            DocBookNames.GetKind(element.Name, Namespace) == DocBookNodeKind.Author);
        string? author = firstAuthor == null ? null : GetAuthorText(firstAuthor);
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
        foreach (XElement authorElement in RootElement.Descendants().Where(element =>
                     DocBookNames.GetKind(element.Name, Namespace) == DocBookNodeKind.Author)) {
            metadata.Add(new OfficeDocumentModelMetadataEntry {
                Id = "docbook-author-" + authorIndex++, Category = "docbook", Name = "author",
                Value = GetAuthorText(authorElement), ValueType = "string"
            });
        }

        var model = new OfficeDocumentModel {
            Format = OfficeDocumentFormat.DocBook,
            Source = new OfficeDocumentModelSource { Path = sourcePath, Title = Title, Author = author },
            CapabilitiesUsed = new[] { "docbook.common-structure", "docbook.extensions", Profile == DocBookProfile.DocBook45 ? "docbook.4.5" : "docbook.5.2" },
            Metadata = metadata,
            Structure = RootElement.Elements().Select(child => Convert(child, 1, string.Empty)).ToArray(),
            Blocks = blocks,
            Tables = tables,
            Links = links
        };
        return new DocBookConversionResult<OfficeDocumentModel>(model, diagnostics);
    }

    /// <summary>Creates an article or book from the shared recursive common structure.</summary>
    public static DocBookConversionResult<DocBookDocument> FromOfficeDocumentModel(
        OfficeDocumentModel model,
        DocBookDocumentKind? kind = null,
        DocBookProfile? profile = null) {
        if (model == null) throw new ArgumentNullException(nameof(model));
        var diagnostics = new List<DocBookDiagnostic>();
        string? sourceKind = model.Metadata.FirstOrDefault(entry => entry.Category == "docbook" && entry.Name == "kind")?.Value;
        string? sourceProfile = model.Metadata.FirstOrDefault(entry => entry.Category == "docbook" && entry.Name == "profile")?.Value;
        bool sourceKindIsSupported = sourceKind == null || sourceKind == "article" || sourceKind == "book";
        bool sourceProfileIsSupported = sourceProfile == null || sourceProfile == "4.5" || sourceProfile == "5.2";
        DocBookDocumentKind inferredKind = sourceKind == "book" ? DocBookDocumentKind.Book : DocBookDocumentKind.Article;
        DocBookProfile inferredProfile = sourceProfile == "4.5" ? DocBookProfile.DocBook45 : DocBookProfile.DocBook52;
        DocBookDocumentKind selectedKind = kind ?? inferredKind;
        DocBookProfile selectedProfile = profile ?? inferredProfile;
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
                    target = parent.AddExtension(XName.Get(expandedName), source.Children.Count == 0 ? source.Text : null);
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

        if (model.Structure.Count > 0) {
            bool hasTitle = model.Structure.Any(ContainsDocumentTitle);
            foreach (OfficeDocumentModelNode node in model.Structure) Add(node, document.Root);
            if (!hasTitle) document.Title = model.Source.Title;
        } else {
            document.Title = model.Source.Title;
            diagnostics.Add(new DocBookDiagnostic("DB103", DocBookDiagnosticSeverity.Warning,
                "The shared model had no recursive Structure; flat Blocks were emitted as paragraphs."));
            foreach (OfficeDocumentModelBlock block in model.Blocks) document.AddParagraph(block.Text);
        }
        return new DocBookConversionResult<DocBookDocument>(document, diagnostics);

        static bool ContainsDocumentTitle(OfficeDocumentModelNode node) =>
            string.Equals(node.Kind, "title", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(node.Kind, "metadata", StringComparison.OrdinalIgnoreCase) &&
            node.Children.Any(child => string.Equals(child.Kind, "title", StringComparison.OrdinalIgnoreCase));
    }

    private static string GetPrimaryText(XElement element, DocBookNodeKind kind) {
        if (kind == DocBookNodeKind.Section || kind == DocBookNodeKind.Table || kind == DocBookNodeKind.Figure || kind == DocBookNodeKind.Info) {
            return element.Element(element.Name.Namespace + "title")?.Value ?? string.Empty;
        }
        if (kind == DocBookNodeKind.Author) return GetAuthorText(element);
        if (kind == DocBookNodeKind.Title || kind == DocBookNodeKind.Subtitle || kind == DocBookNodeKind.Link ||
            kind == DocBookNodeKind.Entry || kind == DocBookNodeKind.Caption) return element.Value;
        return element.HasElements && kind != DocBookNodeKind.Paragraph && kind != DocBookNodeKind.ProgramListing && kind != DocBookNodeKind.Screen
            ? string.Empty : element.Value;
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

    private static string SanitizeRole(string value) => new string((value ?? "unknown").Select(c => char.IsLetterOrDigit(c) || c == '-' ? c : '-').ToArray());

    private static string GetAuthorText(XElement authorElement) {
        if (!authorElement.HasElements) return authorElement.Value;
        string[] parts = authorElement.DescendantNodes().OfType<XText>()
            .Select(text => text.Value.Trim())
            .Where(value => value.Length > 0)
            .ToArray();
        return parts.Length == 0 ? authorElement.Value : string.Join(" ", parts);
    }

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
