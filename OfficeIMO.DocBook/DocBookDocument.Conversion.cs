using OfficeIMO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.DocBook;

public sealed partial class DocBookDocument {
    /// <summary>Converts the typed common structure to the shared recursive document model.</summary>
    public DocBookConversionResult<OfficeDocumentModel> ToOfficeDocumentModel(string? sourcePath = null) {
        var diagnostics = new List<DocBookDiagnostic>();
        var blocks = new List<OfficeDocumentModelBlock>();
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
        XDocumentType? documentType = _xml.DocumentType;
        if (documentType != null && (!string.IsNullOrWhiteSpace(documentType.InternalSubset) ||
            (Profile == DocBookProfile.DocBook45 &&
             (documentType.PublicId != DocBookSchemaProfiles.DocBook45.DtdPublicId || documentType.SystemId != DocBookSchemaProfiles.DocBook45.DtdSystemId)))) {
            diagnostics.Add(new DocBookDiagnostic("DB107", DocBookDiagnosticSeverity.Warning,
                "A custom document type or internal subset remains native but is not represented by the shared document model."));
        }

        OfficeDocumentModelNode Convert(XElement element, int level, string parentPath) {
            DocBookNodeKind kind = DocBookNames.GetKind(element.Name, Namespace);
            string normalizedKind = kind == DocBookNodeKind.Unknown ? "extension:" + element.Name : ToModelKind(kind);
            string text = GetPrimaryText(element, kind);
            string path = kind == DocBookNodeKind.Section
                ? (string.IsNullOrEmpty(parentPath) ? text : parentPath + " / " + text) : parentPath;
            var attributes = element.Attributes().ToDictionary(a => a.Name.ToString(), a => a.Value, StringComparer.Ordinal);
            int nodeIndex = index++;
            if (kind == DocBookNodeKind.Unknown) {
                diagnostics.Add(new DocBookDiagnostic("DB100", DocBookDiagnosticSeverity.Info,
                    $"Extension element '{element.Name}' was represented as a generic shared-model node.", path));
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
                kind == DocBookNodeKind.Link || kind == DocBookNodeKind.Entry || kind == DocBookNodeKind.Caption || kind == DocBookNodeKind.Author;
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

        var model = new OfficeDocumentModel {
            Format = OfficeDocumentFormat.DocBook,
            Source = new OfficeDocumentModelSource { Path = sourcePath, Title = Title },
            CapabilitiesUsed = new[] { "docbook.common-structure", "docbook.extensions", Profile == DocBookProfile.DocBook45 ? "docbook.4.5" : "docbook.5.2" },
            Metadata = new[] {
                new OfficeDocumentModelMetadataEntry {
                    Id = "docbook-profile", Category = "docbook", Name = "profile",
                    Value = Profile == DocBookProfile.DocBook45 ? "4.5" : "5.2", ValueType = "string"
                },
                new OfficeDocumentModelMetadataEntry {
                    Id = "docbook-kind", Category = "docbook", Name = "kind",
                    Value = Kind == DocBookDocumentKind.Book ? "book" : "article", ValueType = "string"
                }
            },
            Structure = RootElement.Elements().Select(child => Convert(child, 1, string.Empty)).ToArray(),
            Blocks = blocks
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
        DocBookDocumentKind inferredKind = sourceKind == "book" ? DocBookDocumentKind.Book : DocBookDocumentKind.Article;
        DocBookProfile inferredProfile = sourceProfile == "4.5" ? DocBookProfile.DocBook45 : DocBookProfile.DocBook52;
        DocBookDocumentKind selectedKind = kind ?? inferredKind;
        DocBookProfile selectedProfile = profile ?? inferredProfile;
        DocBookDocument document = selectedKind == DocBookDocumentKind.Article ? CreateArticle(selectedProfile) : CreateBook(selectedProfile);
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
            } else if (TryMapKind(source.Kind, out DocBookNodeKind nodeKind)) {
                string? directText = NodeAcceptsDirectText(nodeKind) && source.Children.Count == 0 ? source.Text : null;
                target = nodeKind == DocBookNodeKind.Link && selectedProfile == DocBookProfile.DocBook45 && source.Attributes.ContainsKey("url")
                    ? parent.AddRaw("ulink", directText)
                    : parent.Add(nodeKind, directText);
            } else {
                target = parent.Add(DocBookNodeKind.Paragraph, source.Text);
                target.SetAttribute("role", "officeimo-" + SanitizeRole(source.Kind));
                diagnostics.Add(new DocBookDiagnostic("DB101", DocBookDiagnosticSeverity.Warning,
                    $"Shared node kind '{source.Kind}' was represented as a role-qualified paragraph.", source.Location.HeadingPath));
            }
            foreach (KeyValuePair<string, string> attribute in source.Attributes) {
                try { target.SetAttribute(XName.Get(attribute.Key), attribute.Value); } catch (Exception exception) when (exception is ArgumentException || exception is System.Xml.XmlException) {
                    diagnostics.Add(new DocBookDiagnostic("DB102", DocBookDiagnosticSeverity.Warning,
                        $"Attribute name '{attribute.Key}' could not be represented.", source.Location.HeadingPath));
                }
            }
            foreach (OfficeDocumentModelNode child in source.Children) Add(child, target);
        }

        if (model.Structure.Count > 0) {
            bool hasMetadata = model.Structure.Any(node => string.Equals(node.Kind, "metadata", StringComparison.OrdinalIgnoreCase));
            if (!hasMetadata) document.Title = model.Source.Title;
            foreach (OfficeDocumentModelNode node in model.Structure) Add(node, document.Root);
        } else {
            document.Title = model.Source.Title;
            diagnostics.Add(new DocBookDiagnostic("DB103", DocBookDiagnosticSeverity.Warning,
                "The shared model had no recursive Structure; flat Blocks were emitted as paragraphs."));
            foreach (OfficeDocumentModelBlock block in model.Blocks) document.AddParagraph(block.Text);
        }
        return new DocBookConversionResult<DocBookDocument>(document, diagnostics);
    }

    private static string GetPrimaryText(XElement element, DocBookNodeKind kind) {
        if (kind == DocBookNodeKind.Section || kind == DocBookNodeKind.Table || kind == DocBookNodeKind.Figure || kind == DocBookNodeKind.Info) {
            return element.Elements().FirstOrDefault(e => e.Name.LocalName == "title")?.Value ?? string.Empty;
        }
        if (kind == DocBookNodeKind.Title || kind == DocBookNodeKind.Subtitle || kind == DocBookNodeKind.Link ||
            kind == DocBookNodeKind.Entry || kind == DocBookNodeKind.Caption || kind == DocBookNodeKind.Author) return element.Value;
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
        kind == DocBookNodeKind.ProgramListing || kind == DocBookNodeKind.Screen || kind == DocBookNodeKind.Entry || kind == DocBookNodeKind.Link;

    private static string SanitizeRole(string value) => new string((value ?? "unknown").Select(c => char.IsLetterOrDigit(c) || c == '-' ? c : '-').ToArray());
}
