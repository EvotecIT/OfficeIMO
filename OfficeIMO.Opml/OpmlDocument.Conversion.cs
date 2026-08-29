using OfficeIMO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Opml;

public sealed partial class OpmlDocument {
    /// <summary>Converts this document to the shared recursive document model.</summary>
    public OpmlConversionResult<OfficeDocumentModel> ToOfficeDocumentModel(string? sourcePath = null) {
        var diagnostics = new List<OpmlDiagnostic>();
        var blocks = new List<OfficeDocumentModelBlock>();
        int blockIndex = 0;

        OfficeDocumentModelNode Convert(OpmlOutline outline, int level, string parentPath) {
            string headingPath = string.IsNullOrEmpty(parentPath) ? outline.Text : parentPath + " / " + outline.Text;
            var attributes = outline.Attributes.ToDictionary(
                pair => pair.Key.ToString(), pair => pair.Value, StringComparer.Ordinal);
            int nodeIndex = blockIndex++;
            blocks.Add(new OfficeDocumentModelBlock {
                Id = "outline-" + nodeIndex,
                Kind = "outline",
                Text = outline.Text,
                Level = level,
                Location = new OfficeDocumentModelLocation { Path = sourcePath, BlockIndex = nodeIndex, HeadingPath = headingPath }
            });
            if (outline.Element.Elements().Any(element => element.Name != "outline")) {
                diagnostics.Add(new OpmlDiagnostic("OPML200", OpmlDiagnosticSeverity.Warning,
                    "An outline extension element remains in native OPML but is not represented by the shared outline model.", headingPath));
            }
            return new OfficeDocumentModelNode {
                Id = "outline-" + nodeIndex,
                Kind = "outline",
                Text = outline.Text,
                Level = level,
                Attributes = attributes,
                Children = outline.Children.Select(child => Convert(child, level + 1, headingPath)).ToArray(),
                Location = new OfficeDocumentModelLocation { Path = sourcePath, HeadingPath = headingPath }
            };
        }

        var model = new OfficeDocumentModel {
            Format = OfficeDocumentFormat.Opml,
            Source = new OfficeDocumentModelSource { Path = sourcePath, Title = Head.Title },
            CapabilitiesUsed = new[] { "opml.outlines", "opml.attributes", "opml.nesting" },
            Metadata = BuildMetadata(diagnostics),
            Structure = Outlines.Select(outline => Convert(outline, 1, string.Empty)).ToArray(),
            Blocks = blocks
        };
        return new OpmlConversionResult<OfficeDocumentModel>(model, diagnostics);
    }

    /// <summary>Creates OPML from the recursive portion of the shared document model.</summary>
    public static OpmlConversionResult<OpmlDocument> FromOfficeDocumentModel(
        OfficeDocumentModel model,
        OpmlVersion? version = null) {
        if (model == null) throw new ArgumentNullException(nameof(model));
        var diagnostics = new List<OpmlDiagnostic>();
        string? sourceVersion = model.Metadata.FirstOrDefault(entry => entry.Category == "opml" && entry.Name == "version")?.Value;
        OpmlVersion inferredVersion = sourceVersion == "1.0" || sourceVersion == "1.1" ? OpmlVersion.Opml10 : OpmlVersion.Opml20;
        OpmlVersion selectedVersion = version ?? inferredVersion;
        OpmlDocument document = Create(selectedVersion);
        if (version == null && sourceVersion == "1.1") document.DeclaredVersion = "1.1";
        if (version.HasValue && sourceVersion != null && selectedVersion != inferredVersion) {
            diagnostics.Add(new OpmlDiagnostic("OPML103", OpmlDiagnosticSeverity.Warning,
                $"The source OPML profile '{sourceVersion}' was changed to '{document.DeclaredVersion}' by the requested conversion profile."));
        }
        OfficeDocumentModelMetadataEntry? titleMetadata = model.Metadata.FirstOrDefault(entry =>
            entry.Category == "opml.head" && entry.Name == "title");
        document.Head.Title = model.Source.Title ?? (titleMetadata == null ? null : titleMetadata.Value ?? string.Empty);
        if (titleMetadata != null) {
            XElement? titleElement = document.HeadElement.Element("title");
            if (titleElement != null) {
                foreach (KeyValuePair<string, string> attribute in titleMetadata.Attributes) {
                    try {
                        titleElement.SetAttributeValue(XName.Get(attribute.Key), attribute.Value);
                    } catch (Exception exception) when (exception is ArgumentException || exception is System.Xml.XmlException) {
                        diagnostics.Add(new OpmlDiagnostic("OPML102", OpmlDiagnosticSeverity.Warning,
                            $"Title attribute '{attribute.Key}' could not be represented in OPML."));
                    }
                }
            }
        }
        bool primaryTitleConsumed = false;
        foreach (OfficeDocumentModelMetadataEntry metadata in model.Metadata.Where(entry => entry.Category == "opml.head")) {
            if (metadata.Name == "title" && !primaryTitleConsumed) {
                primaryTitleConsumed = true;
                continue;
            }
            if (metadata.Value == null) continue;
            try {
                XName name = XName.Get(metadata.Name);
                var element = new XElement(name, metadata.Value);
                foreach (KeyValuePair<string, string> attribute in metadata.Attributes) {
                    element.SetAttributeValue(XName.Get(attribute.Key), attribute.Value);
                }
                document.HeadElement.Add(element);
            } catch (Exception exception) when (exception is ArgumentException || exception is System.Xml.XmlException) {
                diagnostics.Add(new OpmlDiagnostic("OPML102", OpmlDiagnosticSeverity.Warning,
                    $"Head metadata '{metadata.Name}' could not be represented in OPML."));
            }
        }

        void Add(OfficeDocumentModelNode node, OpmlOutline? parent) {
            OpmlOutline outline = parent == null ? document.AddOutline(node.Text) : parent.AddChild(node.Text);
            foreach (KeyValuePair<string, string> attribute in node.Attributes) {
                try { outline.SetAttribute(XName.Get(attribute.Key), attribute.Value); } catch (Exception exception) when (exception is ArgumentException || exception is System.Xml.XmlException) {
                    diagnostics.Add(new OpmlDiagnostic("OPML100", OpmlDiagnosticSeverity.Warning,
                        $"Attribute name '{attribute.Key}' could not be represented in OPML.", node.Location.HeadingPath));
                }
            }
            outline.Text = node.Text;
            foreach (OfficeDocumentModelNode child in node.Children) Add(child, outline);
        }

        if (model.Structure.Count > 0) {
            foreach (OfficeDocumentModelNode node in model.Structure) Add(node, null);
        } else {
            diagnostics.Add(new OpmlDiagnostic("OPML101", OpmlDiagnosticSeverity.Warning,
                "The shared model had no recursive Structure; flat Blocks were emitted as top-level outlines."));
            foreach (OfficeDocumentModelBlock block in model.Blocks) document.AddOutline(block.Text);
        }
        return new OpmlConversionResult<OpmlDocument>(document, diagnostics);
    }

    private IReadOnlyList<OfficeDocumentModelMetadataEntry> BuildMetadata(List<OpmlDiagnostic> diagnostics) {
        var values = new List<OfficeDocumentModelMetadataEntry> {
            new OfficeDocumentModelMetadataEntry {
                Id = "opml-version", Category = "opml", Name = "version", Value = DeclaredVersion, ValueType = "string"
            }
        };
        int index = 0;
        foreach (XElement element in HeadElement.Elements()) {
            if (element.HasElements) {
                diagnostics.Add(new OpmlDiagnostic("OPML201", OpmlDiagnosticSeverity.Warning,
                    $"Head extension element '{element.Name}' contains nested XML that is not represented by shared metadata.", "/opml/head"));
            }
            values.Add(new OfficeDocumentModelMetadataEntry {
                Id = "opml-head-" + index++,
                Category = "opml.head",
                Name = element.Name.ToString(),
                Value = element.Value,
                ValueType = "string",
                Attributes = element.Attributes().ToDictionary(attribute => attribute.Name.ToString(), attribute => attribute.Value, StringComparer.Ordinal)
            });
        }
        if (Root.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration && attribute.Name != "version")) {
            diagnostics.Add(new OpmlDiagnostic("OPML202", OpmlDiagnosticSeverity.Warning,
                "OPML root extension attributes remain native but are not represented by the shared document model.", "/opml"));
        }
        if (BodyElement.Elements().Any(element => element.Name != "outline")) {
            diagnostics.Add(new OpmlDiagnostic("OPML203", OpmlDiagnosticSeverity.Warning,
                "OPML body extension elements remain native but are not represented by the shared outline model.", "/opml/body"));
        }
        if (Root.Elements().Any(element => element.Name != "head" && element.Name != "body")) {
            diagnostics.Add(new OpmlDiagnostic("OPML205", OpmlDiagnosticSeverity.Warning,
                "OPML root extension elements remain native but are not represented by the shared document model.", "/opml"));
        }
        bool hasUnrepresentedText = Root.Nodes().OfType<XText>().Any(text => !string.IsNullOrWhiteSpace(text.Value)) ||
            HeadElement.Nodes().OfType<XText>().Any(text => !string.IsNullOrWhiteSpace(text.Value)) ||
            BodyElement.Nodes().OfType<XText>().Any(text => !string.IsNullOrWhiteSpace(text.Value)) ||
            BodyElement.Descendants("outline").SelectMany(element => element.Nodes().OfType<XText>())
                .Any(text => !string.IsNullOrWhiteSpace(text.Value));
        if (hasUnrepresentedText) {
            diagnostics.Add(new OpmlDiagnostic("OPML206", OpmlDiagnosticSeverity.Warning,
                "Significant element text remains native but is not represented by the shared outline model."));
        }
        if (BodyElement.Descendants("outline").Any(outline => outline.Elements().Any(element => element.Name != "outline"))) {
            diagnostics.Add(new OpmlDiagnostic("OPML207", OpmlDiagnosticSeverity.Warning,
                "Outline extension elements remain native but are not represented by the shared outline model.", "/opml/body"));
        }
        if (_xml.DescendantNodes().Any(node => node is XComment || node is XProcessingInstruction)) {
            diagnostics.Add(new OpmlDiagnostic("OPML204", OpmlDiagnosticSeverity.Warning,
                "Comments and processing instructions remain native but are not represented by the shared document model."));
        }
        return values;
    }
}
