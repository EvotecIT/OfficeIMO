using OfficeIMO;
using OfficeIMO.Core.Internal;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml.Linq;

namespace OfficeIMO.Opml;

public sealed partial class OpmlDocument {
    /// <summary>Converts this document to the shared recursive document model.</summary>
    public OpmlConversionResult<OfficeDocumentModel> ToOfficeDocumentModel(string? sourcePath = null) =>
        ToOfficeDocumentModel(sourcePath, null, default);

    /// <summary>Converts this document to the shared recursive document model with a source path and bounded diagnostic budget.</summary>
    public OpmlConversionResult<OfficeDocumentModel> ToOfficeDocumentModel(
        string? sourcePath,
        OpmlConversionOptions? options) => ToOfficeDocumentModel(sourcePath, options, default);

    /// <summary>Converts this document to the shared recursive document model with bounded diagnostics and cancellation.</summary>
    public OpmlConversionResult<OfficeDocumentModel> ToOfficeDocumentModel(
        string? sourcePath,
        OpmlConversionOptions? options,
        CancellationToken cancellationToken) {
        options ??= new OpmlConversionOptions();
        options.Validate();
        cancellationToken.ThrowIfCancellationRequested();
        var diagnostics = new OpmlDiagnosticCollector(options.MaxDetailedDiagnosticsPerCode);
        var blocks = new List<OfficeDocumentModelBlock>();
        var links = new List<OfficeDocumentModelLink>();
        int blockIndex = 0;

        OfficeDocumentModelNode Convert(OpmlOutline outline, int level, string parentPath) {
            cancellationToken.ThrowIfCancellationRequested();
            if (level > options.MaxStructureDepth) {
                throw new InvalidDataException($"The OPML shared-model projection exceeds MaxStructureDepth ({options.MaxStructureDepth}).");
            }
            if (blockIndex >= options.MaxStructureNodes) {
                throw new InvalidDataException($"The OPML shared-model projection exceeds MaxStructureNodes ({options.MaxStructureNodes}).");
            }
            string headingPath = OfficeDocumentHeadingPath.Append(parentPath, outline.Text, " / ");
            var attributes = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (XAttribute attribute in outline.Element.Attributes()) {
                cancellationToken.ThrowIfCancellationRequested();
                attributes.Add(attribute.Name.ToString(), attribute.Value);
            }
            int nodeIndex = blockIndex++;
            blocks.Add(new OfficeDocumentModelBlock {
                Id = "outline-" + nodeIndex,
                Kind = "outline",
                Text = outline.Text,
                Level = level,
                Location = new OfficeDocumentModelLocation { Path = sourcePath, BlockIndex = nodeIndex, HeadingPath = headingPath }
            });
            AddLink(outline.Url, "url");
            AddLink(outline.XmlUrl, "subscription");
            AddLink(outline.HtmlUrl, "html");
            if (HasExtensionElement(outline.Element)) {
                diagnostics.Add(new OpmlDiagnostic("OPML200", OpmlDiagnosticSeverity.Warning,
                    "An outline extension element remains in native OPML but is not represented by the shared outline model.", headingPath));
            }
            var children = new List<OfficeDocumentModelNode>();
            foreach (XElement child in outline.Element.Elements("outline")) {
                cancellationToken.ThrowIfCancellationRequested();
                children.Add(Convert(new OpmlOutline(this, child), level + 1, headingPath));
            }
            return new OfficeDocumentModelNode {
                Id = "outline-" + nodeIndex,
                Kind = "outline",
                Text = outline.Text,
                Level = level,
                Attributes = attributes,
                Children = children,
                Location = new OfficeDocumentModelLocation { Path = sourcePath, HeadingPath = headingPath }
            };

            void AddLink(string? uri, string kind) {
                if (string.IsNullOrWhiteSpace(uri)) return;
                links.Add(new OfficeDocumentModelLink {
                    Id = "opml-link-" + nodeIndex + "-" + kind,
                    Kind = kind,
                    Uri = uri,
                    Text = outline.Text,
                    Location = new OfficeDocumentModelLocation { Path = sourcePath, BlockIndex = nodeIndex, HeadingPath = headingPath }
                });
            }
        }

        bool HasExtensionElement(XElement outlineElement) {
            foreach (XElement element in outlineElement.Elements()) {
                cancellationToken.ThrowIfCancellationRequested();
                if (element.Name != "outline") return true;
            }
            return false;
        }

        IReadOnlyList<OfficeDocumentModelNode> ConvertRoots() {
            var roots = new List<OfficeDocumentModelNode>();
            foreach (XElement outline in BodyElement.Elements("outline")) {
                cancellationToken.ThrowIfCancellationRequested();
                roots.Add(Convert(new OpmlOutline(this, outline), 1, string.Empty));
            }
            return roots;
        }

        var model = new OfficeDocumentModel {
            Format = OfficeDocumentFormat.Opml,
            Source = new OfficeDocumentModelSource { Path = sourcePath, Title = Head.Title, Author = Head.OwnerName },
            CapabilitiesUsed = new[] { "opml.outlines", "opml.attributes", "opml.nesting" },
            Metadata = BuildMetadata(diagnostics, cancellationToken),
            Structure = ConvertRoots(),
            Blocks = blocks,
            Links = links
        };
        return new OpmlConversionResult<OfficeDocumentModel>(model, diagnostics.ToArray());
    }

    /// <summary>Creates OPML from the recursive portion of the shared document model.</summary>
    public static OpmlConversionResult<OpmlDocument> FromOfficeDocumentModel(
        OfficeDocumentModel model,
        OpmlVersion? version = null) => FromOfficeDocumentModel(model, version, null);

    /// <summary>Creates OPML from the recursive portion of the shared document model with an explicit profile and bounded diagnostic budget.</summary>
    public static OpmlConversionResult<OpmlDocument> FromOfficeDocumentModel(
        OfficeDocumentModel model,
        OpmlVersion? version,
        OpmlConversionOptions? options) {
        if (model == null) throw new ArgumentNullException(nameof(model));
        options ??= new OpmlConversionOptions();
        options.Validate();
        IReadOnlyList<OfficeDocumentModelNode> structureNodes = OfficeDocumentModelStructureTraversal.ValidateAndFlatten(
            model.Structure, options.MaxStructureDepth, options.MaxStructureNodes);
        var diagnostics = new OpmlDiagnosticCollector(options.MaxDetailedDiagnosticsPerCode);
        string? sourceVersion = model.Metadata.FirstOrDefault(entry => entry.Category == "opml" && entry.Name == "version")?.Value;
        bool sourceVersionIsSupported = sourceVersion == null || sourceVersion == "1.0" || sourceVersion == "1.1" || sourceVersion == "2.0";
        OpmlVersion inferredVersion = sourceVersion == "1.0" || sourceVersion == "1.1" ? OpmlVersion.Opml10 : OpmlVersion.Opml20;
        OpmlVersion selectedVersion = version ?? inferredVersion;
        OpmlDocument document = Create(selectedVersion);
        if (version == null && sourceVersion == "1.1") document.DeclaredVersion = "1.1";
        if (!sourceVersionIsSupported) {
            diagnostics.Add(new OpmlDiagnostic("OPML105", OpmlDiagnosticSeverity.Warning,
                $"The unsupported source OPML version '{sourceVersion}' was normalized to '{document.DeclaredVersion}'."));
        }
        if (version.HasValue && sourceVersion != null && selectedVersion != inferredVersion) {
            diagnostics.Add(new OpmlDiagnostic("OPML103", OpmlDiagnosticSeverity.Warning,
                $"The source OPML profile '{sourceVersion}' was changed to '{document.DeclaredVersion}' by the requested conversion profile."));
        }
        OfficeDocumentModelMetadataEntry? titleMetadata = model.Metadata.FirstOrDefault(entry =>
            entry.Category == "opml.head" && entry.Name == "title");
        OfficeDocumentModelMetadataEntry? ownerMetadata = model.Metadata.FirstOrDefault(entry =>
            entry.Category == "opml.head" && entry.Name == "ownerName");
        document.Head.Title = model.Source.Title ?? (titleMetadata == null ? null : titleMetadata.Value ?? string.Empty);
        document.Head.OwnerName = model.Source.Author ?? (ownerMetadata == null ? null : ownerMetadata.Value ?? string.Empty);
        ReportPrimaryMetadataConflict("title", model.Source.Title, titleMetadata);
        ReportPrimaryMetadataConflict("ownerName", model.Source.Author, ownerMetadata);
        ApplyPrimaryMetadataAttributes("title", titleMetadata);
        ApplyPrimaryMetadataAttributes("ownerName", ownerMetadata);

        void ReportPrimaryMetadataConflict(
            string elementName,
            string? sourceValue,
            OfficeDocumentModelMetadataEntry? metadata) {
            if (sourceValue == null || metadata == null || string.Equals(sourceValue, metadata.Value, StringComparison.Ordinal)) return;
            diagnostics.Add(new OpmlDiagnostic("OPML110", OpmlDiagnosticSeverity.Warning,
                $"Shared Source and opml.head/{elementName} metadata contain conflicting values; the Source value took precedence.",
                "/opml/head/" + elementName));
        }

        void ApplyPrimaryMetadataAttributes(string elementName, OfficeDocumentModelMetadataEntry? metadata) {
            if (metadata == null) return;
            XElement? element = document.HeadElement.Element(elementName);
            if (element != null) {
                foreach (KeyValuePair<string, string> attribute in metadata.Attributes) {
                    try {
                        element.SetAttributeValue(XName.Get(attribute.Key), attribute.Value);
                    } catch (Exception exception) when (exception is ArgumentException || exception is System.Xml.XmlException) {
                        diagnostics.Add(new OpmlDiagnostic("OPML102", OpmlDiagnosticSeverity.Warning,
                            $"Head metadata attribute '{attribute.Key}' could not be represented in OPML."));
                    }
                }
            }
        }
        bool primaryTitleConsumed = false;
        bool primaryOwnerConsumed = false;
        foreach (OfficeDocumentModelMetadataEntry metadata in model.Metadata.Where(entry => entry.Category == "opml.head")) {
            if (metadata.Name == "title" && !primaryTitleConsumed) {
                primaryTitleConsumed = true;
                continue;
            }
            if (metadata.Name == "ownerName" && !primaryOwnerConsumed) {
                primaryOwnerConsumed = true;
                continue;
            }
            try {
                XName name = XName.Get(metadata.Name);
                var element = new XElement(name, metadata.Value ?? string.Empty);
                foreach (KeyValuePair<string, string> attribute in metadata.Attributes) {
                    element.SetAttributeValue(XName.Get(attribute.Key), attribute.Value);
                }
                document.HeadElement.Add(element);
            } catch (Exception exception) when (exception is ArgumentException || exception is System.Xml.XmlException) {
                diagnostics.Add(new OpmlDiagnostic("OPML102", OpmlDiagnosticSeverity.Warning,
                    $"Head metadata '{metadata.Name}' could not be represented in OPML."));
            }
        }
        OfficeDocumentModelMetadataEntry? versionMetadata = model.Metadata.FirstOrDefault(entry =>
            entry.Category == "opml" && entry.Name == "version");
        foreach (OfficeDocumentModelMetadataEntry metadata in model.Metadata) {
            if (ReferenceEquals(metadata, versionMetadata) || metadata.Category == "opml.head") continue;
            diagnostics.Add(new OpmlDiagnostic("OPML109", OpmlDiagnosticSeverity.Warning,
                $"Shared metadata '{metadata.Category}/{metadata.Name}' cannot be represented in OPML and was omitted."));
        }

        void Add(OfficeDocumentModelNode node, OpmlOutline? parent) {
            if (!string.Equals(node.Kind, "outline", StringComparison.OrdinalIgnoreCase)) {
                diagnostics.Add(new OpmlDiagnostic("OPML104", OpmlDiagnosticSeverity.Warning,
                    $"Shared node kind '{node.Kind}' was normalized to an OPML outline.", node.Location.HeadingPath));
            }
            string outlineText = ResolveOutlineText(node);
            OpmlOutline outline = parent == null ? document.AddOutline(outlineText) : parent.AddChild(outlineText);
            foreach (KeyValuePair<string, string> attribute in node.Attributes) {
                try { outline.SetAttribute(XName.Get(attribute.Key), attribute.Value); } catch (Exception exception) when (exception is ArgumentException || exception is System.Xml.XmlException) {
                    diagnostics.Add(new OpmlDiagnostic("OPML100", OpmlDiagnosticSeverity.Warning,
                        $"Attribute name '{attribute.Key}' could not be represented in OPML.", node.Location.HeadingPath));
                }
            }
            outline.Text = outlineText;
            foreach (OfficeDocumentModelNode child in node.Children) Add(child, outline);
        }

        string ResolveOutlineText(OfficeDocumentModelNode node) {
            if (!node.Attributes.TryGetValue("text", out string? attributeText) ||
                string.Equals(attributeText, node.Text, StringComparison.Ordinal)) return node.Text;
            OfficeDocumentModelBlock? projectedBlock = model.Blocks.FirstOrDefault(block =>
                string.Equals(block.Id, node.Id, StringComparison.Ordinal) &&
                string.Equals(block.Kind, node.Kind, StringComparison.OrdinalIgnoreCase) &&
                block.Level == node.Level);
            bool attributeWins = projectedBlock != null &&
                string.Equals(projectedBlock.Text, node.Text, StringComparison.Ordinal) &&
                !string.Equals(projectedBlock.Text, attributeText, StringComparison.Ordinal);
            diagnostics.Add(new OpmlDiagnostic("OPML110", OpmlDiagnosticSeverity.Warning,
                attributeWins
                    ? "The outline text attribute differed from its unchanged primary text projection; the edited attribute took precedence."
                    : "The outline text attribute and primary text contain conflicting values; the primary text took precedence.",
                node.Location.HeadingPath));
            return attributeWins ? attributeText : node.Text;
        }

        bool AddFlatLink(OfficeDocumentModelLink link) {
            string text = link.Text ?? link.Uri ?? link.DestinationName ?? link.Id;
            bool represented = false;
            if (!string.IsNullOrWhiteSpace(link.Uri)) {
                OpmlOutline outline = document.AddOutline(text);
                if (string.Equals(link.Kind, "subscription", StringComparison.OrdinalIgnoreCase)) {
                    outline.Type = "rss";
                    outline.XmlUrl = link.Uri;
                } else if (string.Equals(link.Kind, "html", StringComparison.OrdinalIgnoreCase)) {
                    outline.HtmlUrl = link.Uri;
                } else {
                    outline.Type = "link";
                    outline.Url = link.Uri;
                }
                represented = true;
            }
            bool hasUnsupportedTarget = !string.IsNullOrWhiteSpace(link.DestinationName) || link.DestinationPageNumber.HasValue ||
                !string.IsNullOrWhiteSpace(link.DestinationMode) || !string.IsNullOrWhiteSpace(link.NamedAction) ||
                !string.IsNullOrWhiteSpace(link.RemoteFile) || !string.IsNullOrWhiteSpace(link.RemoteDestinationName) ||
                link.RemoteDestinationPageNumber.HasValue;
            if (!represented || hasUnsupportedTarget) {
                diagnostics.Add(new OpmlDiagnostic("OPML106", OpmlDiagnosticSeverity.Warning,
                    represented
                        ? $"Shared link '{link.Id}' was emitted, but one or more additional target fields could not be represented in OPML."
                        : $"Shared link '{link.Id}' had no OPML-representable URI.",
                    link.Location?.HeadingPath));
            }
            return represented;
        }

        foreach (OfficeDocumentModelTable table in model.Tables) {
            diagnostics.Add(new OpmlDiagnostic("OPML108", OpmlDiagnosticSeverity.Warning,
                $"Shared table '{table.Title ?? table.Kind ?? "unnamed"}' cannot be represented in OPML and was omitted.",
                table.Location?.HeadingPath));
        }
        foreach (OfficeDocumentModelAsset asset in model.Assets) {
            diagnostics.Add(new OpmlDiagnostic("OPML108", OpmlDiagnosticSeverity.Warning,
                $"Shared asset '{asset.Id}' cannot be represented in OPML and was omitted.",
                asset.Location?.HeadingPath));
        }
        foreach (OfficeDocumentModelPage page in model.Pages) {
            diagnostics.Add(new OpmlDiagnostic("OPML108", OpmlDiagnosticSeverity.Warning,
                $"Shared page '{page.Name ?? (page.Number.HasValue ? page.Number.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) : "unnamed")}' cannot be represented in OPML and was omitted.",
                page.Location?.HeadingPath));
        }
        foreach (OfficeDocumentModelFormField form in model.Forms) {
            diagnostics.Add(new OpmlDiagnostic("OPML108", OpmlDiagnosticSeverity.Warning,
                $"Shared form field '{form.Id}' cannot be represented in OPML and was omitted.",
                form.Location?.HeadingPath));
        }
        foreach (OfficeDocumentModelVisual visual in model.Visuals) {
            diagnostics.Add(new OpmlDiagnostic("OPML108", OpmlDiagnosticSeverity.Warning,
                $"Shared visual '{visual.SourceName ?? visual.Kind}' cannot be represented in OPML and was omitted.",
                visual.Location?.HeadingPath));
        }

        if (model.Structure.Count > 0) {
            foreach (OfficeDocumentModelNode node in model.Structure) Add(node, null);
            foreach (OfficeDocumentModelBlock block in model.Blocks.Where(block => !IsDerivedBlock(block, structureNodes))) {
                document.AddOutline(block.Text);
                diagnostics.Add(new OpmlDiagnostic("OPML107", OpmlDiagnosticSeverity.Warning,
                    $"Supplementary shared block '{block.Id}' was appended as a top-level outline because it was not represented by recursive Structure.",
                    block.Location?.HeadingPath));
            }
            foreach (OfficeDocumentModelLink link in model.Links.Where(link => !IsDerivedLink(link, structureNodes))) {
                if (AddFlatLink(link)) {
                    diagnostics.Add(new OpmlDiagnostic("OPML107", OpmlDiagnosticSeverity.Warning,
                        $"Supplementary shared link '{link.Id}' was appended as a top-level outline because it was not represented by recursive Structure.",
                        link.Location?.HeadingPath));
                }
            }
        } else {
            if (model.Blocks.Count > 0 || model.Links.Count > 0) {
                diagnostics.Add(new OpmlDiagnostic("OPML101", OpmlDiagnosticSeverity.Warning,
                    "The shared model had no recursive Structure; flat Blocks and Links were emitted as top-level outlines."));
            }
            foreach (OfficeDocumentModelBlock block in model.Blocks) document.AddOutline(block.Text);
            foreach (OfficeDocumentModelLink link in model.Links) AddFlatLink(link);
        }
        foreach (OpmlDiagnostic diagnostic in document.Validate().Diagnostics.Where(candidate =>
                     candidate.Severity == OpmlDiagnosticSeverity.Error)) {
            diagnostics.Add(diagnostic);
        }
        return new OpmlConversionResult<OpmlDocument>(document, diagnostics.ToArray());

        static bool IsDerivedBlock(OfficeDocumentModelBlock block, IEnumerable<OfficeDocumentModelNode> nodes) =>
            !string.IsNullOrEmpty(block.Id) && block.Marker == null && block.Region == null && nodes.Any(node =>
                string.Equals(node.Id, block.Id, StringComparison.Ordinal) &&
                string.Equals(node.Kind, "outline", StringComparison.OrdinalIgnoreCase) &&
                string.Equals(block.Kind, "outline", StringComparison.OrdinalIgnoreCase) &&
                string.Equals(node.Text, block.Text, StringComparison.Ordinal) && node.Level == block.Level);

        static bool IsDerivedLink(OfficeDocumentModelLink link, IEnumerable<OfficeDocumentModelNode> nodes) {
            const string prefix = "opml-link-";
            if (string.IsNullOrEmpty(link.Id) || !link.Id.StartsWith(prefix, StringComparison.Ordinal) || string.IsNullOrEmpty(link.Uri) ||
                !string.IsNullOrWhiteSpace(link.DestinationName) || link.DestinationPageNumber.HasValue ||
                !string.IsNullOrWhiteSpace(link.DestinationMode) || !string.IsNullOrWhiteSpace(link.NamedAction) ||
                !string.IsNullOrWhiteSpace(link.RemoteFile) || !string.IsNullOrWhiteSpace(link.RemoteDestinationName) ||
                link.RemoteDestinationPageNumber.HasValue) return false;
            int kindSeparator = link.Id.LastIndexOf('-');
            if (kindSeparator <= prefix.Length) return false;
            string nodeId = "outline-" + link.Id.Substring(prefix.Length, kindSeparator - prefix.Length);
            string kind = link.Id.Substring(kindSeparator + 1);
            string? attributeName = string.Equals(kind, "url", StringComparison.Ordinal) ? "url"
                : string.Equals(kind, "subscription", StringComparison.Ordinal) ? "xmlUrl"
                : string.Equals(kind, "html", StringComparison.Ordinal) ? "htmlUrl" : null;
            if (attributeName == null || !string.Equals(kind, link.Kind, StringComparison.OrdinalIgnoreCase)) return false;
            return nodes.Any(node => string.Equals(node.Id, nodeId, StringComparison.Ordinal) &&
                (link.Text == null || string.Equals(node.Text, link.Text, StringComparison.Ordinal)) &&
                node.Attributes.TryGetValue(attributeName, out string? value) && string.Equals(value, link.Uri, StringComparison.Ordinal));
        }
    }

    private IReadOnlyList<OfficeDocumentModelMetadataEntry> BuildMetadata(
        OpmlDiagnosticCollector diagnostics,
        CancellationToken cancellationToken) {
        var values = new List<OfficeDocumentModelMetadataEntry> {
            new OfficeDocumentModelMetadataEntry {
                Id = "opml-version", Category = "opml", Name = "version", Value = DeclaredVersion, ValueType = "string"
            }
        };
        int index = 0;
        foreach (XElement element in HeadElement.Elements()) {
            cancellationToken.ThrowIfCancellationRequested();
            if (element.HasElements) {
                diagnostics.Add(new OpmlDiagnostic("OPML201", OpmlDiagnosticSeverity.Warning,
                    $"Head extension element '{element.Name}' contains nested XML that is not represented by shared metadata.", "/opml/head"));
            }
            var attributes = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (XAttribute attribute in element.Attributes()) {
                cancellationToken.ThrowIfCancellationRequested();
                attributes.Add(attribute.Name.ToString(), attribute.Value);
            }
            values.Add(new OfficeDocumentModelMetadataEntry {
                Id = "opml-head-" + index++,
                Category = "opml.head",
                Name = element.Name.ToString(),
                Value = element.Value,
                ValueType = "string",
                Attributes = attributes
            });
        }
        if (AnyWithCancellation(Root.Attributes(), attribute => !attribute.IsNamespaceDeclaration && attribute.Name != "version")) {
            diagnostics.Add(new OpmlDiagnostic("OPML202", OpmlDiagnosticSeverity.Warning,
                "OPML root extension attributes remain native but are not represented by the shared document model.", "/opml"));
        }
        if (AnyWithCancellation(HeadElement.Attributes(), attribute => !attribute.IsNamespaceDeclaration)) {
            diagnostics.Add(new OpmlDiagnostic("OPML208", OpmlDiagnosticSeverity.Warning,
                "OPML head extension attributes remain native but are not represented by shared metadata.", "/opml/head"));
        }
        if (AnyWithCancellation(BodyElement.Attributes(), attribute => !attribute.IsNamespaceDeclaration)) {
            diagnostics.Add(new OpmlDiagnostic("OPML209", OpmlDiagnosticSeverity.Warning,
                "OPML body extension attributes remain native but are not represented by the shared outline model.", "/opml/body"));
        }
        if (AnyWithCancellation(BodyElement.Elements(), element => element.Name != "outline")) {
            diagnostics.Add(new OpmlDiagnostic("OPML203", OpmlDiagnosticSeverity.Warning,
                "OPML body extension elements remain native but are not represented by the shared outline model.", "/opml/body"));
        }
        if (AnyWithCancellation(Root.Elements(), element => element.Name != "head" && element.Name != "body")) {
            diagnostics.Add(new OpmlDiagnostic("OPML205", OpmlDiagnosticSeverity.Warning,
                "OPML root extension elements remain native but are not represented by the shared document model.", "/opml"));
        }
        bool hasUnrepresentedText = HasSignificantText(Root.Nodes().OfType<XText>()) ||
            HasSignificantText(HeadElement.Nodes().OfType<XText>()) ||
            HasSignificantText(BodyElement.Nodes().OfType<XText>()) || HasSignificantOutlineText();
        if (hasUnrepresentedText) {
            diagnostics.Add(new OpmlDiagnostic("OPML206", OpmlDiagnosticSeverity.Warning,
                "Significant element text remains native but is not represented by the shared outline model."));
        }
        if (HasOutlineExtensionElements()) {
            diagnostics.Add(new OpmlDiagnostic("OPML207", OpmlDiagnosticSeverity.Warning,
                "Outline extension elements remain native but are not represented by the shared outline model.", "/opml/body"));
        }
        if (HasUnrepresentedMarkup()) {
            diagnostics.Add(new OpmlDiagnostic("OPML204", OpmlDiagnosticSeverity.Warning,
                "Comments and processing instructions remain native but are not represented by the shared document model."));
        }
        return values;

        bool HasSignificantText(IEnumerable<XText> nodes) {
            foreach (XText text in nodes) {
                cancellationToken.ThrowIfCancellationRequested();
                if (!string.IsNullOrWhiteSpace(text.Value)) return true;
            }
            return false;
        }

        bool HasSignificantOutlineText() {
            foreach (XElement outline in BodyElement.Descendants("outline")) {
                cancellationToken.ThrowIfCancellationRequested();
                if (HasSignificantText(outline.Nodes().OfType<XText>())) return true;
            }
            return false;
        }

        bool HasOutlineExtensionElements() {
            foreach (XElement outline in BodyElement.Descendants("outline")) {
                cancellationToken.ThrowIfCancellationRequested();
                if (AnyWithCancellation(outline.Elements(), element => element.Name != "outline")) return true;
            }
            return false;
        }

        bool AnyWithCancellation<T>(IEnumerable<T> items, Func<T, bool> predicate) {
            foreach (T item in items) {
                cancellationToken.ThrowIfCancellationRequested();
                if (predicate(item)) return true;
            }
            return false;
        }

        bool HasUnrepresentedMarkup() {
            foreach (XNode node in _xml.DescendantNodes()) {
                cancellationToken.ThrowIfCancellationRequested();
                if (node is XComment || node is XProcessingInstruction) return true;
            }
            return false;
        }
    }
}
