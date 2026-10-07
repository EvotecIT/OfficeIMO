namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    private const string SchemaVocabulary = "http://schema.org/";
    private static readonly string[] AccessibilityProperties = {
        "accessMode", "accessModeSufficient", "accessibilityFeature", "accessibilityHazard", "accessibilitySummary"
    };

    /// <summary>
    /// Atomically replaces the five publication-level discovery properties in EPUB 3. Other metadata and
    /// refined properties remain intact. Existing declarations retain attributes where their positions remain;
    /// removing an identified declaration referenced by a refinement is rejected. Vocabulary tokens are supplied
    /// by the publisher, not inferred or certified. Null summary and empty sufficient-mode sets remove those properties.
    /// </summary>
    public void SetAccessibilityMetadata(EpubAccessibilityMetadata metadata) {
        if (metadata == null) throw new ArgumentNullException(nameof(metadata));
        if (PackageVersion != "3.0") throw new NotSupportedException("Typed accessibility discovery metadata requires EPUB 3.");
        if (EpubVocabulary.Expand(Root, "schema:accessMode") != SchemaVocabulary + "accessMode")
            throw new InvalidOperationException("The schema prefix does not identify the schema.org vocabulary.");
        string[] modes = AccessibilityTokens(metadata.AccessModes, nameof(metadata.AccessModes));
        string[] features = AccessibilityTokens(metadata.Features, nameof(metadata.Features));
        string[] hazards = AccessibilityTokens(metadata.Hazards, nameof(metadata.Hazards));
        if (metadata.SufficientAccessModes == null) throw new ArgumentNullException(nameof(metadata.SufficientAccessModes));
        string[] sufficient = metadata.SufficientAccessModes.Select(set => string.Join(",",
            AccessibilityTokens(set, nameof(metadata.SufficientAccessModes)))).Distinct(StringComparer.Ordinal).ToArray();
        string[] summaries = metadata.Summary == null ? Array.Empty<string>() : new[] { metadata.Summary.Trim() };
        if (summaries.Length != 0) RequireText(summaries[0], nameof(metadata.Summary));
        string[][] values = { modes, sufficient, features, hazards, summaries };
        EditPackageElement(RequireSection("metadata"), proposed => {
            for (int index = 0; index < AccessibilityProperties.Length; index++)
                ReplaceAccessibilityValues(proposed, AccessibilityProperties[index], values[index]);
        });
    }

    private static string[] AccessibilityTokens(IReadOnlyList<string> values, string parameter) {
        if (values == null) throw new ArgumentNullException(parameter);
        string[] result = values.Select(value => {
            if (string.IsNullOrWhiteSpace(value) || value.Any(character => char.IsWhiteSpace(character) || character == ','))
                throw new ArgumentException("Each discovery vocabulary value must be a nonempty token, without commas or whitespace.", parameter);
            XmlConvert.VerifyXmlChars(value);
            return value;
        }).Distinct(StringComparer.Ordinal).ToArray();
        if (result.Length == 0) throw new ArgumentException("At least one reviewed discovery value is required.", parameter);
        return result;
    }

    private void ReplaceAccessibilityValues(XElement metadata, string property, string[] values) {
        XElement root = metadata.Document!.Root!;
        XElement[] existing = metadata.Elements(Opf + "meta").Where(element => element.Attribute("refines") == null &&
            EpubVocabulary.Expand(root, (string?)element.Attribute("property") ?? string.Empty) == SchemaVocabulary + property).ToArray();
        foreach (XElement removed in existing.Skip(values.Length)) {
            string? id = (string?)removed.Attribute("id");
            if (id != null && root.Descendants().Attributes("refines").Any(attribute => ReferencesPackageId(attribute.Value, id)))
                throw new InvalidOperationException("An accessibility declaration is referenced by a refinement: " + id);
            removed.Remove();
        }
        for (int index = 0; index < values.Length; index++) {
            if (index < existing.Length) existing[index].Value = values[index];
            else metadata.Add(new XElement(Opf + "meta", new XAttribute("property", "schema:" + property), values[index]));
        }
    }

    private EpubPreflightCheck CheckAccessibilityDiscovery() {
        if (PackageVersion != "3.0") return new EpubPreflightCheck("accessibility-discovery", EpubPreflightStatus.NotChecked,
            new[] { new EpubDiagnostic { Code = "EPUB_PREFLIGHT_ACCESSIBILITY_VERSION", Severity = EpubDiagnosticSeverity.Info,
                Message = "This discovery check inspects EPUB 3 property metadata. Review EPUB 2 accessibility metadata independently.", Path = PackagePath } });
        var findings = new List<EpubDiagnostic>();
        for (int index = 0; index < AccessibilityProperties.Length; index++) {
            string property = AccessibilityProperties[index];
            bool present = RequireSection("metadata").Elements(Opf + "meta").Any(element => element.Attribute("refines") == null &&
                EpubVocabulary.Expand(Root, (string?)element.Attribute("property") ?? string.Empty) == SchemaVocabulary + property &&
                !string.IsNullOrWhiteSpace(element.Value));
            if (!present) Add(findings, "EPUB_PREFLIGHT_ACCESSIBILITY_METADATA_MISSING",
                property == "accessModeSufficient" || property == "accessibilitySummary" ? EpubDiagnosticSeverity.Warning : EpubDiagnosticSeverity.Error,
                "Supply reviewed schema:" + property + " discovery metadata. Do not infer accessibility or absence of hazards from a successful save.", PackagePath);
        }
        return Result("accessibility-discovery", findings);
    }
}
