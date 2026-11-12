using System.IO.Packaging;

namespace OfficeIMO.Excel.Utilities;

/// <summary>
/// Checks and corrects the content-types part directly on its bytes when the package and both XML parts have the
/// plain shape Office writers produce. The result is the same decision and the same bytes the XML route produces;
/// any other shape returns <see cref="PlainOutcome.Unsure"/> so that route runs and reports its own errors.
/// </summary>
internal static partial class ExcelPackageUtilities {
    private const string ContentTypesNamespace = "http://schemas.openxmlformats.org/package/2006/content-types";
    private const string RelationshipsNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";

    internal enum PlainOutcome { Unsure, Unchanged, Normalized }

    internal static PlainOutcome TryNormalizePlainContentTypes(byte[] package, long maxCharactersInPart, out byte[]? normalized) {
        normalized = null;
        try {
            ExcelPackageEntrySplicer? splicer = ExcelPackageEntrySplicer.TryOpenPlain(package);
            if (splicer == null) return PlainOutcome.Unsure;

            byte[]? typesXml = splicer.ReadExact(ContentTypesEntry);
            if (typesXml == null) return PlainOutcome.Unchanged;
            long characterLimit = maxCharactersInPart > 0 ? Math.Min(maxCharactersInPart, MaximumContentTypesEntryBytes) : MaximumContentTypesEntryBytes;
            if (typesXml.Length > characterLimit) return PlainOutcome.Unsure;
            List<PlainElement>? types = PlainXml.Read(typesXml, "Types", ContentTypesNamespace, out bool typesHaveWhitespace);
            if (types == null) return PlainOutcome.Unsure;
            foreach (PlainElement element in types) {
                if (!element.HasExactly(element.Name == "Default" ? "Extension" : element.Name == "Override" ? "PartName" : null, "ContentType")) {
                    return PlainOutcome.Unsure;
                }
            }

            if (!TryReadPlainRootParts(splicer, maxCharactersInPart, out string? workbook, out string? app, out string? core)) {
                return PlainOutcome.Unsure;
            }

            if (!PlainNeedsNormalization(types, workbook, app, core)) return PlainOutcome.Unchanged;
            // Whitespace between nodes would be carried into the rewritten part, which this writer does not reproduce.
            if (typesHaveWhitespace) return PlainOutcome.Unsure;
            if (!PlainNormalize(types, splicer, typesXml.Length, workbook, app, core)) return PlainOutcome.Unchanged;

            byte[] output = PlainXml.WriteContentTypes(types);
            if (output.Length > MaximumContentTypesEntryBytes) return PlainOutcome.Unsure;
            normalized = splicer.BuildWith(ContentTypesEntry, output);
            return PlainOutcome.Normalized;
        } catch {
            normalized = null;
            return PlainOutcome.Unsure;
        }
    }

    private static bool PlainNeedsNormalization(List<PlainElement> types, string? workbook, string? app, string? core) {
        int xmlDefaultCount = 0;
        bool xmlDefaultIsCorrect = false;
        bool workbookOverrideIsCorrect = workbook == null;
        bool appPropsOverrideIsCorrect = app == null;
        bool corePropsOverrideIsCorrect = core == null;
        foreach (PlainElement element in types) {
            if (element.Name == "Default") {
                if (string.Equals(element.Get("Extension"), "xml", StringComparison.OrdinalIgnoreCase)) {
                    xmlDefaultCount++;
                    xmlDefaultIsCorrect = string.Equals(element.Get("ContentType"), "application/xml", StringComparison.OrdinalIgnoreCase);
                }
                continue;
            }

            string? partName = element.Get("PartName");
            string? contentType = element.Get("ContentType");
            if (string.Equals(partName, workbook, StringComparison.OrdinalIgnoreCase)) {
                workbookOverrideIsCorrect = IsSupportedWorkbookContentType(contentType);
            } else if (string.Equals(partName, app, StringComparison.OrdinalIgnoreCase)) {
                appPropsOverrideIsCorrect = string.Equals(contentType, AppPropsContentType, StringComparison.OrdinalIgnoreCase);
            } else if (string.Equals(partName, core, StringComparison.OrdinalIgnoreCase)) {
                corePropsOverrideIsCorrect = string.Equals(contentType, CorePropsContentType, StringComparison.OrdinalIgnoreCase);
            }
        }

        return xmlDefaultCount != 1 || !xmlDefaultIsCorrect || !workbookOverrideIsCorrect || !appPropsOverrideIsCorrect || !corePropsOverrideIsCorrect;
    }

    /// <summary>Applies the edits <see cref="BuildNormalizedContentTypes"/> makes, in the same order; returns whether any applied.</summary>
    private static bool PlainNormalize(List<PlainElement> types, ExcelPackageEntrySplicer splicer, int originalLength,
        string? workbook, string? app, string? core) {
        string? workbookContentType = workbook == null ? null : PlainEffectiveContentType(types, workbook);
        if (!IsSupportedWorkbookContentType(workbookContentType)) workbookContentType = WorkbookContentType;
        bool changed = false;

        List<PlainElement> xmlDefaults = types.Where(element => element.Name == "Default"
            && string.Equals(element.Get("Extension"), "xml", StringComparison.OrdinalIgnoreCase)).ToList();

        if (xmlDefaults.Count != 0) {
            string? previousXmlType = xmlDefaults[0].Get("ContentType");
            if (!string.IsNullOrWhiteSpace(previousXmlType)
                && !string.Equals(previousXmlType, "application/xml", StringComparison.OrdinalIgnoreCase)) {
                var overriddenParts = new HashSet<string>(types.Where(element => element.Name == "Override")
                    .Select(element => element.Get("PartName") ?? string.Empty), StringComparer.OrdinalIgnoreCase);
                long budget = MaximumContentTypesEntryBytes - originalLength;
                foreach (string name in splicer.Names) {
                    if (name == ContentTypesEntry || !name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)) continue;
                    string partName = "/" + name;
                    if (!overriddenParts.Add(partName)) continue;
                    // A generous bound on the declaration size, so the XML route's exact limit can never be the one reached.
                    budget -= 128 + ContentTypesNamespace.Length + 3L * (partName.Length + previousXmlType!.Length);
                    if (budget < 0 || !PlainXml.IsPlainValue(partName)) throw new InvalidOperationException("Outside the plain shape.");
                    types.Add(PlainElement.Override(partName, previousXmlType));
                    changed = true;
                }
            }
        }

        if (xmlDefaults.Count == 0) {
            types.Insert(0, PlainElement.Default("xml", "application/xml"));
            changed = true;
        } else {
            if (!string.Equals(xmlDefaults[0].Get("ContentType"), "application/xml", StringComparison.OrdinalIgnoreCase)) {
                xmlDefaults[0].Set("ContentType", "application/xml");
                changed = true;
            }
            for (int index = 1; index < xmlDefaults.Count; index++) {
                types.Remove(xmlDefaults[index]);
                changed = true;
            }
        }

        bool EnsureOverride(string partName, string contentType) {
            PlainElement? existing = types.FirstOrDefault(element => element.Name == "Override"
                && string.Equals(element.Get("PartName"), partName, StringComparison.OrdinalIgnoreCase));
            if (existing == null) {
                if (!PlainXml.IsPlainValue(partName)) throw new InvalidOperationException("Outside the plain shape.");
                int firstOverride = types.FindIndex(element => element.Name == "Override");
                if (firstOverride >= 0) types.Insert(firstOverride, PlainElement.Override(partName, contentType));
                else types.Add(PlainElement.Override(partName, contentType));
                return true;
            }
            if (!string.Equals(existing.Get("ContentType"), contentType, StringComparison.OrdinalIgnoreCase)) {
                existing.Set("ContentType", contentType);
                return true;
            }
            return false;
        }

        if (workbook != null && EnsureOverride(workbook, workbookContentType!)) changed = true;
        if (app != null && EnsureOverride(app, AppPropsContentType)) changed = true;
        if (core != null && EnsureOverride(core, CorePropsContentType)) changed = true;
        return changed;
    }

    private static string? PlainEffectiveContentType(List<PlainElement> types, string partName) {
        PlainElement? match = types.FirstOrDefault(element => element.Name == "Override"
            && string.Equals(element.Get("PartName"), partName, StringComparison.OrdinalIgnoreCase));
        if (match?.Get("ContentType") is string overridden) return overridden;
        string extension = Path.GetExtension(partName).TrimStart('.');
        return types.FirstOrDefault(element => element.Name == "Default"
            && string.Equals(element.Get("Extension"), extension, StringComparison.OrdinalIgnoreCase))?.Get("ContentType");
    }

    /// <summary>Resolves the root parts as <see cref="ReadRootParts"/> does; false wherever that route would fail.</summary>
    private static bool TryReadPlainRootParts(ExcelPackageEntrySplicer splicer, long maxCharactersInPart,
        out string? workbook, out string? app, out string? core) {
        workbook = app = core = null;
        byte[]? xml = splicer.ReadExact("_rels/.rels");
        if (xml == null) return true;
        long characterLimit = maxCharactersInPart > 0 ? Math.Min(maxCharactersInPart, MaximumContentTypesEntryBytes) : MaximumContentTypesEntryBytes;
        if (xml.Length > characterLimit) return false;
        List<PlainElement>? relationships = PlainXml.Read(xml, "Relationships", RelationshipsNamespace, out _);
        if (relationships == null) return false;
        foreach (PlainElement relationship in relationships) {
            if (relationship.Name != "Relationship" || !relationship.HasOnly("Id", "Type", "Target", "TargetMode")) return false;
        }

        return TryFindPlainRootPart(splicer, relationships, "/officeDocument", out workbook)
            && TryFindPlainRootPart(splicer, relationships, "/extended-properties", out app)
            && TryFindPlainRootPart(splicer, relationships, "/metadata/core-properties", out core);
    }

    private static bool TryFindPlainRootPart(ExcelPackageEntrySplicer splicer, List<PlainElement> relationships, string suffix, out string? partName) {
        partName = null;
        PlainElement[] matches = relationships.Where(item => IsOfficeRootRelationship(item.Get("Type"), suffix)).ToArray();
        if (matches.Length == 0) return true;
        if (matches.Length != 1 || string.Equals(matches[0].Get("TargetMode"), "External", StringComparison.OrdinalIgnoreCase)) return false;
        string? target = matches[0].Get("Target");
        if (string.IsNullOrEmpty(target) || !Uri.TryCreate(target, UriKind.Relative, out Uri? targetUri)) return false;
        string resolved = PackUriHelper.ResolvePartUri(new Uri("/", UriKind.Relative), targetUri).OriginalString;
        if (!splicer.ContainsExact(resolved.TrimStart('/'))) return false;
        partName = resolved;
        return true;
    }

    /// <summary>An element with attributes only, in document order.</summary>
    private sealed class PlainElement {
        private readonly List<KeyValuePair<string, string>> _attributes;

        internal PlainElement(string name, List<KeyValuePair<string, string>> attributes) {
            Name = name;
            _attributes = attributes;
        }

        internal string Name { get; }
        internal IReadOnlyList<KeyValuePair<string, string>> Attributes => _attributes;

        internal static PlainElement Default(string extension, string contentType) => new PlainElement("Default",
            new List<KeyValuePair<string, string>> { new("Extension", extension), new("ContentType", contentType) });

        internal static PlainElement Override(string partName, string contentType) => new PlainElement("Override",
            new List<KeyValuePair<string, string>> { new("PartName", partName), new("ContentType", contentType) });

        internal string? Get(string name) {
            foreach (KeyValuePair<string, string> attribute in _attributes) {
                if (attribute.Key == name) return attribute.Value;
            }
            return null;
        }

        internal void Set(string name, string value) {
            int index = _attributes.FindIndex(attribute => attribute.Key == name);
            _attributes[index] = new KeyValuePair<string, string>(name, value);
        }

        internal bool HasExactly(string? first, string second) =>
            first != null && _attributes.Count == 2 && Get(first) != null && Get(second) != null;

        internal bool HasOnly(params string[] names) => _attributes.All(attribute => names.Contains(attribute.Key));
    }

    /// <summary>Reads and writes the narrow XML shape of package parts byte by byte.</summary>
    private static class PlainXml {
        private static readonly byte[] Declaration = Encoding.ASCII.GetBytes("<?xml version=\"1.0\" encoding=\"utf-8\"?>");

        /// <summary>
        /// Returns the children of the root, or null unless the part is an optional UTF-8 BOM, an optional plain
        /// declaration, the root with exactly the given default namespace, and childless elements whose attribute
        /// values are printable ASCII without markup characters. Whitespace may separate the declaration and nodes.
        /// </summary>
        internal static List<PlainElement>? Read(byte[] xml, string rootName, string rootNamespace, out bool hasWhitespace) {
            hasWhitespace = false;
            int position = 0;
            if (xml.Length >= 3 && xml[0] == 0xEF && xml[1] == 0xBB && xml[2] == 0xBF) position = 3;
            if (!SkipDeclaration(xml, ref position)) return null;
            hasWhitespace |= SkipWhitespace(xml, ref position);
            if (!Literal(xml, ref position, "<" + rootName + " xmlns=\"" + rootNamespace + "\">")) return null;

            var elements = new List<PlainElement>();
            while (true) {
                hasWhitespace |= SkipWhitespace(xml, ref position);
                if (Literal(xml, ref position, "</" + rootName + ">")) break;
                if (!Literal(xml, ref position, "<")) return null;
                string? name = Name(xml, ref position);
                if (name == null) return null;
                var attributes = new List<KeyValuePair<string, string>>();
                while (true) {
                    bool separated = SkipWhitespace(xml, ref position);
                    if (Literal(xml, ref position, "/>")) break;
                    if (!separated) return null;
                    string? attributeName = Name(xml, ref position);
                    if (attributeName == null || !Literal(xml, ref position, "=") || position >= xml.Length) return null;
                    byte quote = xml[position++];
                    if (quote != (byte)'"' && quote != (byte)'\'') return null;
                    int start = position;
                    while (position < xml.Length && xml[position] != quote) {
                        if (!IsPlainCharacter(xml[position])) return null;
                        position++;
                    }
                    if (position >= xml.Length) return null;
                    string value = Encoding.ASCII.GetString(xml, start, position - start);
                    position++;
                    if (attributes.Any(attribute => attribute.Key == attributeName)) return null;
                    attributes.Add(new KeyValuePair<string, string>(attributeName, value));
                }
                elements.Add(new PlainElement(name, attributes));
            }

            hasWhitespace |= SkipWhitespace(xml, ref position);
            return position == xml.Length ? elements : null;
        }

        /// <summary>Serializes the content types as the XML route's writer does for this shape.</summary>
        internal static byte[] WriteContentTypes(List<PlainElement> types) {
            var output = new MemoryStream(Declaration.Length + 96 + types.Count * 128);
            output.Write(Declaration, 0, Declaration.Length);
            WriteAscii(output, "<Types xmlns=\"" + ContentTypesNamespace + "\">");
            foreach (PlainElement element in types) {
                WriteAscii(output, "<" + element.Name);
                foreach (KeyValuePair<string, string> attribute in element.Attributes) {
                    WriteAscii(output, " " + attribute.Key + "=\"" + attribute.Value + "\"");
                }
                WriteAscii(output, " />");
            }
            WriteAscii(output, "</Types>");
            return output.ToArray();
        }

        internal static bool IsPlainValue(string value) => value.All(character => character < 0x80 && IsPlainCharacter((byte)character));

        // Characters a writer copies unchanged into a double-quoted attribute.
        private static bool IsPlainCharacter(byte value) =>
            value >= 0x20 && value <= 0x7E && value != (byte)'<' && value != (byte)'>' && value != (byte)'&' && value != (byte)'"' && value != (byte)'\'';

        private static bool SkipDeclaration(byte[] xml, ref int position) {
            if (!Literal(xml, ref position, "<?xml")) return true;
            if (!Literal(xml, ref position, " version=\"1.0\"")) return false;
            if (!Literal(xml, ref position, " encoding=\"utf-8\"")) Literal(xml, ref position, " encoding=\"UTF-8\"");
            if (!Literal(xml, ref position, " standalone=\"yes\"")) Literal(xml, ref position, " standalone=\"no\"");
            return Literal(xml, ref position, "?>");
        }

        private static bool SkipWhitespace(byte[] xml, ref int position) {
            int start = position;
            while (position < xml.Length && (xml[position] == ' ' || xml[position] == '\t' || xml[position] == '\r' || xml[position] == '\n')) position++;
            return position > start;
        }

        private static string? Name(byte[] xml, ref int position) {
            int start = position;
            while (position < xml.Length && ((xml[position] | 0x20) >= 'a' && (xml[position] | 0x20) <= 'z')) position++;
            return position > start ? Encoding.ASCII.GetString(xml, start, position - start) : null;
        }

        private static bool Literal(byte[] xml, ref int position, string text) {
            if (position + text.Length > xml.Length) return false;
            for (int index = 0; index < text.Length; index++) {
                if (xml[position + index] != text[index]) return false;
            }
            position += text.Length;
            return true;
        }

        private static void WriteAscii(MemoryStream output, string text) {
            foreach (char character in text) output.WriteByte((byte)character);
        }
    }
}
