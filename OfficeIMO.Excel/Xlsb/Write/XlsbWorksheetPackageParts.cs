using OfficeIMO.Excel.Xlsb.Package;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>A package part that a generated worksheet adds next to itself, such as a comments or table part.</summary>
    internal sealed class XlsbGeneratedPart {
        internal XlsbGeneratedPart(string partName, string contentType, byte[] content) {
            PartName = partName ?? throw new ArgumentNullException(nameof(partName));
            ContentType = contentType ?? throw new ArgumentNullException(nameof(contentType));
            Content = content ?? throw new ArgumentNullException(nameof(content));
        }

        /// <summary>The package path without a leading slash, for example <c>xl/comments1.bin</c>.</summary>
        internal string PartName { get; }

        internal string ContentType { get; }

        internal byte[] Content { get; }

        /// <summary>Relationships from this part, written as its own relationship part when not empty.</summary>
        internal List<XlsbPackageRelationship> Relationships { get; } = new List<XlsbPackageRelationship>();
    }

    /// <summary>Allocates package part names that are unique across one generated package.</summary>
    internal sealed class XlsbPartNameAllocator {
        private readonly HashSet<string> _used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, int> _next = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        internal XlsbPartNameAllocator(IEnumerable<string>? existing = null) {
            if (existing == null) return;
            foreach (string name in existing) _used.Add(name.TrimStart('/'));
        }

        /// <summary>Returns <paramref name="prefix"/> + n + <paramref name="extension"/> for the first unused n from 1.</summary>
        internal string Next(string prefix, string extension) {
            _next.TryGetValue(prefix + extension, out int number);
            string name;
            do {
                number++;
                name = prefix + number.ToString(System.Globalization.CultureInfo.InvariantCulture) + extension;
            } while (!_used.Add(name));
            _next[prefix + extension] = number;
            return name;
        }
    }

    /// <summary>
    /// The parts and worksheet relationships that relationship-backed worksheet content (comments, tables, drawings)
    /// adds to a generated worksheet. Relationship ids never collide with the hyperlink ids already emitted.
    /// </summary>
    internal sealed class XlsbWorksheetPackageParts {
        private readonly HashSet<string> _relationshipIds;
        private readonly List<XlsbPackageRelationship> _relationships = new List<XlsbPackageRelationship>();
        private readonly List<XlsbGeneratedPart> _parts = new List<XlsbGeneratedPart>();

        internal XlsbWorksheetPackageParts(string worksheetPartName, XlsbPartNameAllocator names, IEnumerable<string> reservedRelationshipIds) {
            WorksheetPartName = worksheetPartName ?? throw new ArgumentNullException(nameof(worksheetPartName));
            Names = names ?? throw new ArgumentNullException(nameof(names));
            _relationshipIds = new HashSet<string>(reservedRelationshipIds ?? Array.Empty<string>(), StringComparer.Ordinal);
        }

        /// <summary>The worksheet's own package path without a leading slash, for example <c>xl/worksheets/sheet1.bin</c>.</summary>
        internal string WorksheetPartName { get; }

        internal XlsbPartNameAllocator Names { get; }

        internal IReadOnlyList<XlsbPackageRelationship> Relationships => _relationships;

        internal IReadOnlyList<XlsbGeneratedPart> Parts => _parts;

        internal bool IsEmpty => _relationships.Count == 0 && _parts.Count == 0;

        internal void AddPart(XlsbGeneratedPart part) => _parts.Add(part ?? throw new ArgumentNullException(nameof(part)));

        /// <summary>
        /// Adds a worksheet relationship to <paramref name="targetPartName"/> and returns its id: <paramref name="preferredId"/>,
        /// the id the editable model already uses, when it is still free, otherwise the next unused rIdN.
        /// </summary>
        internal string AddRelationship(string type, string targetPartName, string? preferredId = null) {
            string id = !string.IsNullOrEmpty(preferredId) && _relationshipIds.Add(preferredId!) ? preferredId! : AllocateRelationshipId(_relationshipIds);
            _relationships.Add(new XlsbPackageRelationship(id, type, RelativeTarget(WorksheetPartName, targetPartName), isExternal: false));
            return id;
        }

        /// <summary>The worksheet relationships as the hyperlink plan's preserved-relationship input.</summary>
        internal IReadOnlyDictionary<string, XlsbPackageRelationship> RelationshipsById() =>
            _relationships.ToDictionary(relationship => relationship.Id, StringComparer.Ordinal);

        /// <summary>A relationship target from the folder of <paramref name="sourcePartName"/> to <paramref name="targetPartName"/>.</summary>
        internal static string RelativeTarget(string sourcePartName, string targetPartName) {
            string[] source = sourcePartName.TrimStart('/').Split('/');
            string[] target = targetPartName.TrimStart('/').Split('/');
            int common = 0;
            while (common < source.Length - 1 && common < target.Length - 1
                && string.Equals(source[common], target[common], StringComparison.OrdinalIgnoreCase)) {
                common++;
            }
            var builder = new StringBuilder();
            for (int index = common; index < source.Length - 1; index++) builder.Append("../");
            builder.Append(string.Join("/", target, common, target.Length - common));
            return builder.ToString();
        }

        internal static byte[] CreateRelationshipPart(IEnumerable<XlsbPackageRelationship> relationships) {
            var builder = new StringBuilder(256);
            builder.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            builder.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
            foreach (XlsbPackageRelationship relationship in relationships) {
                builder.Append("<Relationship Id=\"").Append(System.Security.SecurityElement.Escape(relationship.Id))
                    .Append("\" Type=\"").Append(System.Security.SecurityElement.Escape(relationship.Type))
                    .Append("\" Target=\"").Append(System.Security.SecurityElement.Escape(relationship.Target));
                if (relationship.IsExternal) builder.Append("\" TargetMode=\"External");
                builder.Append("\"/>");
            }
            builder.Append("</Relationships>");
            return new UTF8Encoding(encoderShouldEmitUTF8Identifier: false).GetBytes(builder.ToString());
        }

        /// <summary>The relationship part path for <paramref name="partName"/>, for example <c>xl/_rels/comments1.bin.rels</c>.</summary>
        internal static string RelationshipPartName(string partName) {
            string trimmed = partName.TrimStart('/');
            int slash = trimmed.LastIndexOf('/');
            return slash < 0
                ? "_rels/" + trimmed + ".rels"
                : trimmed.Substring(0, slash + 1) + "_rels/" + trimmed.Substring(slash + 1) + ".rels";
        }

        private static string AllocateRelationshipId(HashSet<string> used) {
            for (int number = 1; ; number++) {
                string id = "rId" + number.ToString(System.Globalization.CultureInfo.InvariantCulture);
                if (used.Add(id)) return id;
            }
        }
    }
}
