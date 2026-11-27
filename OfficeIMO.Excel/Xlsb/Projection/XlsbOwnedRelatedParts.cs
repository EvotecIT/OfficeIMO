using System.Globalization;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Excel.Xlsb.Model;
using OfficeIMO.Excel.Xlsb.Package;

namespace OfficeIMO.Excel.Xlsb.Projection {
    /// <summary>Copies native worksheet, chart and VBA carriers into the projected workbook's owned OPC package.</summary>
    internal sealed class XlsbOwnedRelatedParts {
        private readonly OfficeOpcEditablePackage _package;
        private readonly XlsbWorkbook _workbook;
        private readonly CancellationToken _cancellationToken;
        private readonly Dictionary<string, (string Target, string RelationshipType)> _copied = new(StringComparer.OrdinalIgnoreCase);
        private int _nextGeneratedName;

        internal XlsbOwnedRelatedParts(OfficeOpcEditablePackage package, XlsbWorkbook workbook, CancellationToken cancellationToken) {
            _package = package;
            _workbook = workbook;
            _cancellationToken = cancellationToken;
        }

        /// <summary>Links a part copied from the XLSB package under the same logical relationship id.</summary>
        internal void Link(string owner, string source, string relationshipType, string? id = null) {
            var pending = new Queue<(string Source, string Target)>();
            string target = CopyPart(source, relationshipType, pending);
            _package.AddRelationship(owner, relationshipType, OfficeOpcPartNames.GetRelativeTarget(owner, target), id: id);
            CopyRelationships(pending);
        }

        /// <summary>Replaces the owned creation shell's default theme with the original native theme graph.</summary>
        internal void ReplaceTheme(string owner, string themePartName, string source) {
            const string relationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "theme";
            const string contentType = "application/vnd.openxmlformats-officedocument.theme+xml";
            const string replacement = "/xl/theme/theme2.xml";
            XlsbProjectedPart original = _workbook.RelatedParts[source];
            _package.DeletePart(themePartName);
            OfficeXmlTree relationships = _package.GetRelationships(owner);
            var root = new OfficeXmlNode(relationships, relationships.Root);
            foreach (OfficeXmlNode relationship in root.Children(OfficeXmlNamespaces.PackageRelationships, "Relationship")) {
                if (relationship.Attribute("Type") == relationshipType) {
                    relationship.SetAttribute("Target", OfficeOpcPartNames.GetRelativeTarget(owner, replacement));
                    break;
                }
            }
            _package.CreateBinaryPart(replacement, contentType, original.Bytes);
            var pending = new Queue<(string Source, string Target)>();
            Register(source, replacement, relationshipType, pending);
            CopyRelationships(pending);
        }

        private string CopyPart(string source, string relationshipType, Queue<(string Source, string Target)> pending) {
            if (_copied.TryGetValue(source, out var existing)) {
                if (existing.RelationshipType != relationshipType)
                    throw new NotSupportedException("One related part cannot be projected through conflicting relationship types.");
                return existing.Target;
            }

            XlsbProjectedPart part = _workbook.RelatedParts[source];
            string target = "/" + source.TrimStart('/');
            while (_package.PartExists(target)) {
                string extension = Path.GetExtension(source);
                target = "/xl/officeimo-native/part" + (++_nextGeneratedName).ToString(CultureInfo.InvariantCulture)
                    + (extension.Length == 0 ? ".bin" : extension);
            }
            _package.CreateBinaryPart(target, part.ContentType, part.Bytes);
            Register(source, target, relationshipType, pending);
            return target;
        }

        private void Register(string source, string target, string relationshipType, Queue<(string Source, string Target)> pending) {
            _copied.Add(source, (target, relationshipType));
            _workbook.ProjectionPartUris[source] = target;
            pending.Enqueue((source, target));
        }

        private void CopyRelationships(Queue<(string Source, string Target)> pending) {
            while (pending.Count != 0) {
                var owner = pending.Dequeue();
                foreach (XlsbPackageRelationship relationship in _workbook.RelatedParts[owner.Source].Relationships.Values) {
                    _cancellationToken.ThrowIfCancellationRequested();
                    if (relationship.IsExternal) {
                        var uri = new Uri(relationship.Target, UriKind.RelativeOrAbsolute);
                        _package.AddRelationship(owner.Target, relationship.Type, uri.OriginalString, external: true, id: relationship.Id);
                    } else {
                        string source = XlsbPackagePartReader.ResolveTarget(owner.Source, relationship.Target);
                        string target = CopyPart(source, relationship.Type, pending);
                        _package.AddRelationship(owner.Target, relationship.Type,
                            OfficeOpcPartNames.GetRelativeTarget(owner.Target, target), id: relationship.Id);
                    }
                }
            }
        }
    }
}
