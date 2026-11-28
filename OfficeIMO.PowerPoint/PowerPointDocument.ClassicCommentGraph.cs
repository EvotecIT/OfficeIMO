using OfficeIMO.Core.Internal;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        private ClassicCommentRelationshipGraph? _classicCommentGraph;

        private bool CanMutateClassicCommentRelationships(OfficeOpcEditablePackage package, ClassicCommentPartDeletion? comments,
            ClassicCommentPartDeletion? authors) {
            if (_classicCommentGraph?.IsCurrent(package) != true) _classicCommentGraph = ReadClassicCommentGraph(package);
            return _classicCommentGraph is { } graph && graph.CanDelete(comments) && graph.CanDelete(authors);
        }

        private ClassicCommentRelationshipGraph? ReadClassicCommentGraph(OfficeOpcEditablePackage package) {
            OfficeXmlTree catalog = package.ContentTypes;
            if (!IsValidClassicCommentMutationPart(package, OfficeOpcEditablePackage.ContentTypesPartName,
                new OfficeXmlNode(catalog, catalog.Root))) return null;

            var graph = new ClassicCommentRelationshipGraph(package, catalog);
            IReadOnlyList<OfficeOpcRelationship>? ReadRelationships(string? owner) {
                var links = ReadClassicCommentMutationLinks(package, owner, validateXml: true);
                if (links == null) return null;
                string name = OfficeOpcPartNames.GetRelationshipsPartName(owner);
                if (package.GetParsedXml(name) is { } tree) graph.TrackRelationships(name, tree);
                foreach (OfficeOpcRelationship edge in links) {
                    if (!edge.IsExternal && edge.TargetPartName is { } target) graph.AddReference(target, owner, edge.Id);
                }
                return links;
            }

            return TryVisitPackageGraph(ReadRelationships, package.GetPartName, package.GetContentType, default,
                static _ => { }, validateForSave: true) ? graph : null;
        }

        /// <summary>Reuses a validated graph while the part catalog and the reachable relationship trees agree with it.</summary>
        private sealed class ClassicCommentRelationshipGraph {
            private readonly OfficeOpcEditablePackage _package;
            private readonly OfficeXmlTree _catalog;
            private int _catalogVersion;
            private readonly HashSet<string> _parts;
            private readonly Dictionary<string, (OfficeXmlTree Tree, int Version)> _relationships =
                new Dictionary<string, (OfficeXmlTree, int)>(OfficeAsciiIgnoreCaseComparer.Instance);
            private readonly Dictionary<string, (string? Owner, string Id, bool Shared)> _references =
                new Dictionary<string, (string?, string, bool)>(OfficeAsciiIgnoreCaseComparer.Instance);

            internal ClassicCommentRelationshipGraph(OfficeOpcEditablePackage package, OfficeXmlTree catalog) {
                _package = package;
                _catalog = catalog;
                _catalogVersion = catalog.Version;
                _parts = new HashSet<string>(package.PartNames, OfficeAsciiIgnoreCaseComparer.Instance);
            }

            internal bool IsCurrent(OfficeOpcEditablePackage package) {
                if (package != _package || package.ContentTypes != _catalog || _catalog.Version != _catalogVersion) return false;
                int count = 0;
                foreach (string part in package.PartNames) {
                    if (!_parts.Contains(part)) return false;
                    count++;
                }
                if (count != _parts.Count) return false;
                foreach (var relationship in _relationships) {
                    if (relationship.Value.Tree.Version != relationship.Value.Version ||
                        package.GetParsedXml(relationship.Key) != relationship.Value.Tree) return false;
                }
                return true;
            }

            internal void TrackRelationships(string name, OfficeXmlTree tree) => _relationships[name] = (tree, tree.Version);

            internal void AddReference(string target, string? owner, string id) {
                string part = PartIdentity(target);
                _references[part] = (owner, id, _references.ContainsKey(part));
            }

            internal bool HasReference(string part) => _references.ContainsKey(PartIdentity(part));

            internal bool CanDelete(ClassicCommentPartDeletion? deletion) =>
                deletion == null || _references.TryGetValue(PartIdentity(deletion.Part), out var reference) &&
                !reference.Shared && OfficeAsciiIgnoreCaseComparer.Instance.Equals(reference.Owner, deletion.Owner) &&
                reference.Id == deletion.Id && (!deletion.DeleteRelationships ||
                    !_references.ContainsKey(PartIdentity(OfficeOpcPartNames.GetRelationshipsPartName(deletion.Owner))));

            internal void RecordDeletion(ClassicCommentPartDeletion deletion, string ownerRelationships) {
                // The removed part is an unshared leaf. Other relationship targets remain reachable.
                _parts.Remove(deletion.Part);
                _references.Remove(PartIdentity(deletion.Part));
                if (deletion.DeleteRelationships) {
                    _parts.Remove(ownerRelationships);
                    _relationships.Remove(ownerRelationships);
                } else {
                    TrackRelationships(ownerRelationships, deletion.Relationships);
                }
                _catalogVersion = _catalog.Version;
            }
        }
    }
}
