using System.Xml;
using OfficeIMO.Core.Internal;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        private sealed record PackageVbaMutation(
            string Main, string Project, string? AddedRelationship, OfficeXmlTree? Relationships,
            IReadOnlyList<int> RemovedRelationships, IReadOnlyList<string> RemovedParts);

        private bool TryMutatePackageVbaProject(byte[]? project, out bool removed) {
            removed = false;
            if (EditablePackage is not { } package || !TryPlanPackageVbaMutation(package, project == null, out PackageVbaMutation? plan)) {
                return false;
            }
            if (plan == null) return true;

            byte[]? stored = project == null ? null : (byte[])project.Clone();
            if (plan.AddedRelationship is { } id) {
                string part = UniquePartName(package, VbaProjectContentType, plan.Project);
                package.CreateBinaryPart(part, VbaProjectContentType, stored!);
                package.AddRelationship(plan.Main, VbaProjectRelationshipType, part, id: id);
                string mainRelationships = OfficeOpcPartNames.GetRelationshipsPartName(plan.Main);
                if (package.GetContentType(mainRelationships) != OfficeOpcEditablePackage.RelationshipContentType) {
                    package.SetOverride(mainRelationships, OfficeOpcEditablePackage.RelationshipContentType);
                }
            } else {
                if (plan.Relationships != null) {
                    foreach (int edge in plan.RemovedRelationships) plan.Relationships.Remove(edge);
                }
                if (stored != null) {
                    package.ReplacePart(plan.Project, stored, stored.Length);
                    string projectRelationships = OfficeOpcPartNames.GetRelationshipsPartName(plan.Project);
                    if (package.PartExists(projectRelationships)) package.DeletePart(projectRelationships);
                }
                foreach (string part in plan.RemovedParts) {
                    string partRelationships = OfficeOpcPartNames.GetRelationshipsPartName(part);
                    if (package.PartExists(partRelationships)) package.DeletePart(partRelationships);
                    package.DeletePart(part);
                }
            }
            MarkPackageEdited();
            InvalidatePackageRelationships();
            removed = project == null;
            return true;
        }

        private bool TryPlanPackageVbaMutation(OfficeOpcEditablePackage package, bool remove, out PackageVbaMutation? plan) {
            plan = null;
            try {
                if (package.PartNames.Any(name => name.EndsWith(".piece", StringComparison.OrdinalIgnoreCase) &&
                    OfficeOpcPieceAssembler.HasPieceSuffix(name))) return false;
                string? main = FirstRelated(package, null, OfficeOpcPackage.OfficeDocumentRelationshipType, out _);
                if (main == null || package.GetPartName(main) is not { } mainPart) return false;

                var graph = new Dictionary<string, List<PackageGraphRelationship>>(OfficeAsciiIgnoreCaseComparer.Instance);
                var referencedData = new List<(string Owner, string Target)>();
                IReadOnlyList<OfficeOpcRelationship>? ReadRelationships(string? owner) {
                    IReadOnlyList<OfficeOpcRelationship>? edges = ReadVbaMutationRelationships(package, owner);
                    if (edges == null) return null;
                    foreach (OfficeOpcRelationship edge in edges) {
                        if (edge.IsExternal || !IsReferenceRelationship(edge.Type) || edge.Type == PowerPointPartRef.HyperlinkRelationshipType) continue;
                        if (owner == null || edge.TargetPartName is not { } target || package.GetPartName(target) is not { } part ||
                            package.GetContentType(part) is not { } contentType || !MatchesSubtypeSavePart(edge.Type, contentType)) return null;
                        referencedData.Add((PartIdentity(owner), PartIdentity(part)));
                    }
                    return edges;
                }
                if (!TryVisitPackageGraph(ReadRelationships,
                    package.GetPartName, package.GetContentType, default, relationship => {
                        string source = relationship.Owner is { } owner ? PartIdentity(owner) : string.Empty;
                        if (!graph.TryGetValue(source, out List<PackageGraphRelationship>? edges)) {
                            edges = new List<PackageGraphRelationship>();
                            graph.Add(source, edges);
                        }
                        edges.Add(relationship);
                    })) return false;

                string mainIdentity = PartIdentity(mainPart);
                if (!graph.TryGetValue(mainIdentity, out List<PackageGraphRelationship>? children)) {
                    children = new List<PackageGraphRelationship>();
                }
                PackageGraphRelationship[] existing = children.Where(edge => edge.Type == VbaProjectRelationshipType).ToArray();
                if (existing.Length == 0) return remove || TryPlanNewVbaProject(package, mainPart, out plan);
                if (existing.Length != 1 || existing[0].ContentType != VbaProjectContentType) return false;

                PackageGraphRelationship current = existing[0];
                string identity = PartIdentity(current.Part);
                IReadOnlyList<OfficeOpcRelationship>? projectLinks = ReadVbaMutationRelationships(package, current.Part);
                if (projectLinks == null || projectLinks.Any(link => link.IsExternal || IsReferenceRelationship(link.Type))) return false;
                OfficeXmlTree? relationships = null;
                int[] edgesToRemove = Array.Empty<int>();
                if (remove || projectLinks.Count != 0) {
                    relationships = package.GetRelationships(remove ? mainPart : current.Part);
                    edgesToRemove = remove
                        ? new[] { FindRelationship(relationships, current.Id) }
                        : projectLinks.Select(link => FindRelationship(relationships, link.Id)).ToArray();
                    if (edgesToRemove.Any(edge => edge == OfficeXmlTree.NoNode)) return false;
                }

                Dictionary<string, string> kept = ReachVbaPackageParts(graph, string.Empty, mainIdentity,
                    remove ? current.Id : null, remove ? null : identity);
                Dictionary<string, string> projectGraph = ReachVbaPackageParts(graph, current.Part, null, null, null);
                // An SDK data part is accounted for separately from the child parts walked above.
                if (referencedData.Any(edge => projectGraph.ContainsKey(edge.Owner) || projectGraph.ContainsKey(edge.Target))) return false;
                string[] dead = projectGraph.Where(part => !kept.ContainsKey(part.Key)).Select(part => part.Value).ToArray();
                plan = new PackageVbaMutation(mainPart, current.Part, null, relationships, edgesToRemove, dead);
                return true;
            } catch (Exception exception) when (exception is InvalidDataException or XmlException or IOException or
                ArgumentException or NotSupportedException or InvalidOperationException or OverflowException) {
                return false;
            }
        }

        private static IReadOnlyList<OfficeOpcRelationship>? ReadVbaMutationRelationships(OfficeOpcEditablePackage package, string? owner) {
            string part = OfficeOpcPartNames.GetRelationshipsPartName(owner);
            if (package.PartExists(part) && package.GetContentType(part) != OfficeOpcEditablePackage.RelationshipContentType) return null;
            IReadOnlyList<OfficeOpcRelationship>? links = ReadEditablePackageRelationships(package, owner, strictXml: true);
            if (links == null) return null;
            foreach (OfficeOpcRelationship link in links) {
                XmlConvert.VerifyNCName(link.Id);
                if (!Uri.TryCreate(link.Type, UriKind.Absolute, out _) || !Uri.TryCreate(link.Target, UriKind.RelativeOrAbsolute, out _)) return null;
            }
            return links;
        }

        private static bool TryPlanNewVbaProject(OfficeOpcEditablePackage package, string main, out PackageVbaMutation? plan) {
            const string directory = "/ppt/";
            const string part = directory + "vbaProject.bin";
            plan = null;
            if (!main.StartsWith(directory, StringComparison.Ordinal) || main.LastIndexOf('/') != directory.Length - 1 ||
                package.PartNames.Any(name => PartIdentity(name).StartsWith("/ppt/_rels/vbaProject", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(package.GetContentType(name), VbaProjectContentType, StringComparison.OrdinalIgnoreCase))) return false;
            if (ReadVbaMutationRelationships(package, main) is not { } links) return false;
            string id;
            do {
                id = NewRelationshipId();
            } while (links.Any(link => link.Id == id));
            plan = new PackageVbaMutation(main, part, id, null, Array.Empty<int>(), Array.Empty<string>());
            return true;
        }

        private static Dictionary<string, string> ReachVbaPackageParts(Dictionary<string, List<PackageGraphRelationship>> graph,
            string start, string? skipOwner, string? skipId, string? stopAt) {
            string startIdentity = PartIdentity(start);
            var found = new Dictionary<string, string>(OfficeAsciiIgnoreCaseComparer.Instance) { [startIdentity] = start };
            var pending = new Stack<string>();
            pending.Push(startIdentity);
            while (pending.Count > 0) {
                string source = pending.Pop();
                if (OfficeAsciiIgnoreCaseComparer.Instance.Equals(source, stopAt) || !graph.TryGetValue(source, out var edges)) continue;
                bool skip = skipId != null && OfficeAsciiIgnoreCaseComparer.Instance.Equals(source, skipOwner);
                foreach (PackageGraphRelationship edge in edges) {
                    if (skip && edge.Id == skipId) continue;
                    string part = PartIdentity(edge.Part);
                    if (found.ContainsKey(part)) continue;
                    found.Add(part, edge.Part);
                    pending.Push(part);
                }
            }
            return found;
        }
    }
}
