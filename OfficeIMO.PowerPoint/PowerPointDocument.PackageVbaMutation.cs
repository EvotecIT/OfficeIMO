using System.Xml;
using OfficeIMO.Core.Internal;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        private sealed record PackageVbaMutation(
            string Main, string Project, string? AddedRelationship, OfficeXmlTree? Relationships,
            IReadOnlyList<int> RemovedRelationships, IReadOnlyList<string> RemovedParts,
            bool RemoveProjectRelationships = false, IReadOnlyCollection<string>? RecoveryParts = null);

        private bool TryMutatePackageVbaProject(byte[]? project, out bool removed, OfficeVbaWriteOptions? sourceOptions = null) {
            removed = false;
            if (EditablePackage is not { } package || !TryPlanPackageVbaMutation(package, project == null,
                out PackageVbaMutation? plan, preserveNonSignatureParts: sourceOptions != null)) {
                return false;
            }
            if (plan == null) return true;
            if (sourceOptions != null) {
                if (sourceOptions.MaximumRecoveryBytes < 1) throw new ArgumentOutOfRangeException(nameof(sourceOptions.MaximumRecoveryBytes));
                if (plan.RecoveryParts is { } parts) {
                    if (PackageVbaBytesMatch(package, plan.Project, project!)) return true;
                    if (plan.RemovedRelationships.Count != 0 && !sourceOptions.AllowSignatureRemoval) {
                        throw new InvalidOperationException("Changing the VBA project invalidates its signatures. Explicit signature removal is required.");
                    }
                    ValidatePackageVbaRecovery(package, parts, sourceOptions.MaximumRecoveryBytes);
                }
            }

            // The source editor supplies its own byte array. The byte-array API retains caller ownership.
            byte[]? stored = project == null || sourceOptions != null ? project : (byte[])project.Clone();
            if (plan.AddedRelationship is { } id) {
                string part = UniquePartName(package, VbaProjectContentType, plan.Project);
                package.CreateBinaryPart(part, VbaProjectContentType, stored!);
                package.AddRelationship(plan.Main, VbaProjectRelationshipType, part, id: id);
                EnsurePackageRelationshipContentType(package, plan.Main);
            } else {
                if (plan.Relationships != null) {
                    foreach (int edge in plan.RemovedRelationships) plan.Relationships.Remove(edge);
                }
                if (stored != null) {
                    package.ReplacePart(plan.Project, stored, stored.Length);
                    if (plan.RemoveProjectRelationships) {
                        package.DeletePart(OfficeOpcPartNames.GetRelationshipsPartName(plan.Project));
                    }
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

        private bool TryPlanPackageVbaMutation(OfficeOpcEditablePackage package, bool remove, out PackageVbaMutation? plan,
            bool preserveNonSignatureParts = false) {
            plan = null;
            try {
                if (package.PartNames.Any(name => name.EndsWith(".piece", StringComparison.OrdinalIgnoreCase) &&
                    OfficeOpcPieceAssembler.HasPieceSuffix(name))) return false;
                if (preserveNonSignatureParts && !HasUniqueContentTypeEntries(package.ContentTypes)) return false;
                string? main = FirstRelated(package, null, OfficeOpcPackage.OfficeDocumentRelationshipType, out _);
                if (main == null || package.GetPartName(main) is not { } mainPart) return false;

                var graph = new Dictionary<string, List<PackageGraphRelationship>>(OfficeAsciiIgnoreCaseComparer.Instance);
                var referencedData = new List<(string Owner, string Target)>();
                IReadOnlyList<OfficeOpcRelationship>? ReadRelationships(string? owner) {
                    IReadOnlyList<OfficeOpcRelationship>? edges = ReadPackageMutationRelationships(package, owner);
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
                    }, validateForSave: preserveNonSignatureParts)) return false;

                string mainIdentity = PartIdentity(mainPart);
                if (!graph.TryGetValue(mainIdentity, out List<PackageGraphRelationship>? children)) {
                    children = new List<PackageGraphRelationship>();
                }
                PackageGraphRelationship[] existing = children.Where(edge => edge.Type == VbaProjectRelationshipType).ToArray();
                if (existing.Length == 0) return remove || TryPlanNewVbaProject(package, mainPart, out plan);
                if (existing.Length != 1 || existing[0].ContentType != VbaProjectContentType) return false;

                PackageGraphRelationship current = existing[0];
                string identity = PartIdentity(current.Part);
                IReadOnlyList<OfficeOpcRelationship>? projectLinks = ReadPackageMutationRelationships(package, current.Part);
                if (projectLinks == null || projectLinks.Any(link => link.IsExternal || IsReferenceRelationship(link.Type))) return false;
                HashSet<string>? signatureIds = null;
                if (preserveNonSignatureParts) {
                    signatureIds = new HashSet<string>(StringComparer.Ordinal);
                    if (graph.TryGetValue(identity, out List<PackageGraphRelationship>? projectChildren)) {
                        foreach (PackageGraphRelationship edge in projectChildren) {
                            if (IsVbaProjectSignature(edge.Type, edge.ContentType)) signatureIds.Add(edge.Id);
                        }
                    }
                }
                OfficeXmlTree? relationships = null;
                int[] edgesToRemove = Array.Empty<int>();
                if (remove || (signatureIds?.Count ?? projectLinks.Count) != 0) {
                    relationships = package.GetRelationships(remove ? mainPart : current.Part);
                    edgesToRemove = remove
                        ? new[] { FindRelationship(relationships, current.Id) }
                        : projectLinks.Where(link => signatureIds == null || signatureIds.Contains(link.Id))
                            .Select(link => FindRelationship(relationships, link.Id)).ToArray();
                    if (edgesToRemove.Any(edge => edge == OfficeXmlTree.NoNode)) return false;
                }

                Dictionary<string, string> projectGraph = ReachVbaPackageParts(graph, current.Part, null, null, null);
                // An SDK data part is accounted for separately from the child parts walked above.
                if (referencedData.Any(edge => projectGraph.ContainsKey(edge.Owner) || projectGraph.ContainsKey(edge.Target))) return false;
                string[] dead = Array.Empty<string>();
                if (signatureIds == null || signatureIds.Count > 0) {
                    Dictionary<string, string> kept = ReachVbaPackageParts(graph, string.Empty, signatureIds == null ? mainIdentity : identity,
                        remove ? current.Id : null, remove || signatureIds != null ? null : identity, skipIds: signatureIds);
                    dead = projectGraph.Where(part => !kept.ContainsKey(part.Key)).Select(part => part.Value).ToArray();
                }
                bool removeRelationships = !remove && (signatureIds == null || signatureIds.Count > 0 && signatureIds.Count == projectLinks.Count);
                plan = new PackageVbaMutation(mainPart, current.Part, null, relationships, edgesToRemove, dead, removeRelationships,
                    preserveNonSignatureParts ? projectGraph.Values : null);
                return true;
            } catch (Exception exception) when (exception is InvalidDataException or XmlException or IOException or
                ArgumentException or NotSupportedException or InvalidOperationException or OverflowException) {
                return false;
            }
        }

        private static IReadOnlyList<OfficeOpcRelationship>? ReadPackageMutationRelationships(OfficeOpcEditablePackage package, string? owner, bool requireTransitionalTypes = false) {
            string part = OfficeOpcPartNames.GetRelationshipsPartName(owner);
            if (package.PartExists(part) && package.GetContentType(part) != OfficeOpcEditablePackage.RelationshipContentType) return null;
            IReadOnlyList<OfficeOpcRelationship>? links = ReadEditablePackageRelationships(package, owner, strictXml: true, requireTransitionalTypes: requireTransitionalTypes);
            if (links == null) return null;
            foreach (OfficeOpcRelationship link in links) {
                XmlConvert.VerifyNCName(link.Id);
                if (!Uri.TryCreate(link.Type, UriKind.Absolute, out _) || link.Target.Length == 0 ||
                    !Uri.TryCreate(link.Target, UriKind.RelativeOrAbsolute, out Uri? target) ||
                    !link.IsExternal && target.IsAbsoluteUri) return null;
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
            if (ReadPackageMutationRelationships(package, main) is not { } links) return false;
            plan = new PackageVbaMutation(main, part, ReservePackageMutationRelationshipId(links), null,
                Array.Empty<int>(), Array.Empty<string>());
            return true;
        }

        private static string ReservePackageMutationRelationshipId(IReadOnlyList<OfficeOpcRelationship> links) {
            string id;
            do {
                id = NewRelationshipId();
            } while (links.Any(link => link.Id == id));
            return id;
        }

        private static void EnsurePackageRelationshipContentType(OfficeOpcEditablePackage package, string owner) {
            string relationships = OfficeOpcPartNames.GetRelationshipsPartName(owner);
            if (package.GetContentType(relationships) != OfficeOpcEditablePackage.RelationshipContentType) {
                package.SetOverride(relationships, OfficeOpcEditablePackage.RelationshipContentType);
            }
        }

        private static Dictionary<string, string> ReachVbaPackageParts(Dictionary<string, List<PackageGraphRelationship>> graph,
            string start, string? skipOwner, string? skipId, string? stopAt, HashSet<string>? skipIds = null) {
            string startIdentity = PartIdentity(start);
            var found = new Dictionary<string, string>(OfficeAsciiIgnoreCaseComparer.Instance) { [startIdentity] = start };
            var pending = new Stack<string>();
            pending.Push(startIdentity);
            while (pending.Count > 0) {
                string source = pending.Pop();
                if (OfficeAsciiIgnoreCaseComparer.Instance.Equals(source, stopAt) || !graph.TryGetValue(source, out var edges)) continue;
                bool skip = (skipId != null || skipIds != null) && OfficeAsciiIgnoreCaseComparer.Instance.Equals(source, skipOwner);
                foreach (PackageGraphRelationship edge in edges) {
                    if (skip && (edge.Id == skipId || skipIds?.Contains(edge.Id) == true)) continue;
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
