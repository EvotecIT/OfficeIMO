using OfficeIMO.Core.Internal;

namespace OfficeIMO.Access {
    public sealed partial class AccessDocument {
        private void ApplyNativeApplicationPlan(AccessNativeWriter plan, byte[]? projectBytes, OfficeVbaProject? project, ISet<string> hosts, long maximumBytes, CancellationToken cancellationToken, string operation) {
            AccessNativeDatabase native = NativeDatabase ?? throw new NotSupportedException("Native application mutation requires a decoded source.");
            using OfficeBoundedMemoryStream output = new OfficeBoundedMemoryStream(maximumBytes);
            plan.Write(output, cancellationToken); byte[] candidateBytes = output.ToArray();
            using AccessDocument candidate = FromBytes(candidateBytes, new AccessLoadOptions {
                MaxInputBytes = maximumBytes, MaxPages = checked((int)(maximumBytes / 4096)),
                MaxCatalogObjects = native.MaxCatalogObjects, MaxMetadataBytes = native.MaxMetadataBytes,
                MaxValueBytes = native.MaxValueBytes, MaxRows = native.MaxRows, MaxChainLength = native.MaxChainLength,
                TableNames = native.SelectedTables?.ToArray()
            }, cancellationToken);
            if (project != null) {
                if (candidate.VbaProject.CatalogStatus != AccessCatalogStatus.Decoded || candidate.VbaProject.Modules.Count != project.Modules.Count)
                    throw new InvalidDataException("The native candidate did not retain its complete VBA inventory.");
                string[] ordinary = project.Modules.Where(x => !hosts.Contains(x.Name)).Select(x => x.Name).OrderBy(x => x, StringComparer.OrdinalIgnoreCase).ToArray();
                if (!ordinary.SequenceEqual(candidate.Catalog.Where(x => x.NativeType == -32761).Select(x => x.Name).OrderBy(x => x, StringComparer.OrdinalIgnoreCase), StringComparer.OrdinalIgnoreCase))
                    throw new InvalidDataException("The native ordinary-module catalog and VBA inventory disagree.");
            } else {
                AccessStorageStream[] before = ApplicationStreams.Where(x => x.Path.StartsWith("VBA/", StringComparison.OrdinalIgnoreCase)).ToArray();
                AccessStorageStream[] after = candidate.ApplicationStreams.Where(x => x.Path.StartsWith("VBA/", StringComparison.OrdinalIgnoreCase)).ToArray();
                if (before.Length != after.Length || before.Any(x => !after.Any(y => x.Path.Equals(y.Path, StringComparison.OrdinalIgnoreCase)
                    && x.Payload.GetBytes().SequenceEqual(y.Payload.GetBytes()))))
                    throw new InvalidDataException("A designer-only edit changed its unrelated VBA storage.");
            }
            VbaMutationState? previous = _vbaMutation; AccessVbaProjectInfo previousInfo = VbaProject;
            IReadOnlyList<AccessStorageStream> previousStreams = ApplicationStreams;
            AccessCatalogEntry[] previousCatalog = Catalog.Items.ToArray();
            AccessCatalogEntry[] nextModules = candidate.Catalog.Where(x => x.NativeType == -32761).Select(entry => {
                AccessCatalogEntry? retained = previousCatalog.FirstOrDefault(x => x.NativeType == -32761 && x.NativeId == entry.NativeId);
                if (retained?.Name == entry.Name) return retained;
                return new AccessCatalogEntry(this, entry.Name, entry.NativeType, entry.Flags, retained?.Id) {
                    NativeId = entry.NativeId, NativeParentId = entry.NativeParentId, NativeRecord = entry.NativeRecord,
                    Owner = entry.Owner, NativePayloads = entry.NativePayloads, Diagnostics = entry.Diagnostics
                };
            }).ToArray();
            var hostRefresh = PrepareVbaHostRefresh(candidate);
            cancellationToken.ThrowIfCancellationRequested();
            AccessNativeDatabase readProjection = candidate.NativeDatabase!; readProjection.BindReadProjection(this); candidate.NativeDatabase = null;
            var next = new VbaMutationState { Plan = plan, ProjectBytes = projectBytes, Sha256 = candidate.Inspection!.Sha256, ReadProjection = readProjection };
            Action? undoBindings = project == null ? null : BindVbaModuleIdentities(project, nextModules, HasActiveUpdate);
            hostRefresh.Apply();
            if (HasActiveUpdate && previous != null) _vbaUndoStates.Add(previous);
            _vbaMutation = next; VbaProject = candidate.VbaProject; ApplicationStreams = candidate.ApplicationStreams;
            Catalog.Items.RemoveAll(x => x.NativeType == -32761); Catalog.Items.AddRange(nextModules);
            Changed(() => {
                _vbaMutation = previous; VbaProject = previousInfo; ApplicationStreams = previousStreams;
                Catalog.Items.Clear(); Catalog.Items.AddRange(previousCatalog);
                hostRefresh.Undo();
                undoBindings?.Invoke();
                if (previous != null) _vbaUndoStates.Remove(previous);
                next.ReadProjection.Dispose();
            }, operation: operation);
        }

    }
}
