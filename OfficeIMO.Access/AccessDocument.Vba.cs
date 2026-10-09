using OfficeIMO.Core.Internal;

namespace OfficeIMO.Access {
    public sealed partial class AccessDocument {
        private VbaMutationState? _vbaMutation;
        private string? _savedSourceIdentity;
        private readonly List<VbaMutationState> _vbaUndoStates = new List<VbaMutationState>();
        private sealed class VbaMutationState {
            internal AccessNativeWriter Plan = null!;
            internal byte[] ProjectBytes = null!;
            internal string Sha256 = null!;
            internal AccessNativeDatabase ReadProjection = null!;
        }
        /// <summary>Loads a detached editable VBA project from the decoded native application storage.</summary>
        /// <remarks>Editing the returned project does not change this database. Missing or opaque source is rejected by the shared project editor; inspection remains available through <see cref="VbaProject"/>.</remarks>
        public OfficeVbaProject GetVbaProject(OfficeVbaReadOptions? options = null, CancellationToken cancellationToken = default) {
            EnsureNotDisposed(); cancellationToken.ThrowIfCancellationRequested();
            if (VbaProject.CatalogStatus != AccessCatalogStatus.Decoded)
                throw new NotSupportedException("The Access VBA project is not decoded. Its native storage remains preserve-only.");
            options ??= new OfficeVbaReadOptions();
            if (options.MaximumProjectBytes < 1 || options.MaximumExpandedBytes < 1) throw new ArgumentOutOfRangeException(nameof(options));
            OfficeVbaProject project = OfficeVbaProject.Load(_vbaMutation?.ProjectBytes ?? NativeDatabase!.DetachVbaProject(options.MaximumProjectBytes, cancellationToken), options);
            BindVbaModuleIdentities(project, Catalog.Items, undoable: false);
            return project;
        }

        /// <summary>Stages a shared VBA project in a qualified native MDB/ACCDB snapshot. Call Save to persist it.</summary>
        /// <remarks>Native table schemas and unrelated application objects remain unchanged. No VBA is executed or compiled. Storage, host metadata and permission layouts outside the qualified boundary are rejected before document mutation.</remarks>
        public void SetVbaProject(OfficeVbaProject project, OfficeVbaWriteOptions? options = null, CancellationToken cancellationToken = default) {
            if (project == null) throw new ArgumentNullException(nameof(project));
            EnsureMutationAllowed(); cancellationToken.ThrowIfCancellationRequested();
            if (NativeDatabase == null || CatalogStatus != AccessCatalogStatus.Decoded || VbaProject.CatalogStatus != AccessCatalogStatus.Decoded)
                throw new NotSupportedException("VBA application requires a decoded native database and project.");
            ValidateSourceIdentity(cancellationToken);
            AccessNativeDatabase source = _vbaMutation?.ReadProjection ?? NativeDatabase;
            IReadOnlyDictionary<OfficeVbaModule, string> appliedNames = ResolveVbaModuleNames(project);
            options ??= new OfficeVbaWriteOptions();
            byte[] bytes = project.Write(options).GetBytes();
            var limits = new OfficeCompoundReadOptions(maxStreamBytes: options.MaximumProjectBytes, maxTotalStreamBytes: options.MaximumProjectBytes);
            if (!OfficeCompoundFileReader.TryRead(bytes, limits, cancellationToken, out OfficeCompoundFile? compound, out string? error) || compound == null)
                throw new InvalidDataException(error ?? "The VBA writer did not produce a valid project.");
            const string prefix = "VBA/VBAProject/";
            AccessStorageStream[] original = ApplicationStreams.Where(x => x.Path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)).ToArray();
            if (original.Length == compound.Streams.Count && original.All(x => compound.Streams.TryGetValue(x.Path.Substring(prefix.Length), out byte[]? value) && x.Payload.GetBytes().SequenceEqual(value))) return;
            OfficeVbaProject current = original.Length == 0 ? OfficeVbaProject.Create(project.Name, project.CodePage)
                : GetVbaProject(new OfficeVbaReadOptions { MaximumProjectBytes = options.MaximumProjectBytes, MaximumExpandedBytes = options.MaximumExpandedBytes }, cancellationToken);
            if (current.IsProtected || project.IsProtected) throw new InvalidOperationException("Access VBA editing does not bypass project protection.");
            HashSet<string> hosts = source.GetVbaHostNames(current);
            ValidateAccessVbaHostModules(current, project, hosts, appliedNames);
            foreach (AccessApplicationObject host in Forms.Concat(Reports)) {
                string name = (host.CatalogEntry.NativeType == -32768 ? "Form_" : "Report_") + host.Name;
                if (!hosts.Contains(name) && project.Modules.Any(x => x.Kind == OfficeVbaModuleKind.Document && x.Name.Equals(name, StringComparison.OrdinalIgnoreCase)))
                    throw new NotSupportedException("Adding new form/report code-behind requires qualified host metadata authoring.");
            }
            if (options.MaximumRecoveryBytes < 1) throw new ArgumentOutOfRangeException(nameof(options));
            long recovery = NativeDatabase.Snapshot().Length;
            foreach (VbaMutationState state in _vbaUndoStates.Concat(_vbaMutation != null ? new[] { _vbaMutation } : Array.Empty<VbaMutationState>()))
                recovery = checked(recovery + state.Plan.Length * 2 + state.ProjectBytes.Length);
            if (recovery > options.MaximumRecoveryBytes) throw new InvalidDataException("Native Access recovery snapshots exceed MaximumRecoveryBytes.");
            // Access uses a direct native signature carrier rather than OPC relationships.
            // Unknown VBA-side siblings must remain preserve-only until their carrier is qualified.
            var modulePaths = new HashSet<string>(VbaProject.Modules.Select(x => x.StoragePath), StringComparer.OrdinalIgnoreCase);
            if (ApplicationStreams.Any(x => x.Path.StartsWith("VBA/", StringComparison.OrdinalIgnoreCase)
                && !x.Path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) && !x.Path.Equals("VBA/AcessVBAData", StringComparison.OrdinalIgnoreCase)
                || x.Path.IndexOf("DigitalSignature", StringComparison.OrdinalIgnoreCase) >= 0 && !modulePaths.Contains(x.Path)))
                throw new NotSupportedException("An unqualified VBA or signature carrier prevents native Access editing.");
            long maximumBytes = Math.Min(int.MaxValue, checked(Math.Max(_inputLimit, source.Snapshot().Length) + options.MaximumProjectBytes));
            AccessNativeWriter plan;
            using (source.PreserveMutationMetadataAllowance())
                plan = source.BuildVbaMutation(project, compound, maximumBytes, cancellationToken, hosts, appliedNames);
            using OfficeBoundedMemoryStream output = new OfficeBoundedMemoryStream(maximumBytes);
            plan.Write(output, cancellationToken); byte[] candidateBytes = output.ToArray();
            using AccessDocument candidate = FromBytes(candidateBytes, new AccessLoadOptions {
                MaxInputBytes = maximumBytes, MaxPages = checked((int)(maximumBytes / 4096)),
                MaxCatalogObjects = NativeDatabase.MaxCatalogObjects, MaxMetadataBytes = NativeDatabase.MaxMetadataBytes,
                MaxValueBytes = NativeDatabase.MaxValueBytes, MaxRows = NativeDatabase.MaxRows, MaxChainLength = NativeDatabase.MaxChainLength,
                TableNames = NativeDatabase.SelectedTables?.ToArray()
            }, cancellationToken);
            if (candidate.VbaProject.CatalogStatus != AccessCatalogStatus.Decoded || candidate.VbaProject.Modules.Count != project.Modules.Count)
                throw new InvalidDataException("The native candidate did not retain its complete VBA inventory.");
            string[] ordinary = project.Modules.Where(x => !hosts.Contains(x.Name)).Select(x => x.Name).OrderBy(x => x, StringComparer.OrdinalIgnoreCase).ToArray();
            if (!ordinary.SequenceEqual(candidate.Catalog.Where(x => x.NativeType == -32761).Select(x => x.Name).OrderBy(x => x, StringComparer.OrdinalIgnoreCase), StringComparer.OrdinalIgnoreCase))
                throw new InvalidDataException("The native ordinary-module catalog and VBA inventory disagree.");
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
            cancellationToken.ThrowIfCancellationRequested();
            AccessNativeDatabase readProjection = candidate.NativeDatabase!; readProjection.BindReadProjection(this); candidate.NativeDatabase = null;
            var next = new VbaMutationState { Plan = plan, ProjectBytes = bytes, Sha256 = candidate.Inspection!.Sha256, ReadProjection = readProjection };
            Action? undoBindings = BindVbaModuleIdentities(project, nextModules, HasActiveUpdate);
            if (HasActiveUpdate && previous != null) _vbaUndoStates.Add(previous);
            _vbaMutation = next; VbaProject = candidate.VbaProject; ApplicationStreams = candidate.ApplicationStreams;
            Catalog.Items.RemoveAll(x => x.NativeType == -32761); Catalog.Items.AddRange(nextModules);
            Changed(() => {
                _vbaMutation = previous; VbaProject = previousInfo; ApplicationStreams = previousStreams;
                Catalog.Items.Clear(); Catalog.Items.AddRange(previousCatalog);
                undoBindings?.Invoke();
                if (previous != null) _vbaUndoStates.Remove(previous);
                next.ReadProjection.Dispose();
            }, operation: "vba.apply");
        }

        internal AccessNativeTable? ResolveNativeReadTable(AccessNativeTable? original, CancellationToken cancellation = default) {
            if (original == null || _vbaMutation == null || original.Name != "MSysObjects" && original.Name != "MSysACEs"
                && original.Name != "MSysAccessStorage" && original.Name != "MSysAccessObjects") return original;
            return _vbaMutation.ReadProjection.Definition(original.DefinitionPage, original.Name, cancellation);
        }

        private static void ValidateAccessVbaHostModules(OfficeVbaProject original, OfficeVbaProject replacement, ISet<string> boundHosts,
            IReadOnlyDictionary<OfficeVbaModule, string>? appliedNames) {
            foreach (OfficeVbaModule module in replacement.Modules) {
                string identity = appliedNames != null && appliedNames.TryGetValue(module, out string? applied) ? applied : module.IsNew ? module.Name : module.OriginalName;
                OfficeVbaModule? previous = original.Modules.FirstOrDefault(x => x.Name.Equals(identity, StringComparison.OrdinalIgnoreCase));
                if (previous != null && previous.Kind != module.Kind) throw new ArgumentException("Replacing an Access module cannot change its native persistence kind.", nameof(replacement));
            }
            OfficeVbaModule[] hosts = original.Modules.Where(x => boundHosts.Contains(x.Name)).ToArray();
            OfficeVbaModule[] updated = replacement.Modules.Where(x => boundHosts.Contains(x.Name)).ToArray();
            if (hosts.Length != updated.Length || hosts.Any(host => !updated.Any(x => x.Kind == host.Kind && x.Name.Equals(host.Name, StringComparison.OrdinalIgnoreCase)
                && string.Equals(OfficeVbaText.GetBaseIdentity(x.Source), OfficeVbaText.GetBaseIdentity(host.Source), StringComparison.OrdinalIgnoreCase))))
                throw new ArgumentException("Access form/report module identities must remain bound to their existing host.", nameof(replacement));
            if (replacement.Modules.Any(x => x.Kind != OfficeVbaModuleKind.Standard && x.Kind != OfficeVbaModuleKind.Class && !boundHosts.Contains(x.Name)))
                throw new NotSupportedException("Access VBA persistence supports ordinary and qualified form/report class modules.");
        }
    }
}
