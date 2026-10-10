namespace OfficeIMO.Access {
    public sealed partial class AccessDocument {
        internal void SetCodeBehind(AccessApplicationObject host, string source, OfficeVbaWriteOptions? options, CancellationToken cancellation) {
            if (source == null) throw new ArgumentNullException(nameof(source));
            EnsureMutationAllowed(); cancellation.ThrowIfCancellationRequested();
            ValidateVbaHost(host);
            AccessNativeDatabase native = _vbaMutation?.ReadProjection ?? NativeDatabase
                ?? throw new NotSupportedException("Code-behind requires an existing native application.");
            options ??= new OfficeVbaWriteOptions(); ValidateNativeApplicationMutation(options);
            byte[] properties = native.GetVbaHostProperties(host);
            int moduleOffset = AccessNativeDatabase.HostModuleValueOffset(properties);
            string name = (host.CatalogEntry.NativeType == -32768 ? "Form_" : "Report_") + host.Name;
            OfficeVbaProject project = ApplicationStreams.Any(x => x.Path.StartsWith("VBA/VBAProject/", StringComparison.OrdinalIgnoreCase))
                ? GetVbaProject(new OfficeVbaReadOptions { MaximumProjectBytes = options.MaximumProjectBytes, MaximumExpandedBytes = options.MaximumExpandedBytes }, cancellation)
                : OfficeVbaProject.Create("Database");
            if (properties[moduleOffset] == 1) {
                native.GetVbaHostNames(project);
                project.SetModuleSource(name, source);
                SetVbaProjectCore(project, options, cancellation, null, null);
                return;
            }
            if (project.Modules.Any(x => x.Name.Equals(name, StringComparison.OrdinalIgnoreCase)))
                throw new InvalidDataException("The form/report name already has an unbound VBA module.");
            project.AddDocumentModuleWithIdentity(name, source, "0" + Guid.NewGuid().ToString("B").ToUpperInvariant(), documentClass: true);
            properties[moduleOffset] = 1;
            SetVbaProjectCore(project, options, cancellation, new HashSet<string>(StringComparer.OrdinalIgnoreCase) { name },
                new Dictionary<string, byte[]> { [host.StoragePath + "PropData"] = properties });
        }

        private void ValidateVbaHost(AccessApplicationObject host) {
            host.EnsureAttached();
            if (host.Document != this || !Forms.Concat(Reports).Contains(host) || host.StoragePath == null)
                throw new NotSupportedException("Only qualified existing native forms and reports accept code-behind or event edits.");
        }

        private (Action Apply, Action Undo) PrepareVbaHostRefresh(AccessDocument candidate) {
            var updated = candidate.Forms.Concat(candidate.Reports).ToDictionary(x => (x.CatalogEntry.NativeType, x.CatalogEntry.NativeId));
            // Nullable paths intentionally represent preserve-only hosts. Native catalog
            // identity remains stable even when two unrelated objects have no mapping.
            var originals = Forms.Concat(Reports).Select(host => {
                if (!updated.TryGetValue((host.CatalogEntry.NativeType, host.CatalogEntry.NativeId), out AccessApplicationObject? next))
                    throw new InvalidDataException("The native candidate did not retain a form/report catalog identity.");
                return new { Host = host, host.Streams, host.Definition, host.Diagnostics, NextStreams = next.Streams, NextDefinition = next.Definition,
                    NextDiagnostics = Array.AsReadOnly(next.Diagnostics.Select(x => new AccessDiagnostic(x.Code, x.Message, host.Id)).ToArray()) };
            }).ToArray();
            return (() => {
                foreach (var state in originals) {
                    state.Host.Streams = state.NextStreams; state.Host.Definition = state.NextDefinition; state.Host.Diagnostics = state.NextDiagnostics;
                }
            }, () => {
                foreach (var state in originals) {
                    state.Host.Streams = state.Streams; state.Host.Definition = state.Definition; state.Host.Diagnostics = state.Diagnostics;
                }
            });
        }
    }
}
