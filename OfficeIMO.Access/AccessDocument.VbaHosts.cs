namespace OfficeIMO.Access {
    public sealed partial class AccessDocument {
        internal void SetCodeBehind(AccessApplicationObject host, string source, OfficeVbaWriteOptions? options, CancellationToken cancellation) {
            if (source == null) throw new ArgumentNullException(nameof(source));
            EnsureMutationAllowed(); cancellation.ThrowIfCancellationRequested();
            ValidateVbaHost(host);
            AccessNativeDatabase native = _vbaMutation?.ReadProjection ?? NativeDatabase
                ?? throw new NotSupportedException("Code-behind requires an existing native application.");
            byte[] properties = native.GetVbaHostProperties(host);
            int moduleOffset = AccessNativeDatabase.HostModuleValueOffset(properties);
            string name = (host.CatalogEntry.NativeType == -32768 ? "Form_" : "Report_") + host.Name;
            OfficeVbaProject project = ApplicationStreams.Any(x => x.Path.StartsWith("VBA/VBAProject/", StringComparison.OrdinalIgnoreCase))
                ? GetVbaProject(cancellationToken: cancellation) : OfficeVbaProject.Create("Database");
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

        private Action RefreshVbaHostViews(AccessDocument candidate) {
            var originals = Forms.Concat(Reports).Select(x => new { Host = x, x.Streams, x.Definition, x.Diagnostics }).ToArray();
            foreach (var state in originals) {
                AccessApplicationObject updated = candidate.Forms.Concat(candidate.Reports).Single(x => x.StoragePath == state.Host.StoragePath);
                state.Host.Streams = updated.Streams; state.Host.Definition = updated.Definition;
                state.Host.Diagnostics = Array.AsReadOnly(updated.Diagnostics.Select(x => new AccessDiagnostic(x.Code, x.Message, state.Host.Id)).ToArray());
            }
            return () => {
                foreach (var state in originals) {
                    state.Host.Streams = state.Streams; state.Host.Definition = state.Definition; state.Host.Diagnostics = state.Diagnostics;
                }
            };
        }
    }
}
