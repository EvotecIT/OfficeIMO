namespace OfficeIMO.Access {
    public sealed partial class AccessDocument {
        private void ValidateNativeApplicationMutation(OfficeVbaWriteOptions options) {
            const string prefix = "VBA/VBAProject/";
            if (options.MaximumProjectBytes < 1 || options.MaximumExpandedBytes < 1 || options.MaximumRecoveryBytes < 1) throw new ArgumentOutOfRangeException(nameof(options));
            long recovery = (NativeDatabase ?? throw new NotSupportedException("Native application mutation requires a decoded source.")).Snapshot().Length;
            foreach (VbaMutationState state in _vbaUndoStates.Concat(_vbaMutation != null ? new[] { _vbaMutation } : Array.Empty<VbaMutationState>()))
                recovery = checked(recovery + state.Plan.Length * 2 + (state.ProjectBytes?.Length ?? 0));
            if (recovery > options.MaximumRecoveryBytes) throw new InvalidDataException("Native Access recovery snapshots exceed MaximumRecoveryBytes.");
            // Access uses a direct native signature carrier rather than OPC relationships.
            // Unknown VBA-side siblings must remain preserve-only until their carrier is qualified.
            var modulePaths = new HashSet<string>(VbaProject.Modules.Select(x => x.StoragePath), StringComparer.OrdinalIgnoreCase);
            if (ApplicationStreams.Any(x => x.Path.StartsWith("VBA/", StringComparison.OrdinalIgnoreCase)
                && !x.Path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) && !x.Path.Equals("VBA/AcessVBAData", StringComparison.OrdinalIgnoreCase)
                || x.Path.IndexOf("DigitalSignature", StringComparison.OrdinalIgnoreCase) >= 0 && !modulePaths.Contains(x.Path)))
                throw new NotSupportedException("An unqualified VBA or signature carrier prevents native Access editing.");
        }
    }
}
