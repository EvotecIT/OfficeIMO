namespace OfficeIMO.Access {
    public sealed partial class AccessDocument {
        internal void SetEventBinding(AccessApplicationObject host, AccessEventKind eventKind, string? expression, string? controlName,
            OfficeVbaWriteOptions? options, CancellationToken cancellation) {
            EnsureMutationAllowed(); cancellation.ThrowIfCancellationRequested(); ValidateVbaHost(host); ValidateSourceIdentity(cancellation);
            if (expression != null && (expression.IndexOf('\0') >= 0 || expression.Length > 4096))
                throw new ArgumentException("An event binding must be a bounded non-null expression.", nameof(expression));
            if (string.Equals(expression, "[Embedded Macro]", StringComparison.OrdinalIgnoreCase))
                throw new NotSupportedException("Use the preserved embedded macro definition; this operation does not create embedded actions.");
            if ((eventKind == AccessEventKind.Open) != (controlName == null))
                throw new ArgumentException("Open binds the form/report; Click and AfterUpdate require a named control.", nameof(controlName));
            (ushort code, uint id) = eventKind switch {
                AccessEventKind.Open => ((ushort)77, 227U), AccessEventKind.Click => ((ushort)126, 223U),
                AccessEventKind.AfterUpdate => ((ushort)86, 229U), _ => throw new ArgumentOutOfRangeException(nameof(eventKind))
            };
            AccessNativeDatabase native = _vbaMutation?.ReadProjection ?? NativeDatabase
                ?? throw new NotSupportedException("Event authoring requires an existing native application.");
            if (string.Equals(expression, "[Event Procedure]", StringComparison.OrdinalIgnoreCase)) {
                byte[] properties = native.GetVbaHostProperties(host);
                if (properties[AccessNativeDatabase.HostModuleValueOffset(properties)] != 1)
                    throw new InvalidOperationException("Create code-behind before binding an event procedure.");
            }
            options ??= new OfficeVbaWriteOptions(); ValidateNativeApplicationMutation(options);
            byte[] blob = native.GetEventDesigner(host);
            byte[] changed = AccessNativeDesigner.ReplaceEvent(blob, controlName, code, id, expression,
                native.MaxCatalogObjects, native.MaxMetadataBytes, cancellation);
            if (blob.SequenceEqual(changed)) return;
            long maximumBytes = Math.Min(int.MaxValue, checked(Math.Max(_inputLimit, native.Snapshot().Length) + options.MaximumProjectBytes));
            AccessNativeWriter plan = native.BuildApplicationStreamReplacement(new Dictionary<string, byte[]> { [host.StoragePath + "Blob"] = changed }, maximumBytes, cancellation);
            ApplyNativeApplicationPlan(plan, _vbaMutation?.ProjectBytes, null, new HashSet<string>(), maximumBytes, cancellation, "event.set");
        }
    }
}
