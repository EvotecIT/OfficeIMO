namespace OfficeIMO.OpenDocument;

/// <summary>Serialized OpenDocument bytes and the entry-level report from one save operation.</summary>
public sealed class OdfSaveResult {
    internal OdfSaveResult(byte[] value, OdfSaveReport report) {
        Value = value ?? throw new ArgumentNullException(nameof(value));
        Report = report ?? throw new ArgumentNullException(nameof(report));
    }

    /// <summary>The exact bytes produced by the save operation.</summary>
    public byte[] Value { get; }

    /// <summary>Entries rewritten, copied, removed, or projected with loss.</summary>
    public OdfSaveReport Report { get; }

    /// <summary>True when the selected output form could not preserve every source entry losslessly.</summary>
    public bool HasLoss => Report.LossyEntries.Count > 0 || Report.RemovedEntries.Count > 0;

    /// <summary>Returns the serialized bytes.</summary>
    public byte[] RequireValue() => Value;

    /// <summary>Returns the serialized bytes or throws when the save removed or lossily projected content.</summary>
    public byte[] RequireNoLoss() {
        if (HasLoss) {
            throw new InvalidOperationException("OpenDocument save could not preserve every source entry losslessly.");
        }

        return Value;
    }
}
