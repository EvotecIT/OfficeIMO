namespace OfficeIMO.OpenDocument;

/// <summary>Serialized OpenDocument bytes and the entry-level report from one save operation.</summary>
public sealed class OdfSaveResult {
    private readonly byte[] _value;

    internal OdfSaveResult(byte[] bytes, OdfSaveReport report) {
        _value = bytes == null ? throw new ArgumentNullException(nameof(bytes)) : (byte[])bytes.Clone();
        Report = report ?? throw new ArgumentNullException(nameof(report));
    }

    /// <summary>The exact bytes produced by the save operation.</summary>
    public byte[] Value => (byte[])_value.Clone();

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
