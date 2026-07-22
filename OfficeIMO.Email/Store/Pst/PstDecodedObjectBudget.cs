using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

/// <summary>Tracks all materialized data that belongs to one embedded PST message object.</summary>
internal sealed class PstDecodedObjectBudget {
    private readonly long _maximumBytes;

    internal PstDecodedObjectBudget(long maximumBytes) {
        if (maximumBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maximumBytes));
        _maximumBytes = maximumBytes;
    }

    internal long ConsumedBytes { get; private set; }

    internal long RemainingBytes => _maximumBytes - ConsumedBytes;

    internal void Add(long bytes) {
        if (bytes < 0) throw new ArgumentOutOfRangeException(nameof(bytes));
        if (bytes > RemainingBytes) {
            long actual = bytes > long.MaxValue - ConsumedBytes
                ? long.MaxValue
                : ConsumedBytes + bytes;
            throw new EmailStoreLimitExceededException(
                nameof(EmailStoreReaderOptions.MaxAttachmentBytes),
                actual,
                _maximumBytes);
        }
        ConsumedBytes += bytes;
    }

    internal void AddProperties(IEnumerable<MapiProperty> properties) {
        foreach (MapiProperty property in properties) Add(property.RawData?.LongLength ?? 0);
    }

    internal void AddProjectedBodies(EmailDocument document) {
        AddString(document.Body.Text);
        AddString(document.Body.Html);
        AddString(document.Body.Rtf);
    }

    private void AddString(string? value) {
        if (!string.IsNullOrEmpty(value)) Add(checked((long)value!.Length * sizeof(char)));
    }
}
