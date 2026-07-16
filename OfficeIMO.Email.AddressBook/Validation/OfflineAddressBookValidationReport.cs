namespace OfficeIMO.Email.AddressBook;

/// <summary>Aggregate result for selected OAB Full Details components.</summary>
public sealed class OfflineAddressBookValidationReport {
    internal OfflineAddressBookValidationReport(IReadOnlyList<OfflineAddressBookValidationResult> results) {
        Results = results;
        BytesHashed = results.Where(result => result.CalculatedChecksum.HasValue)
            .Sum(result => Math.Max(0, result.AddressList.SourceLength - 12));
        RecordsScanned = results.Sum(result => result.RecordsScanned);
        RecordsDecoded = results.Sum(result => result.RecordsDecoded);
        RecordsSkipped = results.Sum(result => result.RecordsSkipped);
    }

    /// <summary>Per-component results.</summary>
    public IReadOnlyList<OfflineAddressBookValidationResult> Results { get; }
    /// <summary>Payload bytes hashed.</summary>
    public long BytesHashed { get; }
    /// <summary>Record frames walked.</summary>
    public long RecordsScanned { get; }
    /// <summary>Records decoded.</summary>
    public long RecordsDecoded { get; }
    /// <summary>Recoverable decode failures.</summary>
    public long RecordsSkipped { get; }
    /// <summary>Whether every selected component passed the requested validation.</summary>
    public bool IsValid => Results.All(result => result.IsValid);
}
