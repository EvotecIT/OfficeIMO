using OfficeIMO.Email;

namespace OfficeIMO.Email.AddressBook;

/// <summary>Integrity result for one OAB v4 Full Details component.</summary>
public sealed class OfflineAddressBookValidationResult {
    internal OfflineAddressBookValidationResult(
        OfflineAddressBookListInfo addressList,
        uint? calculatedChecksum,
        long recordsScanned,
        long recordsDecoded,
        long recordsSkipped,
        bool framingComplete,
        bool consumedDeclaredPayload,
        IReadOnlyList<EmailDiagnostic> diagnostics) {
        AddressList = addressList;
        CalculatedChecksum = calculatedChecksum;
        RecordsScanned = recordsScanned;
        RecordsDecoded = recordsDecoded;
        RecordsSkipped = recordsSkipped;
        FramingComplete = framingComplete;
        ConsumedDeclaredPayload = consumedDeclaredPayload;
        Diagnostics = diagnostics;
    }

    /// <summary>Validated address-list metadata.</summary>
    public OfflineAddressBookListInfo AddressList { get; }
    /// <summary>Checksum stored in the OAB header.</summary>
    public uint ExpectedChecksum => AddressList.Serial;
    /// <summary>Recalculated checksum, or null when checksum validation was disabled.</summary>
    public uint? CalculatedChecksum { get; }
    /// <summary>Whether the recalculated checksum matches, or null when it was not calculated.</summary>
    public bool? IsChecksumValid => CalculatedChecksum.HasValue
        ? CalculatedChecksum.Value == ExpectedChecksum
        : (bool?)null;
    /// <summary>Record frames walked.</summary>
    public long RecordsScanned { get; }
    /// <summary>Records decoded against the property schema.</summary>
    public long RecordsDecoded { get; }
    /// <summary>Recoverable value-level decode failures.</summary>
    public long RecordsSkipped { get; }
    /// <summary>Whether all declared records were framed within configured bounds.</summary>
    public bool FramingComplete { get; }
    /// <summary>Whether the declared record sequence ends exactly at the component boundary.</summary>
    public bool ConsumedDeclaredPayload { get; }
    /// <summary>Component-specific validation diagnostics.</summary>
    public IReadOnlyList<EmailDiagnostic> Diagnostics { get; }
    /// <summary>Whether the requested pass completed without checksum, framing, or decode errors.</summary>
    public bool IsValid => IsChecksumValid != false && FramingComplete && ConsumedDeclaredPayload &&
        RecordsSkipped == 0 && !Diagnostics.Any(diagnostic =>
            diagnostic.Severity == EmailDiagnosticSeverity.Error);
}
