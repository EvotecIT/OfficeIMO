using System;

namespace OfficeIMO.Core.Internal.Cabinet;

/// <summary>Bounds directory parsing across standalone and embedded cabinets without decoding payloads.</summary>
internal sealed class OfficeCabinetDirectoryBudget {
    private long _remaining;
    internal OfficeCabinetDirectoryBudget(long maximumRecords) {
        if (maximumRecords < 1) throw new ArgumentOutOfRangeException(nameof(maximumRecords));
        _remaining = maximumRecords;
    }
    internal void Consume(long records) {
        if (records < 0 || records > _remaining)
            throw new OfficeCabinetException("CAB_DIRECTORY_LIMIT", "CAB directory work exceeds the configured record limit.");
        _remaining -= records;
    }
}
