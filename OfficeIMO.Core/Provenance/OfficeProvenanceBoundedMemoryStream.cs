using System;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Provenance;

/// <summary>Uses the shared serialization bound with provenance-specific failure classification.</summary>
internal sealed class OfficeProvenanceBoundedMemoryStream(long maximumBytes, int capacityHint = 0)
    : OfficeBoundedMemoryStream(maximumBytes, capacityHint) {
    protected override Exception CreateLimitException(long maximumBytes) =>
        OfficeProvenanceLimitException.CreateOutput(
            $"The rewritten asset exceeds the configured output limit of {maximumBytes} bytes.");
}
