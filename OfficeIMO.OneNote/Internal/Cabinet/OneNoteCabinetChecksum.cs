using OfficeIMO.Core.Internal.Cabinet;

namespace OfficeIMO.OneNote;

/// <summary>Retains the OneNote writer adapter over the shared MS-CAB checksum.</summary>
internal static class OneNoteCabinetChecksum {
    internal static uint Compute(byte[] data, int offset, int count) => OfficeCabinetChecksum.Compute(data, offset, count);
}
