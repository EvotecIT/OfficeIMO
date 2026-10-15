using System;
using System.IO;
using OfficeIMO.Core.Internal;
using OfficeIMO.Core.Internal.Cabinet;

namespace OfficeIMO.Archives;

internal sealed partial class ArchiveLimits {
    /// <summary>Translates only structured native admission failures; format faults retain their corruption identity.</summary>
    internal Exception Translate(OfficeCabinetException error, long? expandedMaximum = null,
        int? entryMaximum = null, int? directoryMaximum = null) {
        switch (error.Code) {
            case "CAB_INPUT_LIMIT": return Exceeded(nameof(ArchiveExtractionOptions.MaxInputBytes), Math.Min(Input, int.MaxValue), error);
            case "CAB_ENTRY_LIMIT": return Exceeded(nameof(ArchiveExtractionOptions.MaxEntryBytes), Math.Min(Entry, int.MaxValue), error);
            case "CAB_ENTRY_COUNT_LIMIT": return Exceeded(nameof(ArchiveExtractionOptions.MaxEntryCount), entryMaximum ?? Entries, error);
            case "CAB_EXPANDED_LIMIT": return Exceeded(nameof(ArchiveExtractionOptions.MaxExpandedBytes), expandedMaximum ?? Expanded, error);
            case "CAB_FOLDER_BUFFER_LIMIT": return Exceeded(nameof(ArchiveExtractionOptions.MaxExpandedBytes), int.MaxValue, error);
            case "CAB_WORK_LIMIT": return Exceeded(nameof(ArchiveExtractionOptions.MaxDecompressionWorkBytes), DecompressionWork.MaximumBytes, error);
            case "CAB_DIRECTORY_LIMIT" when directoryMaximum.HasValue:
                return Exceeded(nameof(ArchiveInspectionOptions.MaxDirectoryRecords), directoryMaximum.Value, error);
            case "CAB_COMPRESSION": case "CAB_MULTIPART": case "CAB_LZX_WINDOW": case "CAB_NAME_ENCODING":
                return new NotSupportedException(error.Code + ": " + error.Message, error);
            default: return new InvalidDataException(error.Code + ": " + error.Message, error);
        }
    }

    internal static Exception Translate(OfficeCompoundStreamLimitExceededException error) {
        string name = error.LimitName switch {
            "MaxDirectoryEntries" => nameof(ArchiveInspectionOptions.MaxDirectoryRecords),
            "MaxStreamCount" => "MaxStorageStreamCount",
            "MaxStreamBytes" or "MaxTotalStreamBytes" or "MaxAllocationTableBytes" => nameof(ArchiveExtractionOptions.MaxInputBytes),
            _ => error.LimitName
        };
        return Exceeded(name, error.MaximumValue, error);
    }
}
