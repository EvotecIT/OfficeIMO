namespace OfficeIMO.Core.Internal.Cabinet;

/// <summary>One immutable authored CAB file record with shared-folder sizes and codec support evidence.</summary>
internal sealed class OfficeCabinetMemberInfo {
    internal OfficeCabinetMemberInfo(int ordinal, string name, long length, int folderOrdinal,
        ushort compression, long compressedBytes, long expandedBytes, int frameCount, int checksumFrameCount,
        long workBytes, string? unsupportedCode) {
        Ordinal = ordinal; Name = name; Length = length; FolderOrdinal = folderOrdinal;
        Compression = compression; FolderCompressedBytes = compressedBytes; FolderExpandedBytes = expandedBytes;
        FolderFrameCount = frameCount; FolderChecksumFrameCount = checksumFrameCount;
        FolderDecompressionWorkBytes = workBytes;
        UnsupportedCode = unsupportedCode;
    }
    internal int Ordinal { get; }
    internal string Name { get; }
    internal long Length { get; }
    internal int FolderOrdinal { get; }
    internal ushort Compression { get; }
    internal long FolderCompressedBytes { get; }
    internal long FolderExpandedBytes { get; }
    internal int FolderFrameCount { get; }
    internal int FolderChecksumFrameCount { get; }
    internal long FolderDecompressionWorkBytes { get; }
    internal string? UnsupportedCode { get; }
}
