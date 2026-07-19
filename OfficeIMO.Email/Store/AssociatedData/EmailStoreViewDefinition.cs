using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

/// <summary>Validated Outlook named-view envelope with lossless binary streams.</summary>
public sealed class EmailStoreViewDefinition {
    private const string ExpectedMessageClass = "IPM.Microsoft.FolderDesign.NamedView";

    internal EmailStoreViewDefinition(EmailDocument document) {
        MessageClass = document.MessageClass;
        Name = document.Mapi.GetValueOrDefault(MapiKnownProperties.PidTag.ViewDescriptorName);
        DescriptorVersion = document.Mapi.GetNullableValue(MapiKnownProperties.PidTag.ViewDescriptorVersion);
        DescriptorFlags = document.Mapi.GetNullableValue(MapiKnownProperties.PidTag.ViewDescriptorFlags);
        LinkTo = Copy(document.Mapi.GetValueOrDefault(MapiKnownProperties.PidTag.ViewDescriptorLinkTo));
        ViewFolder = Copy(document.Mapi.GetValueOrDefault(MapiKnownProperties.PidTag.ViewDescriptorViewFolder));
        Binary = Copy(document.Mapi.GetValueOrDefault(MapiKnownProperties.PidTag.ViewDescriptorBinary));
        Strings = Copy(document.Mapi.GetValueOrDefault(MapiKnownProperties.PidTag.ViewDescriptorStrings));
        if (Binary != null && Binary.Length >= 64) {
            BinaryVersion = ReadUInt32(Binary, 8);
            SortFlags = ReadUInt32(Binary, 12);
            ColumnCount = ReadUInt32(Binary, 20);
            SortColumnIndex = ReadUInt32(Binary, 24);
            GroupByColumnCount = ReadUInt32(Binary, 28);
            GroupSortFlags = ReadUInt32(Binary, 32);
        }
    }

    /// <summary>Actual associated-message class.</summary>
    public string? MessageClass { get; }
    /// <summary>Non-empty view name when valid.</summary>
    public string? Name { get; }
    /// <summary>PidTagViewDescriptorVersion; version 8 is defined.</summary>
    public int? DescriptorVersion { get; }
    /// <summary>PidTagViewDescriptorFlags.</summary>
    public int? DescriptorFlags { get; }
    /// <summary>Raw descriptor link target.</summary>
    public byte[]? LinkTo { get; }
    /// <summary>Raw view-folder identifier.</summary>
    public byte[]? ViewFolder { get; }
    /// <summary>Exact PidTagViewDescriptorBinary stream.</summary>
    public byte[]? Binary { get; }
    /// <summary>Exact PidTagViewDescriptorStrings stream.</summary>
    public byte[]? Strings { get; }
    /// <summary>Version from the binary descriptor header.</summary>
    public uint? BinaryVersion { get; }
    /// <summary>Sort-order flags from the binary header.</summary>
    public uint? SortFlags { get; }
    /// <summary>Declared number of column packets.</summary>
    public uint? ColumnCount { get; }
    /// <summary>Index of the sorted column, or 0xFFFFFFFF for conversation order.</summary>
    public uint? SortColumnIndex { get; }
    /// <summary>Declared number of group-by columns.</summary>
    public uint? GroupByColumnCount { get; }
    /// <summary>Ascending/descending flags for group-by columns.</summary>
    public uint? GroupSortFlags { get; }

    /// <summary>
    /// True when the documented message class, names, versions, and fixed header are valid. Column and restriction
    /// packets remain opaque and lossless until the shared MAPI restriction AST owns their decoding.
    /// </summary>
    public bool IsProtocolEnvelopeValid =>
        string.Equals(MessageClass, ExpectedMessageClass, StringComparison.OrdinalIgnoreCase) &&
        !string.IsNullOrWhiteSpace(Name) && DescriptorVersion == 8 && BinaryVersion == 8 &&
        Binary != null && Binary.Length >= 64 &&
        (SortFlags == 0 || SortFlags == 2) && GroupByColumnCount <= 4 &&
        ColumnCount.HasValue && GroupByColumnCount <= ColumnCount;

    private static uint ReadUInt32(byte[] bytes, int offset) =>
        (uint)(bytes[offset] | (bytes[offset + 1] << 8) | (bytes[offset + 2] << 16) |
               (bytes[offset + 3] << 24));
    private static byte[]? Copy(byte[]? value) => value == null ? null : (byte[])value.Clone();
}
