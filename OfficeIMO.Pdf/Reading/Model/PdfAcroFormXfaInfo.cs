namespace OfficeIMO.Pdf;

/// <summary>
/// Lightweight AcroForm XFA packet metadata. OfficeIMO detects and reports XFA but does not render or fill XFA packets.
/// </summary>
public sealed class PdfAcroFormXfaInfo {
    internal PdfAcroFormXfaInfo(
        string objectKind,
        int? objectNumber,
        int packetCount,
        IReadOnlyList<string> packetNames,
        int streamCount,
        int stringCount,
        int dictionaryCount,
        int totalPayloadBytes,
        bool hasTemplatePacket,
        bool hasDatasetsPacket) {
        ObjectKind = objectKind;
        ObjectNumber = objectNumber;
        PacketCount = packetCount;
        PacketNames = packetNames;
        StreamCount = streamCount;
        StringCount = stringCount;
        DictionaryCount = dictionaryCount;
        TotalPayloadBytes = totalPayloadBytes;
        HasTemplatePacket = hasTemplatePacket;
        HasDatasetsPacket = hasDatasetsPacket;
    }

    /// <summary>Resolved PDF object kind used by the AcroForm /XFA entry, such as array, stream, string, or dictionary.</summary>
    public string ObjectKind { get; }

    /// <summary>Indirect object number for the /XFA value when it is stored by reference.</summary>
    public int? ObjectNumber { get; }

    /// <summary>Number of named XFA packet pairs when /XFA is represented as an alternating name/payload array.</summary>
    public int PacketCount { get; }

    /// <summary>Readable XFA packet names discovered from an array-form /XFA entry.</summary>
    public IReadOnlyList<string> PacketNames { get; }

    /// <summary>Number of stream payloads discovered in the /XFA value.</summary>
    public int StreamCount { get; }

    /// <summary>Number of string payloads discovered in the /XFA value.</summary>
    public int StringCount { get; }

    /// <summary>Number of dictionary payloads discovered in the /XFA value.</summary>
    public int DictionaryCount { get; }

    /// <summary>Total decoded byte length for readable stream or string payloads.</summary>
    public int TotalPayloadBytes { get; }

    /// <summary>True when a named template packet was found.</summary>
    public bool HasTemplatePacket { get; }

    /// <summary>True when a named datasets packet was found.</summary>
    public bool HasDatasetsPacket { get; }
}
