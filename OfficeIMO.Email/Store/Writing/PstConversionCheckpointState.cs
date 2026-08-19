namespace OfficeIMO.Email.Store;

internal sealed class PstConversionCheckpointState {
    private static readonly byte[] Magic = Encoding.ASCII.GetBytes("OIMOMIG1");
    private const int MaximumEntryCount = 1_000_000;

    internal EmailStoreFormat SourceFormat { get; set; }
    internal long SourceLength { get; set; }
    internal string CatalogFingerprint { get; set; } = string.Empty;
    internal string DurableFingerprint { get; set; } = string.Empty;
    internal string OptionsFingerprint { get; set; } = string.Empty;
    internal string WriterDestination { get; set; } = string.Empty;
    internal int InspectedItems { get; set; }
    internal int ConvertedItems { get; set; }
    internal int SkippedItems { get; set; }
    internal string? MappingPath { get; set; }
    internal long MappingLength { get; set; }
    internal int MappingCount { get; set; }
    internal Dictionary<string, string> FolderMap { get; } = new Dictionary<string, string>(StringComparer.Ordinal);
    internal List<EmailStoreDiagnostic> Diagnostics { get; } = new List<EmailStoreDiagnostic>();

    internal EmailStoreSourceIdentity SourceIdentity => new EmailStoreSourceIdentity(
        SourceFormat, SourceLength, CatalogFingerprint, DurableFingerprint);

    internal byte[] Serialize() {
        using (var stream = new MemoryStream())
        using (var writer = new BinaryWriter(stream, Encoding.UTF8, leaveOpen: true)) {
            writer.Write(Magic);
            writer.Write((int)SourceFormat);
            writer.Write(SourceLength);
            writer.Write(CatalogFingerprint);
            writer.Write(DurableFingerprint);
            writer.Write(OptionsFingerprint);
            writer.Write(WriterDestination);
            writer.Write(InspectedItems);
            writer.Write(ConvertedItems);
            writer.Write(SkippedItems);
            WriteNullableString(writer, MappingPath);
            writer.Write(MappingLength);
            writer.Write(MappingCount);
            writer.Write(FolderMap.Count);
            foreach (KeyValuePair<string, string> entry in FolderMap.OrderBy(item => item.Key,
                StringComparer.Ordinal)) {
                writer.Write(entry.Key);
                writer.Write(entry.Value);
            }
            writer.Write(Diagnostics.Count);
            foreach (EmailStoreDiagnostic diagnostic in Diagnostics) WriteDiagnostic(writer, diagnostic);
            writer.Flush();
            return stream.ToArray();
        }
    }

    internal static PstConversionCheckpointState Deserialize(byte[] payload) {
        if (payload == null) throw new ArgumentNullException(nameof(payload));
        using (var stream = new MemoryStream(payload, writable: false))
        using (var reader = new BinaryReader(stream, Encoding.UTF8, leaveOpen: false)) {
            if (!reader.ReadBytes(Magic.Length).SequenceEqual(Magic)) {
                throw new InvalidDataException("The PST checkpoint does not contain OfficeIMO migration state.");
            }
            int format = reader.ReadInt32();
            if (!Enum.IsDefined(typeof(EmailStoreFormat), format)) {
                throw new InvalidDataException("The migration checkpoint source format is invalid.");
            }
            var state = new PstConversionCheckpointState {
                SourceFormat = (EmailStoreFormat)format,
                SourceLength = reader.ReadInt64(),
                CatalogFingerprint = reader.ReadString(),
                DurableFingerprint = reader.ReadString(),
                OptionsFingerprint = reader.ReadString(),
                WriterDestination = Path.GetFullPath(reader.ReadString()),
                InspectedItems = reader.ReadInt32(),
                ConvertedItems = reader.ReadInt32(),
                SkippedItems = reader.ReadInt32(),
                MappingPath = ReadNullableString(reader),
                MappingLength = reader.ReadInt64(),
                MappingCount = reader.ReadInt32()
            };
            int folderCount = ReadCount(reader, "folder-map");
            for (int index = 0; index < folderCount; index++) {
                state.FolderMap.Add(reader.ReadString(), reader.ReadString());
            }
            int diagnosticCount = ReadCount(reader, "diagnostic");
            for (int index = 0; index < diagnosticCount; index++) state.Diagnostics.Add(ReadDiagnostic(reader));
            if (stream.Position != stream.Length || state.SourceLength < 0 ||
                state.InspectedItems < 0 || state.ConvertedItems < 0 || state.SkippedItems < 0 ||
                state.ConvertedItems + state.SkippedItems > state.InspectedItems ||
                state.MappingLength < 0 || state.MappingCount < 0 ||
                state.CatalogFingerprint.Length != 64 || state.DurableFingerprint.Length != 64 ||
                state.OptionsFingerprint.Length != 64) {
                throw new InvalidDataException("The migration checkpoint state is inconsistent.");
            }
            return state;
        }
    }

    private static int ReadCount(BinaryReader reader, string name) {
        int count = reader.ReadInt32();
        if (count < 0 || count > MaximumEntryCount) {
            throw new InvalidDataException("The migration checkpoint " + name + " count is invalid.");
        }
        return count;
    }

    private static void WriteDiagnostic(BinaryWriter writer, EmailStoreDiagnostic diagnostic) {
        writer.Write(diagnostic.Code);
        writer.Write(diagnostic.Message);
        writer.Write((int)diagnostic.Severity);
        WriteNullableString(writer, diagnostic.Location);
        WriteNullableString(writer, diagnostic.Operation);
        WriteNullableInt64(writer, diagnostic.ByteOffset);
        WriteNullableString(writer, diagnostic.LimitName);
        WriteNullableInt64(writer, diagnostic.ActualValue);
        WriteNullableInt64(writer, diagnostic.MaximumValue);
        writer.Write((int)diagnostic.Disposition);
        writer.Write((int)diagnostic.DataLossRisk);
        WriteNullableString(writer, diagnostic.SuggestedAction);
        writer.Write(diagnostic.IsRetryable);
    }

    private static EmailStoreDiagnostic ReadDiagnostic(BinaryReader reader) => new EmailStoreDiagnostic(
        reader.ReadString(), reader.ReadString(), (EmailStoreDiagnosticSeverity)reader.ReadInt32(),
        ReadNullableString(reader), ReadNullableString(reader), ReadNullableInt64(reader),
        ReadNullableString(reader), ReadNullableInt64(reader), ReadNullableInt64(reader),
        (EmailDiagnosticDisposition)reader.ReadInt32(), (EmailDataLossRisk)reader.ReadInt32(),
        ReadNullableString(reader), reader.ReadBoolean());

    private static void WriteNullableString(BinaryWriter writer, string? value) {
        writer.Write(value != null);
        if (value != null) writer.Write(value);
    }

    private static string? ReadNullableString(BinaryReader reader) =>
        reader.ReadBoolean() ? reader.ReadString() : null;

    private static void WriteNullableInt64(BinaryWriter writer, long? value) {
        writer.Write(value.HasValue);
        if (value.HasValue) writer.Write(value.Value);
    }

    private static long? ReadNullableInt64(BinaryReader reader) =>
        reader.ReadBoolean() ? reader.ReadInt64() : (long?)null;
}
