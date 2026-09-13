using OfficeIMO.Core.Internal;
using System.Globalization;

namespace OfficeIMO.Project;

internal static partial class ProjectNativeCodec {
    private static void ReadMetadata(ProjectDocument document, OfficeCompoundFile file, CancellationToken token) {
        foreach (var name in new[] { OfficeOlePropertySetWriter.SummaryInformationStreamName, OfficeOlePropertySetWriter.DocumentSummaryInformationStreamName }) {
            if (!file.Streams.TryGetValue(name, out var bytes)) continue;
            foreach (var section in OfficeOlePropertySetReader.ReadSections(bytes, token)) {
                string? Text(uint id) => section.Properties.TryGetValue(id, out var value) ? value.AsString() : null;
                if (section.FormatId == OfficeOlePropertySetWriter.SummaryInformationFormatId) {
                    document.Title = Text(2); document.Subject = Text(3); document.Author = Text(4);
                    string? producer = Text(18);
                    document.NativeSource!.CreatedByOfficeIMO = string.Equals(producer, ProjectNativeSource.ProducerMarker, StringComparison.Ordinal);
                    string prefix = ProjectNativeSource.ProducerMarker + ":";
                    if (producer != null && producer.StartsWith(prefix, StringComparison.Ordinal)
                        && Guid.TryParseExact(producer.Substring(prefix.Length), "D", out var seed)) {
                        document.NativeSource.CreatedByOfficeIMO = true;
                        document.NativeSource.GeneratedIdentitySeed = seed;
                        document.NativeIdentity = seed;
                        if (document.Guid == seed) document.Guid = null;
                    }
                } else if (section.FormatId == OfficeOlePropertySetWriter.DocumentSummaryInformationFormatId) {
                    document.Manager = Text(14); document.Company = Text(15);
                }
            }
        }
    }
    private static void ReadCustomAliases(ProjectDocument document, OfficeCompoundFile file, ProjectNativeProfile profile, CancellationToken token) {
        foreach (var name in new[] { "Task", "Rsc", "Assn" }) {
            if (!file.Streams.TryGetValue(profile.DataRoot + "/TBknd" + name + "/Props", out var bytes)) continue;
            var properties = ProjectNativeProperties.Read(bytes, token);
            if (!properties.TryGetValue(0x04400001, out var values)) continue;
            int length = checked(values.Int32() + 4), count = values.Int32(8);
            if (length < 12 || length > values.Length || values.Int32() != values.Int32(4) || count < 0 || count > (length - 12) / 112 || length != 12 + count * 112)
                throw new NotSupportedException("Unqualified native custom-field alias table.");
            for (int index = 0; index < count; index++) {
                token.ThrowIfCancellationRequested();
                int offset = 12 + index * 112;
                string id = values.UInt32(offset).ToString(CultureInfo.InvariantCulture);
                if (document.CustomFields.Any(f => f.FieldId == id)) throw new InvalidDataException("Duplicate custom-field alias.");
                var definition = document.CustomFields.Add(); definition.FieldId = id;
                definition.Alias = values.Slice(offset + 8, 104).Unicode().Split('\0')[0];
            }
        }
    }
}
