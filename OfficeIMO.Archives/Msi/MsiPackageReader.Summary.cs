using System;
using System.IO;
using System.Linq;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Archives;

internal sealed partial class MsiPackageReader {
    /// <summary>Reads PID_WORDCOUNT's source-compression bit; an absent property defaults to zero.</summary>
    private bool ReadDefaultCompression() {
        const string name = "\u0005SummaryInformation";
        if (!_streams.ContainsKey(name)) return false;
        using Stream source = _compound.OpenStream(name);
        ArchiveLimits.Add(ref _metadataBytes, source.Length, _limits.Metadata, "MaxMetadataBytes");
        ArchiveLimits.Check(source.Length, Math.Min(_limits.Metadata, int.MaxValue), "MaxMetadataBytes");
        byte[] bytes = OfficeArchiveSafety.ReadEntryBytes(source, source.Length, _limits.Metadata, _token);
        OfficeOlePropertySection[] summaries = OfficeOlePropertySetReader.ReadSections(bytes, _token)
            .Where(section => section.FormatId == OfficeOlePropertySetWriter.SummaryInformationFormatId).ToArray();
        if (summaries.Length != 1) throw new InvalidDataException("MSI SummaryInformation requires one summary property section.");
        if (!summaries[0].Properties.TryGetValue(15, out OfficeOlePropertyValue? value)) return false;
        if (value.Type != 3 || value.Value is not int flags)
            throw new InvalidDataException("MSI PID_WORDCOUNT requires a 32-bit integer.");
        return (flags & 2) != 0;
    }
}
