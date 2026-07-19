using OfficeIMO.Email;

namespace OfficeIMO.Email.AddressBook;

internal static class OabV4MetadataReader {
    internal static OfflineAddressBookListInfo Read(OabSource source, int index,
        OfflineAddressBookReaderOptions options, ICollection<EmailDiagnostic> diagnostics) {
        using (OabStreamLease lease = source.OpenRead()) {
            Stream stream = lease.Stream;
            OabBinary.Seek(source, stream, 0, source.SourcePath);
            if (source.Length < 16) throw new InvalidDataException("OAB v4 Full Details header is truncated.");
            uint version = OabBinary.ReadUInt32(stream, source.SourcePath);
            if (version != 0x00000020U) {
                throw new NotSupportedException(string.Concat(
                    "OAB component is not an uncompressed version 4 Full Details file: 0x",
                    version.ToString("X8", CultureInfo.InvariantCulture), "."));
            }
            uint serial = OabBinary.ReadUInt32(stream, source.SourcePath);
            uint declared = OabBinary.ReadUInt32(stream, source.SourcePath);
            if (declared > options.MaxDeclaredEntries) {
                throw new OfflineAddressBookLimitExceededException(
                    nameof(options.MaxDeclaredEntries), declared, options.MaxDeclaredEntries, source.SourcePath);
            }
            uint encodedMetadataSize = OabBinary.ReadUInt32(stream, source.SourcePath);
            if (encodedMetadataSize > int.MaxValue || encodedMetadataSize > options.MaxMetadataBytes) {
                throw new OfflineAddressBookLimitExceededException(
                    nameof(options.MaxMetadataBytes), encodedMetadataSize, options.MaxMetadataBytes, source.SourcePath);
            }
            int metadataSize = checked((int)encodedMetadataSize);
            if (metadataSize < 12 || 12L + metadataSize > source.Length) {
                throw new InvalidDataException("OAB v4 metadata size is invalid.");
            }
            byte[] metadata = OabBinary.ReadExactly(stream, metadataSize - 4, source.SourcePath);
            int cursor = 0;
            IReadOnlyList<OfflineAddressBookPropertyDefinition> headerDefinitions =
                ReadPropertyTable(metadata, ref cursor, options, source.SourcePath + "/header-schema");
            IReadOnlyList<OfflineAddressBookPropertyDefinition> entryDefinitions =
                ReadPropertyTable(metadata, ref cursor, options, source.SourcePath + "/entry-schema");
            if (cursor != metadata.Length) {
                diagnostics.Add(new EmailDiagnostic(
                    "OAB_METADATA_TRAILING_DATA",
                    string.Concat((metadata.Length - cursor).ToString(CultureInfo.InvariantCulture),
                        " unconsumed byte(s) remain in the OAB metadata structure."),
                    EmailDiagnosticSeverity.Warning,
                    source.SourcePath));
            }
            if (headerDefinitions.Count < 4) {
                throw new InvalidDataException("OAB v4 header schema has fewer than four required properties.");
            }
            if (entryDefinitions.Count < 2 ||
                entryDefinitions[0].PropertyId != MapiKnownProperties.PidTag.EmailAddress.PropertyId ||
                entryDefinitions[1].PropertyId != MapiKnownProperties.PidTag.SmtpAddress.PropertyId) {
                diagnostics.Add(new EmailDiagnostic(
                    "OAB_SCHEMA_REQUIRED_ORDER",
                    "The OAB entry schema does not begin with the required email and SMTP address properties.",
                    EmailDiagnosticSeverity.Warning,
                    source.SourcePath));
            }

            long headerRecordOffset = checked(12L + metadataSize);
            OabBinary.Seek(source, stream, headerRecordOffset, source.SourcePath + "/header-record");
            OabRecordEnvelope headerEnvelope = OabV4RecordReader.ReadEnvelope(
                source, stream, options, source.SourcePath + "/header-record");
            OabParsedRecord header = OabV4RecordReader.Parse(
                headerEnvelope, headerDefinitions, options, source.SourcePath + "/header-record");
            foreach (EmailDiagnostic diagnostic in header.Diagnostics) diagnostics.Add(diagnostic);
            long entriesOffset = checked(headerRecordOffset + headerEnvelope.Size);
            if (entriesOffset > source.Length) throw new InvalidDataException("OAB header record extends beyond the source.");
            return new OfflineAddressBookListInfo(
                string.Concat("list-", index.ToString("D4", CultureInfo.InvariantCulture)),
                index,
                source.SourcePath,
                source.Length,
                serial,
                declared,
                headerDefinitions,
                entryDefinitions,
                header.Properties,
                entriesOffset);
        }
    }

    private static IReadOnlyList<OfflineAddressBookPropertyDefinition> ReadPropertyTable(
        byte[] metadata, ref int cursor, OfflineAddressBookReaderOptions options, string location) {
        uint encodedCount = OabBinary.UInt32(metadata, cursor);
        cursor += 4;
        if (encodedCount > int.MaxValue || encodedCount > options.MaxPropertiesPerTable) {
            throw new OfflineAddressBookLimitExceededException(
                nameof(options.MaxPropertiesPerTable), encodedCount, options.MaxPropertiesPerTable, location);
        }
        int count = checked((int)encodedCount);
        if (cursor > metadata.Length - checked(count * 8)) {
            throw new InvalidDataException(string.Concat("OAB property table is truncated at ", location, "."));
        }
        var result = new OfflineAddressBookPropertyDefinition[count];
        for (int index = 0; index < count; index++) {
            uint propertyTag = OabBinary.UInt32(metadata, cursor);
            uint flags = OabBinary.UInt32(metadata, cursor + 4);
            cursor += 8;
            ValidatePropertyType(propertyTag, location);
            result[index] = new OfflineAddressBookPropertyDefinition(propertyTag, flags);
        }
        return result;
    }

    private static void ValidatePropertyType(uint propertyTag, string location) {
        switch ((MapiPropertyType)unchecked((ushort)propertyTag)) {
            case MapiPropertyType.Integer32:
            case MapiPropertyType.Boolean:
            case MapiPropertyType.Object:
            case MapiPropertyType.String8:
            case MapiPropertyType.Unicode:
            case MapiPropertyType.Binary:
            case MapiPropertyType.MultipleInteger32:
            case MapiPropertyType.MultipleString8:
            case MapiPropertyType.MultipleUnicode:
            case MapiPropertyType.MultipleBinary:
                return;
            default:
                throw new NotSupportedException(string.Concat(
                    "Unsupported OAB schema property type 0x",
                    unchecked((ushort)propertyTag).ToString("X4", CultureInfo.InvariantCulture),
                    " at ", location, "."));
        }
    }
}
