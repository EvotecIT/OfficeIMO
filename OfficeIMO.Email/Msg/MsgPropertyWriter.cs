using OfficeIMO.Shared;

namespace OfficeIMO.Email;

internal static class MsgPropertyWriter {
    internal static void Write(string prefix, MsgPropertyStreamKind kind, IReadOnlyList<MapiProperty> properties,
        int recipientCount, int attachmentCount, MsgNamedPropertyWriter names, IList<OfficeCompoundStream> streams,
        IList<EmailDiagnostic> diagnostics, uint objectReserved = 0) {
        var resolved = properties.Select(property => new ResolvedProperty(property,
                property.Name == null ? property.PropertyId : names.GetPropertyId(property.Name)))
            .GroupBy(item => item.Tag)
            .Select(group => group.Last())
            .OrderBy(item => item.Tag)
            .ToArray();
        int headerLength = kind == MsgPropertyStreamKind.TopLevel ? 32 :
            kind == MsgPropertyStreamKind.EmbeddedMessage ? 24 : 8;
        byte[] propertyStream = new byte[checked(headerLength + resolved.Length * 16)];
        if (kind != MsgPropertyStreamKind.ChildObject) {
            MsgBinary.WriteUInt32(propertyStream, 8, unchecked((uint)recipientCount));
            MsgBinary.WriteUInt32(propertyStream, 12, unchecked((uint)attachmentCount));
            MsgBinary.WriteUInt32(propertyStream, 16, unchecked((uint)recipientCount));
            MsgBinary.WriteUInt32(propertyStream, 20, unchecked((uint)attachmentCount));
        }

        for (int index = 0; index < resolved.Length; index++) {
            ResolvedProperty resolvedProperty = resolved[index];
            MapiProperty property = resolvedProperty.Property;
            int offset = headerLength + index * 16;
            MsgBinary.WriteUInt32(propertyStream, offset, resolvedProperty.Tag);
            MsgBinary.WriteUInt32(propertyStream, offset + 4, property.Flags);
            try {
                if (MsgValueWriter.IsMultiple(property.PropertyType)) {
                    byte[] valueStream = WriteMultiple(prefix, resolvedProperty.Tag, property, streams);
                    MsgBinary.WriteUInt32(propertyStream, offset + 8, unchecked((uint)valueStream.Length));
                } else if (MsgValueWriter.IsVariable(property.PropertyType)) {
                    if (property.PropertyType == MapiPropertyType.Object) {
                        MsgBinary.WriteUInt32(propertyStream, offset + 8, 0xffffffffU);
                        MsgBinary.WriteUInt32(propertyStream, offset + 12, objectReserved);
                    } else {
                        byte[] value = MsgValueWriter.EncodeScalar(property);
                        streams.Add(new OfficeCompoundStream(MsgBinary.CombinePath(prefix,
                            string.Concat("__substg1.0_", resolvedProperty.Tag.ToString("X8", CultureInfo.InvariantCulture))), value));
                        uint size = unchecked((uint)value.Length);
                        if (property.PropertyType == MapiPropertyType.Unicode) size = checked(size + 2);
                        if (property.PropertyType == MapiPropertyType.String8) size = checked(size + 1);
                        MsgBinary.WriteUInt32(propertyStream, offset + 8, size);
                    }
                } else {
                    byte[] value = MsgValueWriter.EncodeFixedValue(property);
                    Buffer.BlockCopy(value, 0, propertyStream, offset + 8, 8);
                }
            } catch (Exception ex) when (ex is ArgumentException || ex is InvalidCastException || ex is FormatException || ex is OverflowException) {
                diagnostics.Add(new EmailDiagnostic("EMAIL_MSG_PROPERTY_WRITE_INVALID",
                    string.Concat("Property 0x", resolvedProperty.Tag.ToString("X8", CultureInfo.InvariantCulture),
                        " could not be serialized: ", ex.Message), EmailDiagnosticSeverity.Error, prefix));
            }
        }
        streams.Add(new OfficeCompoundStream(MsgBinary.CombinePath(prefix, "__properties_version1.0"), propertyStream));
    }

    private static byte[] WriteMultiple(string prefix, uint tag, MapiProperty property,
        IList<OfficeCompoundStream> streams) {
        object[] values = MsgValueWriter.GetMultipleValues(property);
        MapiPropertyType itemType = MsgValueWriter.GetMultipleItemType(property.PropertyType);
        string baseName = string.Concat("__substg1.0_", tag.ToString("X8", CultureInfo.InvariantCulture));
        if (property.PropertyType == MapiPropertyType.MultipleBinary ||
            property.PropertyType == MapiPropertyType.MultipleString8 ||
            property.PropertyType == MapiPropertyType.MultipleUnicode) {
            int lengthEntrySize = property.PropertyType == MapiPropertyType.MultipleBinary ? 8 : 4;
            byte[] lengths = new byte[checked(values.Length * lengthEntrySize)];
            for (int index = 0; index < values.Length; index++) {
                var item = new MapiProperty(property.PropertyId, itemType, values[index]);
                byte[] value = MsgValueWriter.EncodeScalar(item);
                if (itemType == MapiPropertyType.Unicode) value = AppendTerminator(value, 2);
                if (itemType == MapiPropertyType.String8) value = AppendTerminator(value, 1);
                MsgBinary.WriteUInt32(lengths, index * lengthEntrySize, unchecked((uint)value.Length));
                streams.Add(new OfficeCompoundStream(MsgBinary.CombinePath(prefix,
                    string.Concat(baseName, "-", index.ToString("X8", CultureInfo.InvariantCulture))), value));
            }
            streams.Add(new OfficeCompoundStream(MsgBinary.CombinePath(prefix, baseName), lengths));
            return lengths;
        }

        using (MemoryStream output = new MemoryStream()) {
            foreach (object value in values) {
                var item = new MapiProperty(property.PropertyId, itemType, value);
                byte[] bytes = MsgValueWriter.EncodeScalar(item);
                output.Write(bytes, 0, bytes.Length);
            }
            byte[] result = output.ToArray();
            streams.Add(new OfficeCompoundStream(MsgBinary.CombinePath(prefix, baseName), result));
            return result;
        }
    }

    private static byte[] AppendTerminator(byte[] value, int count) {
        byte[] result = new byte[value.Length + count];
        Buffer.BlockCopy(value, 0, result, 0, value.Length);
        return result;
    }

    private sealed class ResolvedProperty {
        internal ResolvedProperty(MapiProperty property, ushort propertyId) {
            Property = property;
            Tag = ((uint)propertyId << 16) | (ushort)property.PropertyType;
        }

        internal MapiProperty Property { get; }

        internal uint Tag { get; }
    }
}
