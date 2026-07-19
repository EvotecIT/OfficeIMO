using OfficeIMO.Email;

namespace OfficeIMO.Email.AddressBook;

/// <summary>One schema-defined property in an OAB v4 header or entry record.</summary>
public sealed class OfflineAddressBookPropertyDefinition {
    internal OfflineAddressBookPropertyDefinition(uint propertyTag, uint flags) {
        PropertyTag = propertyTag;
        PropertyId = unchecked((ushort)(propertyTag >> 16));
        PropertyType = (MapiPropertyType)unchecked((ushort)propertyTag);
        Flags = flags;
    }

    /// <summary>Combined MAPI property tag.</summary>
    public uint PropertyTag { get; }
    /// <summary>High-word MAPI property identifier.</summary>
    public ushort PropertyId { get; }
    /// <summary>Low-word MAPI property type.</summary>
    public MapiPropertyType PropertyType { get; }
    /// <summary>Raw OAB schema flags.</summary>
    public uint Flags { get; }
    /// <summary>The shared OfficeIMO MAPI vocabulary entry for this tag, when known.</summary>
    public MapiPropertyKey? KnownProperty => MapiKnownProperties.Find(PropertyId, PropertyType);
    /// <summary>Whether the property participates in the online ambiguous-name-resolution index.</summary>
    public bool IsAnrIndexed => (Flags & 0x00000001U) != 0;
    /// <summary>Whether the property is an online primary-key value required in every entry.</summary>
    public bool IsPrimaryKey => (Flags & 0x00000002U) != 0;
}
