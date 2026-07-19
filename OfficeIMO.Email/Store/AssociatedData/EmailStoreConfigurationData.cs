using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

/// <summary>Lossless roaming XML and dictionary streams projected from one FAI message.</summary>
public sealed class EmailStoreConfigurationData {
    internal EmailStoreConfigurationData(EmailDocument document, int maxXmlBytes,
        IList<EmailStoreDiagnostic> diagnostics, string location) {
        MessageClass = document.MessageClass;
        RoamingDatatypes = document.Mapi.GetNullableValue(MapiKnownProperties.PidTag.RoamingDatatypes);
        XmlBytes = Copy(document.Mapi.GetValueOrDefault(MapiKnownProperties.PidTag.RoamingXmlStream));
        DictionaryBytes = Copy(document.Mapi.GetValueOrDefault(MapiKnownProperties.PidTag.RoamingDictionary));
        if (XmlBytes != null) {
            try {
                // Parse once to enforce secure bounded XML and expose normalized text without interpreting its schema.
                XmlText = Encoding.UTF8.GetString(EmailStoreConfigurationXml.Serialize(
                    EmailStoreConfigurationXml.Parse(XmlBytes, maxXmlBytes, "The roaming XML stream")));
            } catch (Exception exception) when (exception is InvalidDataException ||
                exception is EmailStoreLimitExceededException) {
                diagnostics.Add(new EmailStoreDiagnostic("EMAIL_STORE_FAI_XML_INVALID", exception.Message,
                    EmailStoreDiagnosticSeverity.Error, location));
            }
        }
        if (DictionaryBytes != null) {
            try {
                Dictionary = EmailStoreConfigurationDictionary.Parse(DictionaryBytes, maxXmlBytes);
            } catch (Exception exception) when (exception is InvalidDataException ||
                exception is EmailStoreLimitExceededException) {
                diagnostics.Add(new EmailStoreDiagnostic("EMAIL_STORE_FAI_DICTIONARY_INVALID", exception.Message,
                    EmailStoreDiagnosticSeverity.Error, location));
            }
        }
        bool xmlFlag = (RoamingDatatypes.GetValueOrDefault() & 0x00000002) != 0;
        bool dictionaryFlag = (RoamingDatatypes.GetValueOrDefault() & 0x00000004) != 0;
        if (XmlBytes != null != xmlFlag || DictionaryBytes != null != dictionaryFlag) {
            diagnostics.Add(new EmailStoreDiagnostic("EMAIL_STORE_FAI_ROAMING_FLAGS_MISMATCH",
                "PidTagRoamingDatatypes does not match the roaming streams present on the associated message.",
                EmailStoreDiagnosticSeverity.Warning, location));
        }
    }

    /// <summary>Configuration message class.</summary>
    public string? MessageClass { get; }
    /// <summary>Raw PidTagRoamingDatatypes bitmask.</summary>
    public int? RoamingDatatypes { get; }
    /// <summary>Exact PidTagRoamingXmlStream bytes.</summary>
    public byte[]? XmlBytes { get; }
    /// <summary>Securely parsed and reserialized UTF-8 XML when valid and within bounds.</summary>
    public string? XmlText { get; }
    /// <summary>Exact PidTagRoamingDictionary bytes.</summary>
    public byte[]? DictionaryBytes { get; }
    /// <summary>Typed lossless dictionary projection when valid.</summary>
    public EmailStoreConfigurationDictionary? Dictionary { get; }

    /// <summary>True when the message advertises or carries either supported roaming stream.</summary>
    public bool HasData => RoamingDatatypes.HasValue || XmlBytes != null || DictionaryBytes != null;

    private static byte[]? Copy(byte[]? value) => value == null ? null : (byte[])value.Clone();
}
