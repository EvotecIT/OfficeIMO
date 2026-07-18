namespace OfficeIMO.Email;

/// <summary>Well-known MAPI named-property sets used by Outlook artifacts.</summary>
public static class MapiPropertySets {
    /// <summary>Appointment properties (PSETID_Appointment).</summary>
    public static readonly Guid Appointment = new Guid("00062002-0000-0000-C000-000000000046");
    /// <summary>Meeting communication properties (PSETID_Meeting).</summary>
    public static readonly Guid Meeting = new Guid("6ED8DA90-450B-101B-98DA-00AA003F1305");
    /// <summary>Task properties (PSETID_Task).</summary>
    public static readonly Guid Task = new Guid("00062003-0000-0000-C000-000000000046");
    /// <summary>Contact and address properties (PSETID_Address).</summary>
    public static readonly Guid Address = new Guid("00062004-0000-0000-C000-000000000046");
    /// <summary>Common Outlook item properties (PSETID_Common).</summary>
    public static readonly Guid Common = new Guid("00062008-0000-0000-C000-000000000046");
    /// <summary>Journal properties (PSETID_Log).</summary>
    public static readonly Guid Log = new Guid("0006200A-0000-0000-C000-000000000046");
    /// <summary>Sticky-note properties (PSETID_Note).</summary>
    public static readonly Guid Note = new Guid("0006200E-0000-0000-C000-000000000046");
    /// <summary>Standard MAPI named properties (PS_MAPI).</summary>
    public static readonly Guid Mapi = new Guid("00020328-0000-0000-C000-000000000046");
    /// <summary>Public string-named properties (PS_PUBLIC_STRINGS).</summary>
    public static readonly Guid PublicStrings = new Guid("00020329-0000-0000-C000-000000000046");
    /// <summary>Internet header string-named properties (PS_INTERNET_HEADERS).</summary>
    public static readonly Guid InternetHeaders = new Guid("00020386-0000-0000-C000-000000000046");
    /// <summary>Calendar assistant properties.</summary>
    public static readonly Guid CalendarAssistant = new Guid("11000E07-B51B-40D6-AF21-CAA85EDAB1D0");
    /// <summary>Sharing properties (PSETID_Sharing).</summary>
    public static readonly Guid Sharing = new Guid("00062040-0000-0000-C000-000000000046");
    /// <summary>Outlook reaction properties.</summary>
    public static readonly Guid Reactions = new Guid("41F28F13-83F4-4114-A584-EEDB5A6B0BFF");
    /// <summary>OfficeIMO.Email.Store artifact provenance properties.</summary>
    public static readonly Guid OfficeImoEmailStore = new Guid("0A5A57A3-CF06-4989-AE40-5A65BFC0C126");
}
