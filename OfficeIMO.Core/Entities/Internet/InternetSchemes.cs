namespace OfficeIMO.Entities;

/// <summary>Internet email address and URL schemes.</summary>
public static class InternetSchemes {
    /// <summary>SMTP mailbox address.</summary>
    public static EntityScheme Email => EmailInstance;
    /// <summary>HTTP, HTTPS, FTP, mailto or www URL.</summary>
    public static EntityScheme Url { get; } = new UrlScheme();

    internal static readonly EmailScheme EmailInstance = new();
}
