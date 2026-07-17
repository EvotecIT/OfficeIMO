namespace OfficeIMO.Email;

/// <summary>A standalone iCalendar stream containing one or more VCALENDAR objects.</summary>
public sealed partial class IcsDocument {
    private const string DefaultProductId = "-//OfficeIMO//OfficeIMO.Email//EN";
    private readonly List<ContentLineComponent> _calendars = new List<ContentLineComponent>();

    /// <summary>Creates a document with one RFC 5545 VCALENDAR object.</summary>
    public IcsDocument(string productId = DefaultProductId) {
        ContentLineComponent calendar = AddCalendar();
        calendar.AddProperty("PRODID", productId ?? throw new ArgumentNullException(nameof(productId)));
        calendar.AddProperty("VERSION", "2.0");
    }

    private IcsDocument(bool createDefault) {
        if (createDefault) AddCalendar();
    }

    /// <summary>Ordered VCALENDAR roots. Unknown properties and nested components are retained.</summary>
    public IList<ContentLineComponent> Calendars => _calendars;

    /// <summary>Adds an empty VCALENDAR root.</summary>
    public ContentLineComponent AddCalendar() {
        var calendar = new ContentLineComponent("VCALENDAR");
        _calendars.Add(calendar);
        return calendar;
    }

    /// <summary>Enumerates matching components across every calendar.</summary>
    public IEnumerable<ContentLineComponent> GetComponents(string name, bool recursive = true) {
        foreach (ContentLineComponent calendar in _calendars) {
            if (string.Equals(calendar.Name, name, StringComparison.OrdinalIgnoreCase)) yield return calendar;
            foreach (ContentLineComponent component in calendar.GetComponents(name, recursive)) yield return component;
        }
    }

    /// <summary>Parses an iCalendar stream from text.</summary>
    public static IcsDocument Parse(string text, ContentLineReaderOptions? options = null) =>
        FromComponents(ContentLineCodec.Parse(text, options ?? ContentLineReaderOptions.Default));

    /// <summary>Loads an iCalendar stream from memory.</summary>
    public static IcsDocument Load(byte[] bytes, ContentLineReaderOptions? options = null) =>
        FromComponents(ContentLineCodec.Parse(bytes, options ?? ContentLineReaderOptions.Default));

    /// <summary>Loads an iCalendar file.</summary>
    public static IcsDocument Load(string filePath, ContentLineReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        ContentLineReaderOptions effective = options ?? ContentLineReaderOptions.Default;
        return Load(ContentLineDocumentIO.Read(filePath, effective, cancellationToken), effective);
    }

    /// <summary>Loads from a caller-owned stream without closing it.</summary>
    public static IcsDocument Load(Stream stream, ContentLineReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        ContentLineReaderOptions effective = options ?? ContentLineReaderOptions.Default;
        return Load(ContentLineDocumentIO.Read(stream, effective, cancellationToken), effective);
    }

    /// <summary>Asynchronously loads an iCalendar file.</summary>
    public static async Task<IcsDocument> LoadAsync(string filePath, ContentLineReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        ContentLineReaderOptions effective = options ?? ContentLineReaderOptions.Default;
        return Load(await ContentLineDocumentIO.ReadAsync(filePath, effective, cancellationToken)
            .ConfigureAwait(false), effective);
    }

    /// <summary>Asynchronously loads from a caller-owned stream without closing it.</summary>
    public static async Task<IcsDocument> LoadAsync(Stream stream, ContentLineReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        ContentLineReaderOptions effective = options ?? ContentLineReaderOptions.Default;
        return Load(await ContentLineDocumentIO.ReadAsync(stream, effective, cancellationToken)
            .ConfigureAwait(false), effective);
    }

    /// <summary>Serializes this document to bytes.</summary>
    public byte[] ToBytes(ContentLineWriterOptions? options = null) {
        ValidateCalendars(_calendars);
        return ContentLineCodec.Serialize(_calendars, options ?? ContentLineWriterOptions.Default);
    }

    /// <summary>Serializes this document to text using the configured output encoding.</summary>
    public string Serialize(ContentLineWriterOptions? options = null) {
        ContentLineWriterOptions effective = options ?? ContentLineWriterOptions.Default;
        return effective.Encoding.GetString(ToBytes(effective));
    }

    /// <summary>Atomically saves this document to a file.</summary>
    public void Save(string filePath, ContentLineWriterOptions? options = null) =>
        ContentLineDocumentIO.Write(filePath, ToBytes(options));

    /// <summary>Saves this document to a caller-owned stream without closing it.</summary>
    public void Save(Stream stream, ContentLineWriterOptions? options = null) =>
        ContentLineDocumentIO.Write(stream, ToBytes(options));

    /// <summary>Asynchronously atomically saves this document to a file.</summary>
    public Task SaveAsync(string filePath, ContentLineWriterOptions? options = null,
        CancellationToken cancellationToken = default) =>
        ContentLineDocumentIO.WriteAsync(filePath, ToBytes(options), cancellationToken);

    /// <summary>Asynchronously saves to a caller-owned stream without closing it.</summary>
    public Task SaveAsync(Stream stream, ContentLineWriterOptions? options = null,
        CancellationToken cancellationToken = default) =>
        ContentLineDocumentIO.WriteAsync(stream, ToBytes(options), cancellationToken);

    private static IcsDocument FromComponents(IReadOnlyList<ContentLineComponent> components) {
        ValidateCalendars(components);
        var document = new IcsDocument(false);
        foreach (ContentLineComponent component in components) document._calendars.Add(component);
        return document;
    }

    private static void ValidateCalendars(IEnumerable<ContentLineComponent> calendars) {
        int count = 0;
        foreach (ContentLineComponent calendar in calendars) {
            count++;
            if (!string.Equals(calendar.Name, "VCALENDAR", StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException(
                    "The iCalendar stream contains a non-VCALENDAR root component.");
        }
        if (count == 0) throw new InvalidDataException(
            "The iCalendar document does not contain a calendar.");
    }
}
