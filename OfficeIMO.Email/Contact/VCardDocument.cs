namespace OfficeIMO.Email;

/// <summary>Supported vCard syntax versions.</summary>
public enum VCardVersion {
    /// <summary>Legacy vCard 2.1 compatibility syntax.</summary>
    V2_1 = 21,
    /// <summary>vCard 3.0 syntax.</summary>
    V3_0 = 30,
    /// <summary>RFC 6350 vCard 4.0 syntax.</summary>
    V4_0 = 40
}

/// <summary>A standalone vCard stream containing one or more ordered VCARD objects.</summary>
public sealed partial class VCardDocument {
    private readonly List<ContentLineComponent> _cards = new List<ContentLineComponent>();

    /// <summary>Creates a document with one vCard 4.0 card.</summary>
    public VCardDocument() {
        AddCard(VCardVersion.V4_0);
    }

    private VCardDocument(bool createDefault) {
        if (createDefault) AddCard(VCardVersion.V4_0);
    }

    /// <summary>Ordered VCARD roots. Repeated, grouped, unknown, and extension properties are retained.</summary>
    public IList<ContentLineComponent> Cards => _cards;

    /// <summary>Adds an empty card with an explicit VERSION property.</summary>
    public ContentLineComponent AddCard(VCardVersion version = VCardVersion.V4_0) {
        var card = new ContentLineComponent("VCARD");
        card.AddProperty("VERSION", FormatVersion(version));
        _cards.Add(card);
        return card;
    }

    /// <summary>Adds an RFC 6350 group card with ordered MEMBER URI properties.</summary>
    public ContentLineComponent AddGroup(string formattedName, IEnumerable<string> memberUris) {
        if (formattedName == null) throw new ArgumentNullException(nameof(formattedName));
        if (memberUris == null) throw new ArgumentNullException(nameof(memberUris));
        ContentLineComponent card;
        if (_cards.Count == 1 && _cards[0].Properties.Count == 1 &&
            string.Equals(_cards[0].Properties[0].Name, "VERSION", StringComparison.OrdinalIgnoreCase)) {
            card = _cards[0];
            SetVersion(card, VCardVersion.V4_0);
        } else card = AddCard(VCardVersion.V4_0);
        card.SetVCardText("FN", formattedName);
        card.AddProperty("KIND", "group");
        foreach (string memberUri in memberUris) {
            if (string.IsNullOrWhiteSpace(memberUri))
                throw new ArgumentException("A group MEMBER URI cannot be empty.", nameof(memberUris));
            card.AddProperty("MEMBER", memberUri);
        }
        return card;
    }

    /// <summary>Gets a card's declared syntax version.</summary>
    public static VCardVersion GetVersion(ContentLineComponent card) {
        if (card == null) throw new ArgumentNullException(nameof(card));
        ContentLineProperty? property = card.GetFirstProperty("VERSION");
        if (property == null) throw new InvalidDataException("A VCARD component does not declare VERSION.");
        return ParseVersion(property.Value);
    }

    /// <summary>Replaces a card's declared syntax version.</summary>
    public static void SetVersion(ContentLineComponent card, VCardVersion version) {
        if (card == null) throw new ArgumentNullException(nameof(card));
        for (int index = card.Properties.Count - 1; index >= 0; index--) {
            if (string.Equals(card.Properties[index].Name, "VERSION", StringComparison.OrdinalIgnoreCase))
                card.Properties.RemoveAt(index);
        }
        card.Properties.Insert(0, new ContentLineProperty("VERSION", FormatVersion(version)));
    }

    /// <summary>Parses a vCard stream from text.</summary>
    public static VCardDocument Parse(string text, ContentLineReaderOptions? options = null) =>
        FromComponents(ContentLineCodec.Parse(text, options ?? ContentLineReaderOptions.Default,
            decodeRfc6868Parameters: false));

    /// <summary>Loads a vCard stream from memory.</summary>
    public static VCardDocument Load(byte[] bytes, ContentLineReaderOptions? options = null) =>
        FromComponents(ContentLineCodec.Parse(bytes, options ?? ContentLineReaderOptions.Default,
            decodeRfc6868Parameters: false));

    /// <summary>Loads a vCard file.</summary>
    public static VCardDocument Load(string filePath, ContentLineReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        ContentLineReaderOptions effective = options ?? ContentLineReaderOptions.Default;
        return Load(ContentLineDocumentIO.Read(filePath, effective, cancellationToken), effective);
    }

    /// <summary>Loads from a caller-owned stream without closing it.</summary>
    public static VCardDocument Load(Stream stream, ContentLineReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        ContentLineReaderOptions effective = options ?? ContentLineReaderOptions.Default;
        return Load(ContentLineDocumentIO.Read(stream, effective, cancellationToken), effective);
    }

    /// <summary>Asynchronously loads a vCard file.</summary>
    public static async Task<VCardDocument> LoadAsync(string filePath, ContentLineReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        ContentLineReaderOptions effective = options ?? ContentLineReaderOptions.Default;
        return Load(await ContentLineDocumentIO.ReadAsync(filePath, effective, cancellationToken)
            .ConfigureAwait(false), effective);
    }

    /// <summary>Asynchronously loads from a caller-owned stream without closing it.</summary>
    public static async Task<VCardDocument> LoadAsync(Stream stream, ContentLineReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        ContentLineReaderOptions effective = options ?? ContentLineReaderOptions.Default;
        return Load(await ContentLineDocumentIO.ReadAsync(stream, effective, cancellationToken)
            .ConfigureAwait(false), effective);
    }

    /// <summary>Serializes this document to bytes.</summary>
    public byte[] ToBytes(ContentLineWriterOptions? options = null) {
        ValidateCards(_cards);
        return ContentLineCodec.Serialize(_cards, options ?? ContentLineWriterOptions.Default,
            card => GetVersion(card) == VCardVersion.V4_0
                ? ContentLineParameterEncoding.Rfc6868
                : ContentLineParameterEncoding.Legacy);
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

    private static VCardDocument FromComponents(IReadOnlyList<ContentLineComponent> components) {
        if (components.Count == 0) throw new InvalidDataException("The vCard stream does not contain VCARD.");
        ValidateCards(components);
        var document = new VCardDocument(false);
        foreach (ContentLineComponent component in components) {
            if (GetVersion(component) == VCardVersion.V4_0)
                ContentLineCodec.DecodeRfc6868Parameters(component);
            document._cards.Add(component);
        }
        return document;
    }

    private static void ValidateCards(IEnumerable<ContentLineComponent> cards) {
        int count = 0;
        foreach (ContentLineComponent card in cards) {
            count++;
            if (!string.Equals(card.Name, "VCARD", StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException("The vCard stream contains a non-VCARD root component.");
            GetVersion(card);
        }
        if (count == 0) throw new InvalidDataException("The vCard document does not contain a card.");
    }

    private static VCardVersion ParseVersion(string value) {
        if (string.Equals(value, "2.1", StringComparison.Ordinal)) return VCardVersion.V2_1;
        if (string.Equals(value, "3.0", StringComparison.Ordinal)) return VCardVersion.V3_0;
        if (string.Equals(value, "4.0", StringComparison.Ordinal)) return VCardVersion.V4_0;
        throw new InvalidDataException("Unsupported vCard VERSION '" + value + "'.");
    }

    private static string FormatVersion(VCardVersion version) {
        switch (version) {
            case VCardVersion.V2_1: return "2.1";
            case VCardVersion.V3_0: return "3.0";
            case VCardVersion.V4_0: return "4.0";
            default: throw new ArgumentOutOfRangeException(nameof(version));
        }
    }
}
