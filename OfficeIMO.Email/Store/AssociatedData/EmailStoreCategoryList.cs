using OfficeIMO.Email;
using System.Xml.Linq;

namespace OfficeIMO.Email.Store;

/// <summary>One Outlook master-category definition from a CategoryList configuration stream.</summary>
public sealed class EmailStoreCategoryDefinition {
    internal EmailStoreCategoryDefinition(XElement element) {
        Name = (string?)element.Attribute("name") ?? string.Empty;
        Color = ParseInt((string?)element.Attribute("color"));
        KeyboardShortcut = ParseUInt((string?)element.Attribute("keyboardShortcut"));
        UsageCount = ParseUInt((string?)element.Attribute("usageCount"));
        LastUsed = EmailStoreConfigurationXml.ParseDate((string?)element.Attribute("lastTimeUsed"));
        LastUsedMail = EmailStoreConfigurationXml.ParseDate((string?)element.Attribute("lastTimeUsedMail"));
        LastUsedCalendar = EmailStoreConfigurationXml.ParseDate((string?)element.Attribute("lastTimeUsedCalendar"));
        LastUsedTasks = EmailStoreConfigurationXml.ParseDate((string?)element.Attribute("lastTimeUsedTasks"));
        LastUsedContacts = EmailStoreConfigurationXml.ParseDate((string?)element.Attribute("lastTimeUsedContacts"));
        LastUsedJournal = EmailStoreConfigurationXml.ParseDate((string?)element.Attribute("lastTimeUsedJournal"));
        LastUsedNotes = EmailStoreConfigurationXml.ParseDate((string?)element.Attribute("lastTimeUsedNotes"));
        LastSessionUsed = ParseInt((string?)element.Attribute("lastSessionUsed"));
        string? guid = (string?)element.Attribute("guid");
        Id = Guid.TryParse(guid, out Guid parsed) ? parsed : (Guid?)null;
        RenameOnFirstUse = string.Equals((string?)element.Attribute("renameOnFirstUse"), "1", StringComparison.Ordinal);
        UnknownAttributes = element.Attributes()
            .Where(attribute => !KnownAttributes.Contains(attribute.Name.LocalName))
            .ToDictionary(attribute => attribute.Name.ToString(), attribute => attribute.Value, StringComparer.Ordinal);
    }

    private static readonly ISet<string> KnownAttributes = new HashSet<string>(StringComparer.Ordinal) {
        "name", "color", "keyboardShortcut", "usageCount", "lastTimeUsedNotes", "lastTimeUsedJournal",
        "lastTimeUsedContacts", "lastTimeUsedTasks", "lastTimeUsedCalendar", "lastTimeUsedMail",
        "lastTimeUsed", "lastSessionUsed", "guid", "renameOnFirstUse"
    };

    /// <summary>Case-insensitively unique display name.</summary>
    public string Name { get; }
    /// <summary>Outlook color index; -1 through 24 are defined.</summary>
    public int? Color { get; }
    /// <summary>Outlook shortcut index; 0 through 11 are defined.</summary>
    public uint? KeyboardShortcut { get; }
    /// <summary>Optional usage counter.</summary>
    public uint? UsageCount { get; }
    /// <summary>Most recent use across item families.</summary>
    public DateTimeOffset? LastUsed { get; }
    /// <summary>Most recent use on a mail item.</summary>
    public DateTimeOffset? LastUsedMail { get; }
    /// <summary>Most recent use on a calendar item.</summary>
    public DateTimeOffset? LastUsedCalendar { get; }
    /// <summary>Most recent use on a task.</summary>
    public DateTimeOffset? LastUsedTasks { get; }
    /// <summary>Most recent use on a contact.</summary>
    public DateTimeOffset? LastUsedContacts { get; }
    /// <summary>Most recent use on a journal item.</summary>
    public DateTimeOffset? LastUsedJournal { get; }
    /// <summary>Most recent use on a note.</summary>
    public DateTimeOffset? LastUsedNotes { get; }
    /// <summary>Reserved Outlook session counter.</summary>
    public int? LastSessionUsed { get; }
    /// <summary>Stable category identity.</summary>
    public Guid? Id { get; }
    /// <summary>Whether Outlook should offer to rename this category on first use.</summary>
    public bool RenameOnFirstUse { get; }
    /// <summary>Attributes not interpreted by OfficeIMO and retained on rewrite.</summary>
    public IReadOnlyDictionary<string, string> UnknownAttributes { get; }

    internal bool IsValid => !string.IsNullOrWhiteSpace(Name) && Name.Length <= 255 && Name.IndexOf(',') < 0 &&
        Color is >= -1 and <= 24 && KeyboardShortcut is <= 11 && LastUsed.HasValue &&
        LastSessionUsed.HasValue && Id.HasValue;

    private static int? ParseInt(string? value) => int.TryParse(value, NumberStyles.Integer,
        CultureInfo.InvariantCulture, out int parsed) ? parsed : (int?)null;
    private static uint? ParseUInt(string? value) => uint.TryParse(value, NumberStyles.Integer,
        CultureInfo.InvariantCulture, out uint parsed) ? parsed : (uint?)null;
}

/// <summary>
/// Lossless Outlook master category list. Known category elements can be edited while unknown XML is preserved.
/// </summary>
public sealed class EmailStoreCategoryList {
    private const string MessageClass = "IPM.Configuration.CategoryList";
    private static readonly XNamespace CategoryNamespace = "CategoryList.xsd";
    private readonly XDocument _xml;
    private readonly IReadOnlyList<MapiProperty> _sourceProperties;
    private readonly string? _sourceSubject;
    private readonly int? _sourceCodePage;

    private EmailStoreCategoryList(XDocument xml, IReadOnlyList<MapiProperty>? sourceProperties = null,
        string? sourceSubject = null, int? sourceCodePage = null) {
        _xml = xml;
        _sourceProperties = sourceProperties ?? Array.Empty<MapiProperty>();
        _sourceSubject = sourceSubject;
        _sourceCodePage = sourceCodePage;
    }

    /// <summary>Creates an empty category list; add at least one category before serialization.</summary>
    public static EmailStoreCategoryList Create() {
        var root = new XElement(CategoryNamespace + "categories",
            new XAttribute("default", string.Empty),
            new XAttribute("lastSavedSession", "0"),
            new XAttribute("lastSavedTime", EmailStoreConfigurationXml.FormatDate(
                new DateTimeOffset(1601, 1, 1, 0, 0, 0, TimeSpan.Zero))));
        return new EmailStoreCategoryList(new XDocument(new XDeclaration("1.0", "utf-8", null), root));
    }

    /// <summary>Parses a category-list associated message using a strict XML byte bound.</summary>
    public static EmailStoreCategoryList Parse(EmailDocument document, int maxXmlBytes = 4 * 1024 * 1024) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        byte[] bytes = document.Mapi.GetValueOrDefault(MapiKnownProperties.PidTag.RoamingXmlStream) ??
            throw new InvalidDataException("The category-list message has no PidTagRoamingXmlStream value.");
        XDocument xml = EmailStoreConfigurationXml.Parse(bytes, maxXmlBytes, "The category list");
        if (xml.Root == null || xml.Root.Name != CategoryNamespace + "categories") {
            throw new InvalidDataException("The category list does not have the CategoryList.xsd categories root.");
        }
        return new EmailStoreCategoryList(xml,
            document.MapiProperties.Select(MapiPropertySnapshot.Clone).ToArray(),
            document.Subject, document.OutlookCodePage);
    }

    /// <summary>Every parsed category in XML order.</summary>
    public IReadOnlyList<EmailStoreCategoryDefinition> Categories => Root.Elements(CategoryNamespace + "category")
        .Select(element => new EmailStoreCategoryDefinition(element)).ToArray();

    /// <summary>One-click default category name, or an empty string.</summary>
    public string DefaultCategoryName {
        get => (string?)Root.Attribute("default") ?? string.Empty;
        set {
            if (value == null) throw new ArgumentNullException(nameof(value));
            if (value.Length != 0 && !Categories.Any(category =>
                string.Equals(category.Name, value, StringComparison.OrdinalIgnoreCase))) {
                throw new ArgumentException("The default category must exist in this category list.", nameof(value));
            }
            Root.SetAttributeValue("default", value);
        }
    }

    /// <summary>Last saved UTC time declared by the XML.</summary>
    public DateTimeOffset? LastSavedAt =>
        EmailStoreConfigurationXml.ParseDate((string?)Root.Attribute("lastSavedTime"));

    /// <summary>True when the known protocol envelope and required category fields are valid.</summary>
    public bool IsProtocolEnvelopeValid {
        get {
            IReadOnlyList<EmailStoreCategoryDefinition> categories = Categories;
            return categories.Count > 0 && categories.All(category => category.IsValid) &&
                categories.Select(category => category.Name).Distinct(StringComparer.OrdinalIgnoreCase).Count() ==
                    categories.Count &&
                Root.Attribute("default") != null && Root.Attribute("lastSavedSession") != null && LastSavedAt.HasValue;
        }
    }

    /// <summary>Adds or replaces one category while retaining unknown attributes on an existing element.</summary>
    public EmailStoreCategoryDefinition Set(string name, int color = -1, uint keyboardShortcut = 0,
        Guid? id = null, uint? usageCount = null, DateTimeOffset? lastUsed = null,
        bool renameOnFirstUse = false) {
        ValidateName(name);
        if (color < -1 || color > 24) throw new ArgumentOutOfRangeException(nameof(color));
        if (keyboardShortcut > 11) throw new ArgumentOutOfRangeException(nameof(keyboardShortcut));
        XElement? element = Root.Elements(CategoryNamespace + "category").FirstOrDefault(candidate =>
            string.Equals((string?)candidate.Attribute("name"), name, StringComparison.OrdinalIgnoreCase));
        Guid? existingId = null;
        if (element != null && Guid.TryParse((string?)element.Attribute("guid"), out Guid parsedId)) {
            existingId = parsedId;
        }
        if (element == null) {
            element = new XElement(CategoryNamespace + "category");
            Root.Add(element);
        }
        DateTimeOffset used = (lastUsed ?? new DateTimeOffset(1601, 1, 1, 0, 0, 0, TimeSpan.Zero)).ToUniversalTime();
        element.SetAttributeValue("name", name);
        element.SetAttributeValue("color", color.ToString(CultureInfo.InvariantCulture));
        element.SetAttributeValue("keyboardShortcut", keyboardShortcut.ToString(CultureInfo.InvariantCulture));
        if (usageCount.HasValue) element.SetAttributeValue("usageCount", usageCount.Value.ToString(CultureInfo.InvariantCulture));
        else element.Attribute("usageCount")?.Remove();
        element.SetAttributeValue("lastTimeUsed", EmailStoreConfigurationXml.FormatDate(used));
        element.SetAttributeValue("lastSessionUsed", "0");
        element.SetAttributeValue("guid", string.Concat("{", (id ?? existingId ?? Guid.NewGuid()).ToString("D").ToUpperInvariant(), "}"));
        element.SetAttributeValue("renameOnFirstUse", renameOnFirstUse ? "1" : "0");
        return new EmailStoreCategoryDefinition(element);
    }

    /// <summary>Removes one category case-insensitively.</summary>
    public bool Remove(string name) {
        ValidateName(name);
        XElement? element = Root.Elements(CategoryNamespace + "category").FirstOrDefault(candidate =>
            string.Equals((string?)candidate.Attribute("name"), name, StringComparison.OrdinalIgnoreCase));
        if (element == null) return false;
        element.Remove();
        if (string.Equals(DefaultCategoryName, name, StringComparison.OrdinalIgnoreCase)) {
            Root.SetAttributeValue("default", string.Empty);
        }
        return true;
    }

    /// <summary>Serializes the category XML, updating its last-saved time on the serialized snapshot.</summary>
    public byte[] ToXml(DateTimeOffset? savedAt = null) {
        ValidateForWrite();
        var copy = new XDocument(_xml);
        copy.Root!.SetAttributeValue("lastSavedTime",
            EmailStoreConfigurationXml.FormatDate((savedAt ?? DateTimeOffset.UtcNow).ToUniversalTime()));
        return EmailStoreConfigurationXml.Serialize(copy);
    }

    /// <summary>Creates an Outlook-compatible FAI document for an associated Calendar-folder write.</summary>
    public EmailDocument ToAssociatedDocument(DateTimeOffset? savedAt = null) {
        DateTimeOffset timestamp = (savedAt ?? DateTimeOffset.UtcNow).ToUniversalTime();
        var document = new EmailDocument {
            MessageClass = MessageClass,
            Subject = _sourceSubject,
            OutlookCodePage = _sourceCodePage
        };
        foreach (MapiProperty property in _sourceProperties) document.MapiProperties.Add(MapiPropertySnapshot.Clone(property));
        document.Mapi.Set(MapiKnownProperties.PidTag.MessageClass, MessageClass);
        int flags = document.Mapi.GetNullableValue(MapiKnownProperties.PidTag.RoamingDatatypes) ?? 0;
        document.Mapi.Set(MapiKnownProperties.PidTag.RoamingDatatypes, flags | 0x00000002);
        document.Mapi.Set(MapiKnownProperties.PidTag.RoamingXmlStream, ToXml(timestamp));
        document.MessageMetadata.ModifiedDate = timestamp;
        return document;
    }

    private XElement Root => _xml.Root!;

    private void ValidateForWrite() {
        if (!IsProtocolEnvelopeValid) {
            throw new InvalidOperationException("The category list cannot be written until required fields are valid and at least one unique category exists.");
        }
    }

    private static void ValidateName(string name) {
        if (name == null) throw new ArgumentNullException(nameof(name));
        if (string.IsNullOrWhiteSpace(name) || name.Length > 255 || name.IndexOf(',') >= 0) {
            throw new ArgumentException("A category name must contain 1-255 non-whitespace characters and cannot contain a comma.", nameof(name));
        }
    }
}
