namespace OfficeIMO.Email;

public sealed partial class VCardDocument {
    private static readonly string[] SingletonProperties = {
        "VERSION", "N", "BDAY", "ANNIVERSARY", "GENDER", "KIND", "PRODID", "REV", "UID"
    };

    /// <summary>Validates version-specific stable vCard contracts while retaining extension properties.</summary>
    public IReadOnlyList<ContentLineValidationIssue> Validate() {
        var issues = new List<ContentLineValidationIssue>();
        if (_cards.Count == 0) {
            issues.Add(new ContentLineValidationIssue("VCARD_ROOT_REQUIRED",
                "The vCard document must contain at least one VCARD root.",
                ContentLineValidationSeverity.Error, "VCARD"));
        }
        foreach (ContentLineComponent card in _cards) {
            if (!string.Equals(card.Name, "VCARD", StringComparison.OrdinalIgnoreCase)) {
                issues.Add(new ContentLineValidationIssue("VCARD_ROOT_INVALID",
                    "Every vCard document root must be VCARD.",
                    ContentLineValidationSeverity.Error, card.Name));
            }
            foreach (ContentLineComponent nested in card.Components) {
                issues.Add(new ContentLineValidationIssue("VCARD_COMPONENT_NESTING_INVALID",
                    "VCARD cannot contain nested components.",
                    ContentLineValidationSeverity.Error, nested.Name));
            }
            VCardVersion version;
            try { version = GetVersion(card); }
            catch (InvalidDataException exception) {
                issues.Add(new ContentLineValidationIssue("VCARD_VERSION_INVALID", exception.Message,
                    ContentLineValidationSeverity.Error, card.Name, "VERSION"));
                continue;
            }
            if (card.Properties.Count == 0 ||
                !string.Equals(card.Properties[0].Name, "VERSION", StringComparison.OrdinalIgnoreCase)) {
                issues.Add(Issue("VCARD_VERSION_ORDER", "VERSION must be the first property after BEGIN:VCARD.",
                    ContentLineValidationSeverity.Error, card, "VERSION"));
            }
            Require(card, "FN", version == VCardVersion.V2_1
                ? ContentLineValidationSeverity.Warning
                : ContentLineValidationSeverity.Error, issues);
            Require(card, "N", version == VCardVersion.V4_0
                ? ContentLineValidationSeverity.Warning
                : ContentLineValidationSeverity.Error, issues);
            if (version != VCardVersion.V4_0 && card.GetProperties("FN").Skip(1).Any())
                issues.Add(Issue("VCARD_PROPERTY_CARDINALITY",
                    "FN must not occur more than once before vCard 4.0.",
                    ContentLineValidationSeverity.Error, card, "FN"));
            foreach (string propertyName in SingletonProperties) {
                if (card.GetProperties(propertyName).Skip(1).Any())
                    issues.Add(Issue("VCARD_PROPERTY_CARDINALITY",
                        propertyName + " must not occur more than once.", ContentLineValidationSeverity.Error,
                        card, propertyName));
            }
            if (version != VCardVersion.V4_0 && card.GetFirstProperty("ANNIVERSARY") != null)
                issues.Add(Issue("VCARD_PROPERTY_VERSION_MISMATCH",
                    "ANNIVERSARY is a vCard 4.0 property; use VERSION:4.0 or an explicitly named extension.",
                    ContentLineValidationSeverity.Error, card, "ANNIVERSARY"));
            string? kind = card.GetFirstProperty("KIND")?.Value;
            if (string.Equals(kind, "group", StringComparison.OrdinalIgnoreCase) &&
                !card.GetProperties("MEMBER").Any())
                issues.Add(Issue("VCARD_GROUP_MEMBER_REQUIRED", "A group card should contain MEMBER properties.",
                    ContentLineValidationSeverity.Warning, card, "MEMBER"));

            foreach (ContentLineProperty property in card.Properties) {
                if (version == VCardVersion.V4_0 && property.GetParameter("CHARSET") != null)
                    issues.Add(Issue("VCARD4_CHARSET_FORBIDDEN",
                        "vCard 4.0 is UTF-8 and does not use the CHARSET parameter.",
                        ContentLineValidationSeverity.Error, card, property.Name));
                bool hasEncodingParameter = property.Parameters.Any(parameter => string.Equals(
                    parameter.Name, "ENCODING", StringComparison.OrdinalIgnoreCase));
                if (version == VCardVersion.V4_0 && hasEncodingParameter)
                    issues.Add(Issue("VCARD4_ENCODING_FORBIDDEN",
                        "vCard 4.0 does not use the ENCODING parameter.",
                        ContentLineValidationSeverity.Error, card, property.Name));
                if (version == VCardVersion.V4_0) ValidatePreference(card, property, issues);
            }
        }
        return issues.AsReadOnly();
    }

    private static void ValidatePreference(ContentLineComponent card, ContentLineProperty property,
        ICollection<ContentLineValidationIssue> issues) {
        ContentLineParameter[] preferences = property.Parameters.Where(parameter =>
            string.Equals(parameter.Name, "PREF", StringComparison.OrdinalIgnoreCase)).ToArray();
        if (preferences.Length == 0) return;
        if (preferences.Length != 1 || preferences[0].Values.Count != 1) {
            issues.Add(Issue("VCARD4_PREF_CARDINALITY",
                "PREF must occur at most once and contain exactly one value.",
                ContentLineValidationSeverity.Error, card, property.Name));
        }
        foreach (string value in preferences.SelectMany(preference => preference.Values)) {
            if (!int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out int number) ||
                number < 1 || number > 100) {
                issues.Add(Issue("VCARD4_PREF_INVALID", "PREF must be an integer from 1 through 100.",
                    ContentLineValidationSeverity.Error, card, property.Name));
            }
        }
    }

    private static void Require(ContentLineComponent card, string propertyName,
        ContentLineValidationSeverity severity, ICollection<ContentLineValidationIssue> issues) {
        if (card.GetFirstProperty(propertyName) == null)
            issues.Add(Issue("VCARD_PROPERTY_REQUIRED", propertyName + " is required for this vCard version.",
                severity, card, propertyName));
    }

    private static ContentLineValidationIssue Issue(string code, string message,
        ContentLineValidationSeverity severity, ContentLineComponent card, string propertyName) =>
        new ContentLineValidationIssue(code, message, severity, card.Name, propertyName);
}
