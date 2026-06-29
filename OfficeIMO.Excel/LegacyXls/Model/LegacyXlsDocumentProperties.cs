namespace OfficeIMO.Excel.LegacyXls.Model {
    internal sealed class LegacyXlsDocumentProperties {
        private readonly Dictionary<string, LegacyXlsDocumentPropertyValue> _customProperties = new(StringComparer.OrdinalIgnoreCase);

        internal string? Title { get; set; }

        internal string? Subject { get; set; }

        internal string? Creator { get; set; }

        internal string? Keywords { get; set; }

        internal string? Description { get; set; }

        internal string? Category { get; set; }

        internal string? LastModifiedBy { get; set; }

        internal string? Revision { get; set; }

        internal DateTime? Created { get; set; }

        internal DateTime? Modified { get; set; }

        internal DateTime? LastPrinted { get; set; }

        internal string? Company { get; set; }

        internal string? Manager { get; set; }

        internal IReadOnlyDictionary<string, LegacyXlsDocumentPropertyValue> CustomProperties => _customProperties;

        internal bool HasAnyProperties =>
            HasBuiltInProperties
            || HasApplicationProperties
            || _customProperties.Count > 0;

        internal bool HasBuiltInProperties =>
            Title != null
            || Subject != null
            || Creator != null
            || Keywords != null
            || Description != null
            || Category != null
            || LastModifiedBy != null
            || Revision != null
            || Created.HasValue
            || Modified.HasValue
            || LastPrinted.HasValue;

        internal bool HasApplicationProperties => Company != null || Manager != null;

        internal void SetCustomProperty(string name, LegacyXlsDocumentPropertyValue value) {
            if (string.IsNullOrWhiteSpace(name)) {
                return;
            }

            _customProperties[name.Trim()] = value;
        }
    }
}
