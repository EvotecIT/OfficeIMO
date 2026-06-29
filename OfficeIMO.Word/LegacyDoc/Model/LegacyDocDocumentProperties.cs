namespace OfficeIMO.Word.LegacyDoc.Model {
    internal sealed class LegacyDocDocumentProperties {
        private readonly Dictionary<string, LegacyDocDocumentPropertyValue> _customProperties = new(StringComparer.OrdinalIgnoreCase);

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

        internal IReadOnlyDictionary<string, LegacyDocDocumentPropertyValue> CustomProperties => _customProperties;

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

        internal int Count {
            get {
                int count = 0;
                if (Title != null) count++;
                if (Subject != null) count++;
                if (Creator != null) count++;
                if (Keywords != null) count++;
                if (Description != null) count++;
                if (Category != null) count++;
                if (LastModifiedBy != null) count++;
                if (Revision != null) count++;
                if (Created.HasValue) count++;
                if (Modified.HasValue) count++;
                if (LastPrinted.HasValue) count++;
                if (Company != null) count++;
                if (Manager != null) count++;
                return count + _customProperties.Count;
            }
        }

        internal void SetCustomProperty(string name, LegacyDocDocumentPropertyValue value) {
            if (string.IsNullOrWhiteSpace(name)) {
                return;
            }

            _customProperties[name.Trim()] = value;
        }
    }
}
