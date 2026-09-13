using System.Globalization;

namespace OfficeIMO.Project;

/// <summary>Resolves modern shared or legacy inline scalar lookup references without conflating the two source representations.</summary>
internal static class ProjectCustomFieldLookup {
    internal static string? Text(ProjectDocument document, ProjectCustomFieldDefinition? definition, ProjectCustomFieldValue? stored) {
        string? text = stored?.Value;
        var table = definition == null ? null : document.GetCustomFieldLookupTable(definition);
        if (stored != null && (stored.ValueId != null || stored.ValueGuid != null)) {
            string? resolved;
            if (table != null) {
                var matches = table.Values.Where(v => (stored.ValueId == null || v.ValueId?.ToString(CultureInfo.InvariantCulture) == stored.ValueId) &&
                    (stored.ValueGuid == null || ProjectOutlineCodeIndex.SameGuid(v.Guid, stored.ValueGuid))).ToArray();
                if (matches.Length != 1) throw new InvalidDataException("The shared lookup reference has no unique modeled entry.");
                if ((matches[0].ParentValueId ?? 0) != 0) throw new NotSupportedException("Scalar custom fields support flat shared lookup values.");
                resolved = matches[0].Value;
            } else {
                var matches = definition?.LookupValues.Where(v => (stored.ValueId == null || v.Id?.ToString(CultureInfo.InvariantCulture) == stored.ValueId) &&
                    (stored.ValueGuid == null || ProjectOutlineCodeIndex.SameGuid(v.Guid, stored.ValueGuid))).ToArray();
                if (matches == null || matches.Length != 1) throw new InvalidDataException("The lookup reference has no unique modeled entry.");
                resolved = matches[0].Value;
            }
            if (text != null && text != resolved) throw new InvalidDataException("The stored custom value disagrees with its lookup reference.");
            text = resolved;
        }
        bool restricted = definition?.RestrictValues == true || table?.OnlyTableValuesAllowed == true;
        if (restricted && (text == null || !(table != null ? table.Values.Any(v => v.Value == text) : definition!.LookupValues.Any(v => v.Value == text))))
            throw new InvalidDataException("The custom value is outside its restricted lookup list.");
        return text;
    }
}
