namespace OfficeIMO.Email;

internal static partial class VCardCodec {
    private static bool HasUnsupportedEmailSlot(VCardProperty property, int index, string types) {
        bool preferred = ContainsType(types, "PREF") || property.Parameters.ContainsKey("PREF");
        bool expectedSlot = index == 0
            ? ContainsType(types, "WORK") && preferred
            : index == 1
                ? ContainsType(types, "HOME") && !preferred
                : ContainsType(types, "OTHER") && !preferred;
        string[] allowedTypes = index == 0
            ? new[] { "WORK", "PREF", "INTERNET" }
            : index == 1
                ? new[] { "HOME", "INTERNET" }
                : new[] { "OTHER", "INTERNET" };
        return !expectedSlot || HasUnsupportedParameters(property, "TYPE", "PREF") ||
            HasUnsupportedTypeTokens(types, allowedTypes);
    }

    private static bool HasUnsupportedPhoneSlot(VCardProperty property, VCardPhoneSlot slot, int count,
        string types) {
        if (HasUnsupportedParameters(property, "TYPE", "PREF")) return true;
        bool preferred = ContainsType(types, "PREF") || property.Parameters.ContainsKey("PREF");
        string[] allowedTypes;
        bool requiredTypePresent;
        switch (slot) {
            case VCardPhoneSlot.Mobile:
                allowedTypes = new[] { "CELL", "VOICE" }; requiredTypePresent = ContainsType(types, "CELL"); break;
            case VCardPhoneSlot.Assistant:
                allowedTypes = new[] { "X-ASSISTANT" }; requiredTypePresent = ContainsType(types, "X-ASSISTANT"); break;
            case VCardPhoneSlot.Company:
                allowedTypes = new[] { "WORK", "X-COMPANY" }; requiredTypePresent = ContainsType(types, "X-COMPANY"); break;
            case VCardPhoneSlot.Car:
                allowedTypes = new[] { "CAR" }; requiredTypePresent = ContainsType(types, "CAR"); break;
            case VCardPhoneSlot.Radio:
                allowedTypes = new[] { "X-RADIO" }; requiredTypePresent = ContainsType(types, "X-RADIO"); break;
            case VCardPhoneSlot.Callback:
                allowedTypes = new[] { "X-CALLBACK" }; requiredTypePresent = ContainsType(types, "X-CALLBACK"); break;
            case VCardPhoneSlot.Telex:
                allowedTypes = new[] { "X-TELEX" }; requiredTypePresent = ContainsType(types, "X-TELEX"); break;
            case VCardPhoneSlot.Text:
                allowedTypes = new[] { "TEXT" }; requiredTypePresent = ContainsType(types, "TEXT"); break;
            case VCardPhoneSlot.Isdn:
                allowedTypes = new[] { "ISDN" }; requiredTypePresent = ContainsType(types, "ISDN"); break;
            case VCardPhoneSlot.PrimaryFax:
                allowedTypes = new[] { "PREF", "FAX" }; requiredTypePresent = preferred && ContainsType(types, "FAX"); break;
            case VCardPhoneSlot.HomeFax:
                allowedTypes = new[] { "HOME", "FAX" }; requiredTypePresent = ContainsType(types, "HOME") && ContainsType(types, "FAX"); break;
            case VCardPhoneSlot.BusinessFax:
                allowedTypes = new[] { "WORK", "FAX" }; requiredTypePresent = ContainsType(types, "WORK") && ContainsType(types, "FAX"); break;
            case VCardPhoneSlot.Home:
                allowedTypes = new[] { "HOME", "VOICE" }; requiredTypePresent = ContainsType(types, "HOME") && !preferred; break;
            case VCardPhoneSlot.Pager:
                allowedTypes = new[] { "PAGER" }; requiredTypePresent = ContainsType(types, "PAGER"); break;
            case VCardPhoneSlot.Other:
                allowedTypes = new[] { "OTHER", "VOICE" }; requiredTypePresent = ContainsType(types, "OTHER") && !preferred; break;
            case VCardPhoneSlot.Work:
                allowedTypes = new[] { "WORK", "VOICE" }; requiredTypePresent = ContainsType(types, "WORK") && !preferred; break;
            default:
                allowedTypes = new[] { "PREF", "VOICE" }; requiredTypePresent = preferred && count == 1; break;
        }
        return !requiredTypePresent || HasUnsupportedTypeTokens(types, allowedTypes);
    }

    private static bool HasAddressSlotOverflow(IEnumerable<VCardProperty> properties) =>
        HasAddressSlotOverflow(properties, "ADR") || HasAddressSlotOverflow(properties, "LABEL");

    private static bool HasUnprojectedAddressComponents(IEnumerable<VCardProperty> properties) =>
        properties.Where(property => property.Name == "ADR").Any(property => {
            string[] values = SplitEscaped(property.Value, ';');
            return !string.IsNullOrWhiteSpace(ValueAt(values, 1));
        });

    private static bool HasUnsupportedAddressTypes(IEnumerable<VCardProperty> properties) =>
        properties.Where(property => property.Name == "ADR" || property.Name == "LABEL").Any(property => {
            string types = property.Parameters.TryGetValue("TYPE", out string? type) ? type : string.Empty;
            int addressTypes = new[] { "HOME", "WORK", "OTHER" }.Count(value => ContainsType(types, value));
            string[] allowedParameters = property.Name == "ADR" ? new[] { "TYPE", "LABEL" } : new[] { "TYPE" };
            return addressTypes > 1 || HasUnsupportedParameters(property, allowedParameters) ||
                HasUnsupportedTypeTokens(types, "HOME", "WORK", "OTHER");
        });

    private static bool HasAddressSlotOverflow(IEnumerable<VCardProperty> properties, string propertyName) {
        int homeCount = 0;
        int workCount = 0;
        int otherCount = 0;
        foreach (VCardProperty property in properties.Where(property => property.Name == propertyName)) {
            string types = property.Parameters.TryGetValue("TYPE", out string? type) ? type : string.Empty;
            if (ContainsType(types, "HOME")) homeCount++;
            else if (ContainsType(types, "WORK")) workCount++;
            else otherCount++;
        }
        return homeCount > 1 || workCount > 2 || otherCount > 1;
    }

    private static bool HasUrlSlotOverflow(IEnumerable<VCardProperty> properties) {
        int homeCount = 0;
        int workCount = 0;
        foreach (VCardProperty property in properties.Where(property => property.Name == "URL")) {
            string types = property.Parameters.TryGetValue("TYPE", out string? type) ? type : string.Empty;
            if (ContainsType(types, "HOME")) homeCount++;
            else workCount++;
        }
        return homeCount > 1 || workCount > 1;
    }

    private static bool HasUnsupportedUrlTypes(IEnumerable<VCardProperty> properties) =>
        properties.Where(property => property.Name == "URL").Any(property => {
            string types = property.Parameters.TryGetValue("TYPE", out string? type) ? type : string.Empty;
            return HasUnsupportedParameters(property, "TYPE") ||
                HasUnsupportedTypeTokens(types, "HOME", "WORK") ||
                ContainsType(types, "HOME") && ContainsType(types, "WORK");
        });

    private static bool HasUnsupportedParameters(VCardProperty property, params string[] allowed) =>
        property.Parameters.Keys.Any(parameter => !allowed.Any(value =>
            value.Equals(parameter, StringComparison.OrdinalIgnoreCase)));

    private static bool HasUnsupportedTypeTokens(string types, params string[] allowed) => types.Split(',')
        .Select(value => value.Trim()).Where(value => value.Length > 0)
        .Any(value => !allowed.Any(allowedType => allowedType.Equals(value, StringComparison.OrdinalIgnoreCase)));
}
