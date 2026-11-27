namespace OfficeIMO.Entities.Internal;

/// <summary>Registered IBAN country and BBAN structures.</summary>
internal static class IbanStructure {
    // SWIFT IBAN Registry release 102 (June 2026), registered BBAN formats:
    // https://www.swift.com/swift-resource/9606/download
    // Fixed-length n/a/c fields denote digits, letters, and alphanumeric characters.
    // Unknown prefixes are not inferred to be IBANs. This does not verify account existence.
    private static string? GetBbanFormat(string country) => country switch {
        "AD" => "4!n4!n12!c",
        "AE" => "3!n16!n",
        "AL" => "8!n16!c",
        "AT" => "5!n11!n",
        "AZ" => "4!a20!c",
        "BA" => "3!n3!n8!n2!n",
        "BE" => "3!n7!n2!n",
        "BG" => "4!a4!n2!n8!c",
        "BH" => "4!a14!c",
        "BI" => "5!n5!n11!n2!n",
        "BR" => "8!n5!n10!n1!a1!c",
        "BY" => "4!c4!n16!c",
        "CH" => "5!n12!c",
        "CR" => "4!n14!n",
        "CY" => "3!n5!n16!c",
        "CZ" => "4!n6!n10!n",
        "DE" => "8!n10!n",
        "DJ" => "5!n5!n11!n2!n",
        "DK" => "4!n9!n1!n",
        "DO" => "4!c20!n",
        "EE" => "2!n14!n",
        "EG" => "4!n4!n17!n",
        "ES" => "4!n4!n1!n1!n10!n",
        "FI" => "3!n11!n",
        "FK" => "2!a12!n",
        "FO" => "4!n9!n1!n",
        "FR" => "5!n5!n11!c2!n",
        "GB" => "4!a6!n8!n",
        "GE" => "2!a16!n",
        "GI" => "4!a15!c",
        "GL" => "4!n9!n1!n",
        "GR" => "3!n4!n16!c",
        "GT" => "4!c20!c",
        "HN" => "4!a20!n",
        "HR" => "7!n10!n",
        "HU" => "3!n4!n1!n15!n1!n",
        "IE" => "4!a6!n8!n",
        "IL" => "3!n3!n13!n",
        "IQ" => "4!a3!n12!n",
        "IS" => "4!n2!n6!n10!n",
        "IT" => "1!a5!n5!n12!c",
        "JO" => "4!a4!n18!c",
        "KW" => "4!a22!c",
        "KZ" => "3!n13!c",
        "LB" => "4!n20!c",
        "LC" => "4!a24!c",
        "LI" => "5!n12!c",
        "LT" => "5!n11!n",
        "LU" => "3!n13!c",
        "LV" => "4!a13!c",
        "LY" => "3!n3!n15!n",
        "MC" => "5!n5!n11!c2!n",
        "MD" => "2!c18!c",
        "ME" => "3!n13!n2!n",
        "MK" => "3!n10!c2!n",
        "MN" => "4!n12!n",
        "MR" => "5!n5!n11!n2!n",
        "MT" => "4!a5!n18!c",
        "MU" => "4!a2!n2!n12!n3!n3!a",
        "NI" => "4!a20!n",
        "NL" => "4!a10!n",
        "NO" => "4!n6!n1!n",
        "OM" => "3!n16!c",
        "PK" => "4!a16!c",
        "PL" => "8!n16!n",
        "PS" => "4!a21!c",
        "PT" => "4!n4!n11!n2!n",
        "QA" => "4!a21!c",
        "RO" => "4!a16!c",
        "RS" => "3!n13!n2!n",
        "RU" => "9!n5!n15!c",
        "SA" => "2!n18!c",
        "SC" => "4!a2!n2!n16!n3!a",
        "SD" => "2!n12!n",
        "SE" => "3!n16!n1!n",
        "SI" => "5!n8!n2!n",
        "SK" => "4!n6!n10!n",
        "SM" => "1!a5!n5!n12!c",
        "SO" => "4!n3!n12!n",
        "ST" => "4!n4!n11!n2!n",
        "SV" => "4!a20!n",
        "TL" => "3!n14!n2!n",
        "TN" => "2!n3!n13!n2!n",
        "TR" => "5!n1!n16!c",
        "UA" => "6!n19!c",
        "VA" => "3!n15!n",
        "VG" => "4!a16!n",
        "XK" => "4!n10!n2!n",
        "YE" => "4!a4!n18!c",
        _ => null
    };

    internal static bool HasRegisteredStructure(string iban) {
        string? format = GetBbanFormat(iban.Substring(0, 2));
        if (format == null) return false;

        int position = 4;
        for (int index = 0; index < format.Length;) {
            int length = 0;
            while (EntityCharacters.IsDigit(format[index])) {
                length = length * 10 + format[index++] - '0';
            }
            index++; // The registry's fixed-length marker (!).
            char kind = format[index++];
            if (length > iban.Length - position) return false;
            for (int end = position + length; position < end; position++) {
                char current = iban[position];
                bool digit = EntityCharacters.IsDigit(current);
                bool letter = EntityCharacters.IsLetter(current);
                if (!(kind == 'n' ? digit : kind == 'a' ? letter : digit || letter)) return false;
            }
        }
        return position == iban.Length;
    }
}
