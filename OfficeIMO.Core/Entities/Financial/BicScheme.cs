using System;
using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

/// <summary>Business identifier code structure from <see href="https://www.swift.com/standards/data-standards/bic-business-identifier-code">SWIFT</see> and ISO 9362.</summary>
internal sealed class BicScheme : CompactEntityScheme {
    // Assigned alpha-2 codes from ISO 3166-1 (https://www.iso.org/obp/ui/#search/code/).
    private const string CountryCodes =
        " AD AE AF AG AI AL AM AO AQ AR AS AT AU AW AX AZ BA BB BD BE BF BG BH BI BJ BL BM BN BO BQ BR BS BT BV BW BY BZ" +
        " CA CC CD CF CG CH CI CK CL CM CN CO CR CU CV CW CX CY CZ DE DJ DK DM DO DZ EC EE EG EH ER ES ET FI FJ FK FM FO FR" +
        " GA GB GD GE GF GG GH GI GL GM GN GP GQ GR GS GT GU GW GY HK HM HN HR HT HU ID IE IL IM IN IO IQ IR IS IT JE JM JO JP" +
        " KE KG KH KI KM KN KP KR KW KY KZ LA LB LC LI LK LR LS LT LU LV LY MA MC MD ME MF MG MH MK ML MM MN MO MP MQ MR MS MT MU MV MW MX MY MZ" +
        " NA NC NE NF NG NI NL NO NP NR NU NZ OM PA PE PF PG PH PK PL PM PN PR PS PT PW PY QA RE RO RS RU RW" +
        " SA SB SC SD SE SG SH SI SJ SK SL SM SN SO SR SS ST SV SX SY SZ TC TD TF TG TH TJ TK TL TM TN TO TR TT TV TW TZ" +
        " UA UG UM US UY UZ VA VC VE VG VI VN VU WF WS YE YT ZA ZM ZW ";

    public override string Id => "BIC";
    public override EntityKind Kind => EntityKind.BankIdentifier;
    public override string DisplayName => "BIC";
    public override string Issuer => "SWIFT (ISO 9362 registration authority)";
    public override Uri Specification { get; } = new("https://www.swift.com/standards/data-standards/bic-business-identifier-code");
    public override int MinLength => 8;
    public override int MaxLength => 11;
    public override EntitySensitivity Sensitivity => EntitySensitivity.Organizational;

    protected override EntityValidationResult ValidateCompact(string value) {
        if (value.Length != 8 && value.Length != 11) return Result(EntityValidationStatus.InvalidLength);
        if (!EntityCharacters.IsLetters(value, 0, 6) || CountryCodes.IndexOf(" " + value.Substring(4, 2) + " ", StringComparison.Ordinal) < 0) {
            return Result(EntityValidationStatus.InvalidFormat);
        }
        return Result(EntityValidationStatus.Valid, value);
    }
}
