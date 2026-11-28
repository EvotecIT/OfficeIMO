using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

/// <summary>Latvian company VAT checksum from <see href="https://meta.cdq.com/Public:Rule_Category/IDENTIFIER_CHECKDIGIT">CDQ</see>.</summary>
internal sealed class LvVatScheme : PrefixedVatScheme {
    internal LvVatScheme() : base("LV", 11, 11, EntityCheckKind.Mod11, CdqSpecification) { }

    protected override EntityCheckProvenance DefaultCheckProvenance => EntityCheckProvenance.CommunityDocumented;

    protected override EntityValidationResult ValidateNational(string number) {
        if (!EntityCharacters.IsDigits(number, 0, 11)) return Result(EntityValidationStatus.InvalidFormat);
        if (number[0] < '4') return NationalResult(EntityValidationStatus.Valid, number);
        int sum = CheckDigits.WeightedSum(number, "9148302576") + (number[5] - '0') * 10;
        int check = (14 - sum % 11) % 11;
        return NationalCheck(number, check == number[10] - '0');
    }
}
