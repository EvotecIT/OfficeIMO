using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

/// <summary>Bulgarian VAT formats and company check from <see href="https://vatdb.com/guides/validate-bg-vat-number/">VatDB's documented check rules</see>.</summary>
internal sealed class BgVatScheme : PrefixedVatScheme {
    internal BgVatScheme() : base("BG", 9, 10, EntityCheckKind.Mod11, "https://vatdb.com/guides/validate-bg-vat-number/") { }

    protected override EntityCheckProvenance DefaultCheckProvenance => EntityCheckProvenance.CommunityDocumented;

    protected override EntityValidationResult ValidateNational(string number) {
        if (!EntityCharacters.IsDigits(number, 0, number.Length)) return Result(EntityValidationStatus.InvalidFormat);
        // Ten-digit service and legacy numbers share a shape with several individual schemes.
        if (number.Length == 10) return NationalResult(EntityValidationStatus.Valid, number);

        int check = CheckDigits.WeightedSum(number, "12345678") % 11;
        if (check == 10) check = (CheckDigits.WeightedSum(number, "34567890") + (number[7] - '0') * 10) % 11 % 10;
        return NationalCheck(number, check == number[8] - '0');
    }
}
