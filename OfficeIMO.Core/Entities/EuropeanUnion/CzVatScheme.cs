using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

/// <summary>Czech VAT formats and IČO check from <see href="https://vatdb.com/guides/validate-cz-vat-number/">VatDB's documented check rules</see>.</summary>
internal sealed class CzVatScheme : PrefixedVatScheme {
    internal CzVatScheme() : base("CZ", 8, 10, EntityCheckKind.Mod11, "https://vatdb.com/guides/validate-cz-vat-number/") { }

    protected override EntityCheckProvenance DefaultCheckProvenance => EntityCheckProvenance.CommunityDocumented;

    protected override EntityValidationResult ValidateNational(string number) {
        if (!EntityCharacters.IsDigits(number, 0, number.Length)) return Result(EntityValidationStatus.InvalidFormat);
        if (number.Length != 8) return NationalResult(EntityValidationStatus.Valid, number);
        int check = (11 - CheckDigits.WeightedSum(number, "8765432") % 11) % 10;
        return NationalCheck(number, check == number[7] - '0');
    }
}
