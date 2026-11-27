using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

/// <summary>Cypriot VAT validation from <see href="https://vatdb.com/guides/validate-cy-vat-number/">VatDB’s format and checksum documentation</see>.</summary>
internal sealed class CyVatScheme : PrefixedVatScheme {
    internal CyVatScheme() : base("CY", 9, 9, EntityCheckKind.Other, "https://vatdb.com/guides/validate-cy-vat-number/") { }

    private static readonly int[] OddPositionValues = { 1, 0, 5, 7, 9, 13, 15, 17, 19, 21 };

    protected override EntityCheckProvenance DefaultCheckProvenance => EntityCheckProvenance.CommunityDocumented;

    protected override EntityValidationResult ValidateNational(string number) {
        if (!EntityCharacters.IsDigits(number, 0, 8) || !EntityCharacters.IsLetter(number[8]))
            return Result(EntityValidationStatus.InvalidFormat);

        if (number[0] == '1' && number[1] == '2') return NationalResult(EntityValidationStatus.NotAllocated, number);
        int sum = 0;
        for (int index = 0; index < 8; index += 2) {
            sum += OddPositionValues[number[index] - '0'] + number[index + 1] - '0';
        }
        return NationalCheck(number, 'A' + sum % 26 == number[8]);
    }
}
