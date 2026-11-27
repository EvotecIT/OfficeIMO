using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

/// <summary>Cypriot VAT validation from <see href="https://digitoverificador.com/en/chipre/vat-chipre-check-digit">Dígito Verificador’s positional check</see>.</summary>
internal sealed class CyVatScheme : PrefixedVatScheme {
    internal CyVatScheme() : base("CY", 9, 9, EntityCheckKind.Other, "https://digitoverificador.com/en/chipre/vat-chipre-check-digit") { }

    private static readonly int[] OddPositionValues = { 1, 0, 5, 7, 9, 13, 15, 17, 19, 21 };

    protected override EntityCheckProvenance DefaultCheckProvenance => EntityCheckProvenance.CommunityDocumented;

    protected override EntityValidationResult ValidateNational(string number) {
        if (!EntityCharacters.IsDigits(number, 0, 8) || !EntityCharacters.IsLetter(number[8]))
            return Result(EntityValidationStatus.InvalidFormat);

        int sum = 0;
        for (int index = 0; index < 8; index += 2) {
            sum += OddPositionValues[number[index] - '0'] + number[index + 1] - '0';
        }
        return NationalCheck(number, 'A' + sum % 26 == number[8]);
    }
}
