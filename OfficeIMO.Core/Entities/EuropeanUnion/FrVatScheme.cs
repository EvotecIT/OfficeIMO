using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

/// <summary>French VAT validation from <see href="https://meta.cdq.com/Public:Rule_Category/IDENTIFIER_CHECKDIGIT">CDQ’s documented numeric and alphanumeric French VAT checks</see>.</summary>
internal sealed class FrVatScheme : PrefixedVatScheme {
    internal FrVatScheme() : base("FR", 11, 11, EntityCheckKind.Other, CdqSpecification) { }

    protected override EntityCheckProvenance DefaultCheckProvenance => EntityCheckProvenance.CommunityDocumented;

    protected override EntityValidationResult ValidateNational(string number) {
        if (!EntityCharacters.IsDigits(number, 2, 9) || "IO".IndexOf(number[0]) >= 0 || "IO".IndexOf(number[1]) >= 0)
            return Result(EntityValidationStatus.InvalidFormat);

        int siren = EntityCharacters.Number(number, 2, 9);
        if (EntityCharacters.IsDigits(number, 0, 2)) {
            return NationalCheck(number, (12 + 3 * (siren % 97)) % 97 == EntityCharacters.Number(number, 0, 2));
        }
        int first = KeyValue(number[0]);
        int second = KeyValue(number[1]);
        int key = first < 10 ? first * 24 + second - 10 : first * 34 + second - 100;
        return NationalCheck(number, key % 11 == (siren + key / 11 + 1) % 11);
    }

    private static int KeyValue(char value) => "0123456789ABCDEFGHJKLMNPQRSTUVWXYZ".IndexOf(value);
}
