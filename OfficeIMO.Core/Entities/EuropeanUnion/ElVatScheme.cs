using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

/// <summary>Greek VAT validation from <see href="https://meta.cdq.com/Public:Rule_Category/IDENTIFIER_CHECKDIGIT">CDQ’s documented Greek VAT check</see>.</summary>
internal sealed class ElVatScheme : PrefixedVatScheme {
    internal ElVatScheme() : base("EL", 9, 9, EntityCheckKind.Mod11, CdqSpecification, "GR") { }

    protected override EntityCheckProvenance DefaultCheckProvenance => EntityCheckProvenance.CommunityDocumented;

    protected override EntityValidationResult ValidateNational(string number) =>
        NationalDigits(number, static value => PassesCheck(value));

    private static bool PassesCheck(string value) {
        int sum = 0;
        for (int index = 0; index < 8; index++) sum = sum * 2 + (value[index] - '0');
        return sum * 2 % 11 % 10 == value[8] - '0';
    }
}
