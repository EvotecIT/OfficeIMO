using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

/// <summary>Romanian VAT validation from <see href="https://meta.cdq.com/Public:Rule_Category/IDENTIFIER_CHECKDIGIT">CDQ’s documented Romanian VAT check</see>.</summary>
internal sealed class RoVatScheme : PrefixedVatScheme {
    internal RoVatScheme() : base("RO", 2, 10, EntityCheckKind.Mod11, CdqSpecification) { }

    protected override EntityCheckProvenance DefaultCheckProvenance => EntityCheckProvenance.CommunityDocumented;

    protected override EntityValidationResult ValidateNational(string number) {
        if (number[0] == '0') return Result(EntityValidationStatus.InvalidFormat);
        return NationalDigits(number, static value => CheckDigits.WeightedSum(value.PadLeft(10, '0'), "753217532") * 10 % 11 % 10 == value[value.Length - 1] - '0');
    }
}
