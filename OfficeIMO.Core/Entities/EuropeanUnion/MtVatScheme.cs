using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

/// <summary>Maltese VAT validation from <see href="https://meta.cdq.com/Public:Rule_Category/IDENTIFIER_CHECKDIGIT">CDQ’s documented Maltese VAT check</see>.</summary>
internal sealed class MtVatScheme : PrefixedVatScheme {
    internal MtVatScheme() : base("MT", 8, 8, EntityCheckKind.Other, CdqSpecification) { }

    protected override EntityCheckProvenance DefaultCheckProvenance => EntityCheckProvenance.CommunityDocumented;

    protected override EntityValidationResult ValidateNational(string number) =>
        NationalDigits(number, static value => 37 - CheckDigits.WeightedSum(value, "346789") % 37 == EntityCharacters.Number(value, 6, 2));
}
