using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

/// <summary>German VAT validation from <see href="https://www.bmf.gv.at/dam/jcr:d6794f8f-d321-43df-9840-1a841f9bf5dc/BMF_UID_Konstruktionsregeln_Stand_November%202020.pdf">the BMF construction rules</see>.</summary>
internal sealed class DeVatScheme : PrefixedVatScheme {
    internal DeVatScheme() : base("DE", 9, 9, EntityCheckKind.Other, BmfSpecification) { }

    protected override EntityValidationResult ValidateNational(string number) {
        if (!EntityCharacters.IsDigits(number, 0, 9) || number[0] == '0') return Result(EntityValidationStatus.InvalidFormat);
        return NationalCheck(number, CheckDigits.HybridMod11_10(number));
    }
}
