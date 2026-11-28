using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

/// <summary>Estonian VAT validation from <see href="https://www.bmf.gv.at/dam/jcr:d6794f8f-d321-43df-9840-1a841f9bf5dc/BMF_UID_Konstruktionsregeln_Stand_November%202020.pdf">the BMF construction rules</see>.</summary>
internal sealed class EeVatScheme : PrefixedVatScheme {
    internal EeVatScheme() : base("EE", 9, 9, EntityCheckKind.Mod10, BmfSpecification) { }

    protected override EntityValidationResult ValidateNational(string number) {
        if (!EntityCharacters.IsDigits(number, 0, 9) || number[0] != '1' || number[1] != '0') return Result(EntityValidationStatus.InvalidFormat);
        return NationalCheck(number, CheckDigits.WeightedSum(number, "371371371") % 10 == 0);
    }
}
