using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

/// <summary>Danish VAT validation from <see href="https://www.bmf.gv.at/dam/jcr:d6794f8f-d321-43df-9840-1a841f9bf5dc/BMF_UID_Konstruktionsregeln_Stand_November%202020.pdf">the BMF construction rules</see>.</summary>
internal sealed class DkVatScheme : PrefixedVatScheme {
    internal DkVatScheme() : base("DK", 8, 8, EntityCheckKind.Mod11, BmfSpecification) { }

    protected override EntityValidationResult ValidateNational(string number) {
        if (!EntityCharacters.IsDigits(number, 0, 8) || number[0] == '0') return Result(EntityValidationStatus.InvalidFormat);
        return NationalCheck(number, CheckDigits.WeightedSum(number, "27654321") % 11 == 0);
    }
}
