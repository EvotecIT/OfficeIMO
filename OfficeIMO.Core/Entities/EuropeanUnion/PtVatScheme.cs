using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

/// <summary>Portuguese VAT validation from <see href="https://www.bmf.gv.at/dam/jcr:d6794f8f-d321-43df-9840-1a841f9bf5dc/BMF_UID_Konstruktionsregeln_Stand_November%202020.pdf">the BMF construction rules</see>.</summary>
internal sealed class PtVatScheme : PrefixedVatScheme {
    internal PtVatScheme() : base("PT", 9, 9, EntityCheckKind.Mod11, BmfSpecification) { }

    protected override EntityValidationResult ValidateNational(string number) {
        if (!EntityCharacters.IsDigits(number, 0, 9) || number[0] == '0') return Result(EntityValidationStatus.InvalidFormat);
        int check = 11 - CheckDigits.WeightedSum(number, "98765432") % 11;
        return NationalCheck(number, (check > 9 ? 0 : check) == number[8] - '0');
    }
}
