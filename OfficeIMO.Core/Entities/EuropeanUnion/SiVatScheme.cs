using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

/// <summary>Slovenian VAT validation from <see href="https://www.bmf.gv.at/dam/jcr:d6794f8f-d321-43df-9840-1a841f9bf5dc/BMF_UID_Konstruktionsregeln_Stand_November%202020.pdf">the BMF construction rules</see>.</summary>
internal sealed class SiVatScheme : PrefixedVatScheme {
    internal SiVatScheme() : base("SI", 8, 8, EntityCheckKind.Mod11, BmfSpecification) { }

    protected override EntityValidationResult ValidateNational(string number) {
        if (!EntityCharacters.IsDigits(number, 0, 8) || number[0] == '0') return Result(EntityValidationStatus.InvalidFormat);
        int check = 11 - CheckDigits.WeightedSum(number, "8765432") % 11;
        return NationalCheck(number, (check == 10 ? 0 : check) == number[7] - '0');
    }
}
