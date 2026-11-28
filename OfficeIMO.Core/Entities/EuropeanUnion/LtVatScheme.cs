using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

/// <summary>Lithuanian VAT validation from <see href="https://www.bmf.gv.at/dam/jcr:d6794f8f-d321-43df-9840-1a841f9bf5dc/BMF_UID_Konstruktionsregeln_Stand_November%202020.pdf">the BMF construction rules</see>.</summary>
internal sealed class LtVatScheme : PrefixedVatScheme {
    internal LtVatScheme() : base("LT", 9, 12, EntityCheckKind.Mod11, BmfSpecification) { }

    protected override EntityValidationResult ValidateNational(string number) {
        if (number.Length is not (9 or 12)) return Result(EntityValidationStatus.InvalidLength);
        if (!EntityCharacters.IsDigits(number, 0, number.Length) || number[number.Length - 2] != '1')
            return Result(EntityValidationStatus.InvalidFormat);

        bool shorter = number.Length == 9;
        int check = CheckDigits.WeightedSum(number, shorter ? "12345678" : "12345678912") % 11;
        if (check == 10) check = CheckDigits.WeightedSum(number, shorter ? "34567891" : "34567891234") % 11 % 10;
        return NationalCheck(number, check == number[number.Length - 1] - '0');
    }
}
