using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

/// <summary>Finnish VAT validation from <see href="https://www.bmf.gv.at/dam/jcr:d6794f8f-d321-43df-9840-1a841f9bf5dc/BMF_UID_Konstruktionsregeln_Stand_November%202020.pdf">the BMF construction rules</see>.</summary>
internal sealed class FiVatScheme : PrefixedVatScheme {
    internal FiVatScheme() : base("FI", 8, 8, EntityCheckKind.Mod11, BmfSpecification) { }

    protected override EntityValidationResult ValidateNational(string number) {
        if (!EntityCharacters.IsDigits(number, 0, 8)) return Result(EntityValidationStatus.InvalidFormat);
        int sum = CheckDigits.WeightedSum(number, "7905842") + (number[2] - '0') * 10;
        return NationalCheck(number, (11 - sum % 11) % 11 == number[7] - '0');
    }
}
