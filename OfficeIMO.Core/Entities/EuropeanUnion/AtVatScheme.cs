using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

/// <summary>Austrian VAT validation from <see href="https://www.bmf.gv.at/dam/jcr:d6794f8f-d321-43df-9840-1a841f9bf5dc/BMF_UID_Konstruktionsregeln_Stand_November%202020.pdf">the BMF construction rules</see>.</summary>
internal sealed class AtVatScheme : PrefixedVatScheme {
    internal AtVatScheme() : base("AT", 9, 9, EntityCheckKind.Mod10, BmfSpecification) { }

    protected override EntityValidationResult ValidateNational(string number) {
        if (number[0] != 'U' || !EntityCharacters.IsDigits(number, 1, 8)) return Result(EntityValidationStatus.InvalidFormat);
        int sum = CheckDigits.WeightedSum(number, "1212121", 1, expandProducts: true) + 4;
        return NationalCheck(number, (sum + number[8] - '0') % 10 == 0);
    }
}
