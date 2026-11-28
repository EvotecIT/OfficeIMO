using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

/// <summary>Slovak VAT validation from <see href="https://www.bmf.gv.at/dam/jcr:d6794f8f-d321-43df-9840-1a841f9bf5dc/BMF_UID_Konstruktionsregeln_Stand_November%202020.pdf">the BMF construction rules</see>.</summary>
internal sealed class SkVatScheme : PrefixedVatScheme {
    internal SkVatScheme() : base("SK", 10, 10, EntityCheckKind.Mod11, BmfSpecification) { }

    protected override EntityValidationResult ValidateNational(string number) {
        if (!EntityCharacters.IsDigits(number, 0, 10) || number[0] == '0' || "234789".IndexOf(number[2]) < 0)
            return Result(EntityValidationStatus.InvalidFormat);
        return NationalCheck(number, (EntityCharacters.Number(number, 0, 9) % 11 * 10 + number[9] - '0') % 11 == 0);
    }
}
