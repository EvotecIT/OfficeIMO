using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

/// <summary>Swedish VAT validation from <see href="https://www.bmf.gv.at/dam/jcr:d6794f8f-d321-43df-9840-1a841f9bf5dc/BMF_UID_Konstruktionsregeln_Stand_November%202020.pdf">the BMF construction rules</see>.</summary>
internal sealed class SeVatScheme : PrefixedVatScheme {
    internal SeVatScheme() : base("SE", 12, 12, EntityCheckKind.Luhn, BmfSpecification) { }

    protected override EntityValidationResult ValidateNational(string number) {
        if (!EntityCharacters.IsDigits(number, 0, 12)) return Result(EntityValidationStatus.InvalidFormat);
        if (EntityCharacters.Number(number, 10, 2) is < 1 or > 94) return NationalResult(EntityValidationStatus.NotAllocated, number);
        return NationalCheck(number, CheckDigits.Luhn(number, length: 10));
    }
}
