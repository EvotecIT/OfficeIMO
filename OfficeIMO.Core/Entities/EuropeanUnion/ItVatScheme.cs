using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

/// <summary>Italian VAT validation from <see href="https://www.bmf.gv.at/dam/jcr:d6794f8f-d321-43df-9840-1a841f9bf5dc/BMF_UID_Konstruktionsregeln_Stand_November%202020.pdf">the BMF construction rules</see>.</summary>
internal sealed class ItVatScheme : PrefixedVatScheme {
    internal ItVatScheme() : base("IT", 11, 11, EntityCheckKind.Luhn, BmfSpecification) { }

    protected override EntityValidationResult ValidateNational(string number) {
        if (!EntityCharacters.IsDigits(number, 0, 11)) return Result(EntityValidationStatus.InvalidFormat);
        int office = EntityCharacters.Number(number, 7, 3);
        if (office is not (> 0 and <= 100 or 120 or 121 or 888 or 999))
            return NationalResult(EntityValidationStatus.NotAllocated, number);
        return NationalCheck(number, CheckDigits.Luhn(number));
    }
}
