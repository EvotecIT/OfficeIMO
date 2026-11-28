using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

/// <summary>Belgian VAT validation from <see href="https://www.bmf.gv.at/dam/jcr:d6794f8f-d321-43df-9840-1a841f9bf5dc/BMF_UID_Konstruktionsregeln_Stand_November%202020.pdf">the BMF construction rules</see>.</summary>
internal sealed class BeVatScheme : PrefixedVatScheme {
    internal BeVatScheme() : base("BE", 10, 10, EntityCheckKind.Other, BmfSpecification) { }

    protected override EntityValidationResult ValidateNational(string number) {
        if (!EntityCharacters.IsDigits(number, 0, 10) || number[0] > '1') return Result(EntityValidationStatus.InvalidFormat);
        return NationalCheck(number, 97 - EntityCharacters.Number(number, 0, 8) % 97 == EntityCharacters.Number(number, 8, 2));
    }
}
