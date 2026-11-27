using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

/// <summary>Luxembourg VAT validation from <see href="https://www.bmf.gv.at/dam/jcr:d6794f8f-d321-43df-9840-1a841f9bf5dc/BMF_UID_Konstruktionsregeln_Stand_November%202020.pdf">the BMF construction rules</see>.</summary>
internal sealed class LuVatScheme : PrefixedVatScheme {
    internal LuVatScheme() : base("LU", 8, 8, EntityCheckKind.Other, BmfSpecification) { }

    protected override EntityValidationResult ValidateNational(string number) =>
        NationalDigits(number, static value => EntityCharacters.Number(value, 0, 6) % 89 == EntityCharacters.Number(value, 6, 2));
}
