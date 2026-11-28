using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

/// <summary>Hungarian VAT validation from <see href="https://www.bmf.gv.at/dam/jcr:d6794f8f-d321-43df-9840-1a841f9bf5dc/BMF_UID_Konstruktionsregeln_Stand_November%202020.pdf">the BMF construction rules</see>.</summary>
internal sealed class HuVatScheme : PrefixedVatScheme {
    internal HuVatScheme() : base("HU", 8, 8, EntityCheckKind.Mod10, BmfSpecification) { }

    protected override EntityValidationResult ValidateNational(string number) =>
        NationalDigits(number, static value => CheckDigits.WeightedSum(value, "97319731") % 10 == 0);
}
