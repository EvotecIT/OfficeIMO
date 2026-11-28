using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

/// <summary>Croatian VAT validation from <see href="https://regos.hr/app/uploads/2018/07/KONTROLA-OIB-a.pdf">Croatia’s central register of insured persons</see>.</summary>
internal sealed class HrVatScheme : PrefixedVatScheme {
    internal HrVatScheme() : base("HR", 11, 11, EntityCheckKind.Other, "https://regos.hr/app/uploads/2018/07/KONTROLA-OIB-a.pdf") { }

    protected override EntityValidationResult ValidateNational(string number) =>
        NationalDigits(number, static value => CheckDigits.HybridMod11_10(value));
}
