using System;
using System.Collections.Generic;
using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

/// <summary>Dutch VAT validation from <see href="https://www.bmf.gv.at/dam/jcr:d6794f8f-d321-43df-9840-1a841f9bf5dc/BMF_UID_Konstruktionsregeln_Stand_November%202020.pdf">the BMF construction rules</see>.</summary>
internal sealed class NlVatScheme : PrefixedVatScheme {
    internal NlVatScheme() : base("NL", 12, 12, EntityCheckKind.Other, BmfSpecification) { }

    protected override IReadOnlyList<EntityCheckInfo> DescribeChecks() => Array.AsReadOnly(new[] {
        new EntityCheckInfo("Legacy VAT modulus 11", EntityCheckKind.Mod11, EntityCheckProvenance.IssuerPublished, Specification),
        new EntityCheckInfo("VAT modulus 97", EntityCheckKind.Mod97_10, EntityCheckProvenance.IssuerPublished, Specification),
    });

    protected override bool IsAdditionalCharacter(char value) => value is '+' or '*';

    protected override EntityValidationResult ValidateNational(string number) {
        if (!EntityCharacters.IsDigits(number, 10, 2)) return Result(EntityValidationStatus.InvalidFormat);

        int suffix = EntityCharacters.Number(number, 10, 2);
        bool legacy = number[9] == 'B' && EntityCharacters.IsDigits(number, 0, 9);
        if (legacy && suffix > 0 && CheckDigits.WeightedSum(number, "98765432") % 11 == number[8] - '0')
            return NationalCheck(number, true);

        if (suffix is < 2 or > 98) return NationalCheck(number, false);
        int remainder = CheckDigits.Mod97(Prefix);
        foreach (char character in number) {
            remainder = character switch {
                '+' => (remainder * 100 + 36) % 97,
                '*' => (remainder * 100 + 37) % 97,
                _ => CheckDigits.AppendMod97(remainder, character),
            };
        }
        return NationalCheck(number, remainder == 1, Checks[1]);
    }
}
