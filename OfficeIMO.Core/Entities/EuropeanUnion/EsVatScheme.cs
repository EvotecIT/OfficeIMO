using System;
using System.Collections.Generic;
using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

/// <summary>Spanish VAT validation from <see href="https://josegoval.github.io/nif-dni-nie-cif-validation/es/guides/control-character/">the documented NIF and CIF rules</see>.</summary>
/// <remarks>DNI and NIE letter checks are published by the <see href="https://www.interior.gob.es/opencms/es/servicios-al-ciudadano/tramites-y-gestiones/dni/calculo-del-digito-de-control-del-nif-nie/">Spanish Ministry of the Interior</see>.
/// The <see href="https://www.caib.es/sites/civitasmanualsisuport/f/41519">Spanish Tax Agency technical note</see> defines the seven-digit checksum input for K, L and M.</remarks>
internal sealed class EsVatScheme : PrefixedVatScheme {
    internal EsVatScheme() : base("ES", 9, 9, EntityCheckKind.Other, "https://josegoval.github.io/nif-dni-nie-cif-validation/es/guides/control-character/") { }

    protected override IReadOnlyList<EntityCheckInfo> DescribeChecks() => Array.AsReadOnly(new[] {
        new EntityCheckInfo("DNI and NIE control letter", EntityCheckKind.Mod23, EntityCheckProvenance.IssuerPublished,
            new Uri("https://www.interior.gob.es/opencms/es/servicios-al-ciudadano/tramites-y-gestiones/dni/calculo-del-digito-de-control-del-nif-nie/")),
        new EntityCheckInfo("K, L and M NIF control letter", EntityCheckKind.Mod23, EntityCheckProvenance.IssuerPublished,
            new Uri("https://www.caib.es/sites/civitasmanualsisuport/f/41519")),
        new EntityCheckInfo("CIF control character", EntityCheckKind.Mod10, EntityCheckProvenance.CommunityDocumented, Specification),
    });

    protected override EntityValidationResult ValidateNational(string number) {
        const string personalCheckLetters = "TRWAGMYFPDXBNJZSQVHLCKE";
        if (EntityCharacters.IsDigits(number, 0, 8) && EntityCharacters.IsLetter(number[8]))
            return NationalCheck(number, personalCheckLetters[EntityCharacters.Number(number, 0, 8) % 23] == number[8]);

        if (!EntityCharacters.IsDigits(number, 1, 7)) return Result(EntityValidationStatus.InvalidFormat);
        char prefix = number[0];
        if (prefix is 'X' or 'Y' or 'Z') {
            if (!EntityCharacters.IsLetter(number[8])) return Result(EntityValidationStatus.InvalidFormat);
            int position = ((prefix - 'X') * 10000000 + EntityCharacters.Number(number, 1, 7)) % 23;
            return NationalCheck(number, personalCheckLetters[position] == number[8]);
        }

        if (prefix is 'K' or 'L' or 'M') {
            if (!EntityCharacters.IsLetter(number[8])) return Result(EntityValidationStatus.InvalidFormat);
            int position = EntityCharacters.Number(number, 1, 7) % 23;
            return NationalCheck(number, personalCheckLetters[position] == number[8], Checks[1]);
        }

        bool legalEntityDigit = "ABCDEFGHJUV".IndexOf(prefix) >= 0;
        bool legalEntityLetter = "NPQRSW".IndexOf(prefix) >= 0;
        if (!legalEntityDigit && !legalEntityLetter) return Result(EntityValidationStatus.InvalidFormat);

        int digit = (10 - CheckDigits.WeightedSum(number, "2121212", 1, expandProducts: true) % 10) % 10;
        char expected = legalEntityDigit ? (char)('0' + digit) : "JABCDEFGHI"[digit];
        return NationalCheck(number, number[8] == expected, Checks[2]);
    }
}
