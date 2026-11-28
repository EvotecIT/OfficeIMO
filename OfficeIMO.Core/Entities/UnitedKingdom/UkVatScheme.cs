using System;
using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

/// <summary>UK and Northern Ireland VAT formats and checks from the <see href="https://www.bmf.gv.at/dam/jcr:d6794f8f-d321-43df-9840-1a841f9bf5dc/BMF_UID_Konstruktionsregeln_Stand_November%202020.pdf">GB and XI construction rules</see>.</summary>
internal sealed class UkVatScheme : PrefixedVatScheme {
    internal UkVatScheme(string prefix, string id) : base(prefix, 5, 12, EntityCheckKind.Other,
        BmfSpecification, "GB", id) { }

    public override string Issuer => "HM Revenue and Customs";

    protected override System.Collections.Generic.IReadOnlyList<EntityCheckInfo> DescribeChecks() =>
        Array.AsReadOnly(new[] {
            new EntityCheckInfo("UK VAT modulus 97", EntityCheckKind.Other, EntityCheckProvenance.IssuerPublished, Specification),
            new EntityCheckInfo("UK VAT adjusted modulus 97", EntityCheckKind.Other, EntityCheckProvenance.IssuerPublished, Specification),
        });

    protected override EntityValidationResult ValidateNational(string number) {
        if (number.Length == 5) {
            if (!EntityCharacters.IsDigits(number, 2, 3)) return Result(EntityValidationStatus.InvalidFormat);
            int serial = EntityCharacters.Number(number, 2, 3);
            if (number.StartsWith("GD", StringComparison.Ordinal))
                return NationalResult(serial <= 499 ? EntityValidationStatus.Valid : EntityValidationStatus.NotAllocated, number);
            if (number.StartsWith("HA", StringComparison.Ordinal))
                return NationalResult(serial >= 500 ? EntityValidationStatus.Valid : EntityValidationStatus.NotAllocated, number);
            return Result(EntityValidationStatus.InvalidFormat);
        }
        if (number.Length != 9 && number.Length != 12) return Result(EntityValidationStatus.InvalidLength);
        if (!EntityCharacters.IsDigits(number, 0, number.Length)) return Result(EntityValidationStatus.InvalidFormat);

        int firstSeven = EntityCharacters.Number(number, 0, 7);
        int check = EntityCharacters.Number(number, 7, 2);
        if ((firstSeven == 0 && check == 0) || (number.Length == 12 && EntityCharacters.Number(number, 9, 3) == 0))
            return NationalResult(EntityValidationStatus.NotAllocated, number);

        int sum = CheckDigits.WeightedSum(number, "8765432") + check;
        bool original = sum % 97 == 0;
        bool extended = (sum + 55) % 97 == 0;
        if (!original && !extended) return NationalCheck(number, false);

        bool excluded = original && (firstSeven is >= 100000 and <= 999999 or >= 9490001 and <= 9700000 or >= 9990001) ||
                        extended && firstSeven is >= 1 and <= 1000000;
        EntityCheckInfo applied = Checks[extended ? 1 : 0];
        return excluded
            ? new EntityValidationResult(this, EntityValidationStatus.NotAllocated, Prefix + number, EntityCheckOutcome.Passed, applied)
            : NationalCheck(number, true, applied);
    }
}
