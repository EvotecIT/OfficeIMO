using System;
using System.Collections.Generic;
using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

/// <summary>Polish REGON format from <see href="https://api.sejm.gov.pl/eli/acts/DU/2015/2009/text.html">the REGON regulation</see>
/// and checksum weights from <see href="https://numerik.slashlab.pl/guide/algorithms/">Numerik's algorithm documentation</see>.</summary>
internal sealed class RegonScheme : CompactEntityScheme {
    public override string Id => "PL.REGON";
    public override EntityKind Kind => EntityKind.CompanyId;
    public override string Country => "PL";
    public override string DisplayName => "REGON";
    public override string Issuer => "Statistics Poland";
    public override Uri Specification { get; } = new("https://numerik.slashlab.pl/guide/algorithms/");
    public override int MinLength => 9;
    public override int MaxLength => 14;
    public override EntityCheckKind Check => EntityCheckKind.Mod11;
    public override EntitySensitivity Sensitivity => EntitySensitivity.Organizational;
    public override IReadOnlyList<string> Contains { get; } = Array.AsReadOnly(new[] { "PL.REGON" });

    protected override IReadOnlyList<EntityCheckInfo> DescribeChecks() => Array.AsReadOnly(new[] {
        new EntityCheckInfo("REGON 9 checksum", EntityCheckKind.Mod11, EntityCheckProvenance.CommunityDocumented, Specification),
        new EntityCheckInfo("REGON 14 checksum", EntityCheckKind.Mod11, EntityCheckProvenance.CommunityDocumented, Specification),
    });

    protected override EntityValidationResult ValidateCompact(string value) {
        if (value.Length != 9 && value.Length != 14) return Result(EntityValidationStatus.InvalidLength);
        if (!EntityCharacters.IsDigits(value, 0, value.Length)) return Result(EntityValidationStatus.InvalidCharacters);
        bool validParent = PassesCheck(value, "89234567");
        if (!validParent || value.Length == 9) return Checked(value, validParent, Checks[0]);
        return Checked(value, PassesCheck(value, "2485097361248"), Checks[1]);
    }

    internal override void RecognizeContained(EntityRecognitionContext context, ScoredEntity outer) {
        string normalized = outer.Candidate.Validation.Normalized!;
        if (normalized.Length != 14 || IdentifierReadings.IsRepeated(normalized, 0, 9)) return;
        int end = outer.Start;
        int count = 0;
        while (end < outer.End && count < 9) {
            if (EntityCharacters.IsDigit(context.Text[end])) count++;
            end++;
        }
        if (count == 9)
            IdentifierReadings.Submit(context, outer.Scheme, outer.Start, end, EntityCandidateStrength.ShapeOnly);
    }

    private static bool PassesCheck(string value, string weights) =>
        CheckDigits.WeightedSum(value, weights) % 11 % 10 == value[weights.Length] - '0';

    public override EntityDetails? Decode(string normalized) {
        string? value = Normalize(normalized);
        return value?.Length == 14
            ? new EntityDetails(country: Country, variant: "REGON14", properties: new Dictionary<string, string> { ["parent"] = value.Substring(0, 9) })
            : null;
    }
}
