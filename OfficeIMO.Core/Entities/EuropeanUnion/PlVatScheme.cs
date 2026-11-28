using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

/// <summary>Polish VAT validation from <see href="https://www.bmf.gv.at/dam/jcr:d6794f8f-d321-43df-9840-1a841f9bf5dc/BMF_UID_Konstruktionsregeln_Stand_November%202020.pdf">the BMF construction rules</see>.</summary>
internal sealed class PlVatScheme : PrefixedVatScheme {
    internal PlVatScheme() : base("PL", 10, 10, EntityCheckKind.Mod11, BmfSpecification) { }
    public override System.Collections.Generic.IReadOnlyList<string> Contains { get; } = System.Array.AsReadOnly(new[] { "PL.NIP" });

    internal override void RecognizeContained(EntityRecognitionContext context, ScoredEntity outer) {
        if (!context.IsEnabled(PolandSchemes.Nip)) return;
        if ((outer.Evidence & EntityEvidence.CountryPrefix) == 0) return;
        int start = outer.Start;
        int skip = 0;
        while (start < outer.End && skip < 2) {
            if (EntityCharacters.IsLetter(EntityCharacters.Upper(context.Text[start]))) skip++;
            start++;
        }
        while (start < outer.End && !EntityCharacters.IsDigit(context.Text[start])) start++;
        if (start >= outer.End) return;

        var token = new EntityToken(EntityTokenShape.DigitRun, start, outer.End - start, 10, false);
        EntityEvidence evidence = PolandRecognizers.MatchesNipGrouping(context.Text, token) ? EntityEvidence.IssuerFormatting : EntityEvidence.None;
        IdentifierReadings.Submit(context, PolandSchemes.Nip, start, outer.End, EntityCandidateStrength.DigitCheck, evidence);
    }

    protected override EntityValidationResult ValidateNational(string number) =>
        NationalDigits(number, static value => NipScheme.PassesCheck(value));
}
