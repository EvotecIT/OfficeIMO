using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

/// <summary>Irish VAT validation from <see href="https://www.bmf.gv.at/dam/jcr:d6794f8f-d321-43df-9840-1a841f9bf5dc/BMF_UID_Konstruktionsregeln_Stand_November%202020.pdf">the BMF construction rules</see>.</summary>
internal sealed class IeVatScheme : PrefixedVatScheme {
    internal IeVatScheme() : base("IE", 8, 9, EntityCheckKind.Mod23, BmfSpecification) { }

    internal override bool AllowsVatPunctuation => true;

    protected override EntityValidationResult ValidateNational(string number) {
        const string checkLetters = "WABCDEFGHIJKLMNOPQRSTUV";
        if (checkLetters.IndexOf(number[7]) < 0) return Result(EntityValidationStatus.InvalidFormat);

        int sum;
        if (number.Length == 8 && (EntityCharacters.IsLetter(number[1]) || IsAdditionalCharacter(number[1]))) {
            if (!EntityCharacters.IsDigit(number[0]) || !EntityCharacters.IsDigits(number, 2, 5)) return Result(EntityValidationStatus.InvalidFormat);
            sum = CheckDigits.WeightedSum(number, "76543", 2) + 2 * (number[0] - '0');
        } else {
            if (!EntityCharacters.IsDigits(number, 0, 7)) return Result(EntityValidationStatus.InvalidFormat);
            sum = CheckDigits.WeightedSum(number, "8765432");
            if (number.Length == 9) {
                if (number[8] < 'A' || number[8] > 'I') return Result(EntityValidationStatus.InvalidFormat);
                sum += (number[8] - 'A' + 1) * 9;
            }
        }
        return NationalCheck(number, checkLetters[sum % 23] == number[7]);
    }
}
