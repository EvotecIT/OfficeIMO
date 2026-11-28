using System;
using System.Text;

namespace OfficeIMO.Entities.Internal;

/// <summary>Parses ASCII identifier spellings before delegating structure and checksum validation.</summary>
internal abstract class CompactEntityScheme : EntityScheme {
    protected virtual bool IsSeparator(char value) => EntityCharacters.IsSeparator(value);
    protected virtual bool IsAdditionalCharacter(char value) => false;
    protected virtual string Canonicalize(string compact) => compact;
    protected virtual int MaxWrittenCharacters => MaxLength * 2 + 4;

    public sealed override EntityValidationResult Validate(string? value) {
        if (string.IsNullOrWhiteSpace(value)) return Result(EntityValidationStatus.Empty);
        string written = value!.Trim();
        if (written.Length > MaxWrittenCharacters) return Result(EntityValidationStatus.InvalidLength);
        var compact = new StringBuilder(Math.Min(written.Length, MaxLength + 2));
        bool previousSeparator = true;
        foreach (char character in written) {
            if (IsSeparator(character)) {
                if (previousSeparator) return Result(EntityValidationStatus.InvalidFormat);
                previousSeparator = true;
                continue;
            }
            char upper = EntityCharacters.Upper(character);
            if (!EntityCharacters.IsLetter(upper) && !EntityCharacters.IsDigit(upper) && !IsAdditionalCharacter(upper)) return Result(EntityValidationStatus.InvalidCharacters);
            compact.Append(upper);
            previousSeparator = false;
        }
        if (previousSeparator) return Result(EntityValidationStatus.InvalidFormat);
        string normalized = Canonicalize(compact.ToString());
        if (normalized.Length < MinLength || normalized.Length > MaxLength) return Result(EntityValidationStatus.InvalidLength);
        return ValidateCompact(normalized);
    }

    protected abstract EntityValidationResult ValidateCompact(string value);

    protected EntityValidationResult Result(EntityValidationStatus status, string? normalized = null) => new(this, status, normalized);

    protected EntityValidationResult Checked(string value, bool valid, EntityCheckInfo? check = null) =>
        new(this, valid ? EntityValidationStatus.Valid : EntityValidationStatus.InvalidCheckDigit, value,
            valid ? EntityCheckOutcome.Passed : EntityCheckOutcome.Failed, check ?? Checks[0]);
}
