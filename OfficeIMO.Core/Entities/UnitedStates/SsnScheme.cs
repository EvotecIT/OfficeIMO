using System;
using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

/// <summary>US Social Security number format and exclusions from the <see href="https://www.ssa.gov/employer/randomizationfaqs.html">Social Security Administration</see>.</summary>
internal sealed class SsnScheme : CompactEntityScheme {
    public override string Id => "US.SSN";
    public override EntityKind Kind => EntityKind.NationalId;
    public override string Country => "US";
    public override string DisplayName => "Social Security number";
    public override string Issuer => "Social Security Administration";
    public override Uri Specification { get; } = new("https://www.ssa.gov/employer/randomizationfaqs.html");
    public override int MinLength => 9;
    public override int MaxLength => 9;

    protected override EntityValidationResult ValidateCompact(string value) {
        if (!EntityCharacters.IsDigits(value, 0, value.Length)) return Result(EntityValidationStatus.InvalidCharacters);
        int area = EntityCharacters.Number(value, 0, 3);
        bool valid = area != 0 && area != 666 && area < 900 &&
                     EntityCharacters.Number(value, 3, 2) != 0 && EntityCharacters.Number(value, 5, 4) != 0;
        return Result(valid ? EntityValidationStatus.Valid : EntityValidationStatus.NotAllocated, value);
    }

    public override string Format(string normalized, EntityFormatStyle style = EntityFormatStyle.Display) {
        string value = base.Format(normalized, style);
        return style == EntityFormatStyle.Display && value.Length == 9 ? value.Insert(5, "-").Insert(3, "-") : value;
    }
}
