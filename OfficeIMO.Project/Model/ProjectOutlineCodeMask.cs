namespace OfficeIMO.Project;

/// <summary>Formatting rule for one level of a hierarchical outline code.</summary>
public sealed class ProjectOutlineCodeMask : ProjectObject {
    internal ProjectOutlineCodeMask(ProjectDocument document) : base(document) {
    }
    private int? _level;
    /// <summary>One-based mask level.</summary>
    public int? Level { get => _level; set => Set(ref _level, value); }
    private int? _type;
    /// <summary>Mask character class: 0 numbers, 1 uppercase, 2 lowercase, 3 arbitrary characters.</summary>
    public int? Type { get => _type; set => Set(ref _type, value); }
    private int? _length;
    /// <summary>Required component length; zero permits any positive length.</summary>
    public int? Length { get => _length; set => Set(ref _length, value); }
    private string? _separator;
    /// <summary>Separator following this level when another level follows.</summary>
    public string? Separator { get => _separator; set => Set(ref _separator, value); }
}
