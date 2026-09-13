namespace OfficeIMO.Project;

/// <summary>One compact timephased interval, retained without expanding it into daily records.</summary>
public sealed class ProjectTimephasedValue : ProjectObject {
    internal ProjectTimephasedValue(ProjectDocument document) : base(document) {
    }

    private int? _type;
    /// <summary>Stored type; null represents an absent source value.</summary>
    public int? Type { get => _type; set => Set(ref _type, value, true); }

    private int? _uid;
    /// <summary>Stored uid; null represents an absent source value.</summary>
    public int? Uid { get => _uid; set => Set(ref _uid, value); }

    private DateTime? _start;
    /// <summary>Stored local project date; null preserves an absent value.</summary>
    public DateTime? Start { get => _start; set => Set(ref _start, value, true); }

    private DateTime? _finish;
    /// <summary>Stored local project date; null preserves an absent value.</summary>
    public DateTime? Finish { get => _finish; set => Set(ref _finish, value, true); }

    private int? _unit;
    /// <summary>Stored unit; null represents an absent source value.</summary>
    public int? Unit { get => _unit; set => Set(ref _unit, value, true); }

    private string? _value;
    /// <summary>Stored lexical value interpreted by the containing field or interval type.</summary>
    public string? Value { get => _value; set => Set(ref _value, value, true); }
}
