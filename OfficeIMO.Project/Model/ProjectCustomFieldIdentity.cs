namespace OfficeIMO.Project;

/// <summary>Qualified scalar field identities shared by file codecs and custom-field calculation.</summary>
internal sealed class ProjectCustomFieldIdentity {
    internal readonly bool IsTask;
    internal readonly uint Id, Relative, DurationFormat;
    internal readonly string Kind;
    private readonly string _nameKind;
    internal readonly int Number;
    internal static string NormalizeId(string value) => uint.TryParse(value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out uint id)
        ? id.ToString(System.Globalization.CultureInfo.InvariantCulture) : value;
    internal static bool SameId(string? first, string? second) => first != null && second != null && NormalizeId(first) == NormalizeId(second);
    internal string Name => _nameKind + Number.ToString(System.Globalization.CultureInfo.InvariantCulture);
    private ProjectCustomFieldIdentity(bool task, uint relative, string kind, int number, string? nameKind = null) {
        IsTask = task;
        Relative = relative; Id = (task ? 0x0b400000u : 0x0c400000u) | relative; Kind = kind; Number = number; _nameKind = nameKind ?? kind;
        DurationFormat = task && kind == "Duration" ? 0x0b400000u | (number <= 3 ? 0xb7u + (uint)number - 1 : 0x151u + (uint)number - 4) : 0;
    }
    internal static readonly IReadOnlyList<ProjectCustomFieldIdentity> TaskFields = Build(true), ResourceFields = Build(false);
    private static IReadOnlyList<ProjectCustomFieldIdentity> Build(bool task) {
        uint[] taskText = { 0x33, 0x36, 0x39, 0x3c, 0x3f, 0x42, 0x43, 0x44, 0x45, 0x46 };
        uint[] resourceText = { 8, 9, 0x1e, 0x1f, 0x20, 0x61, 0x62, 0x63, 0x64, 0x65 };
        var fields = new List<ProjectCustomFieldIdentity>();
        void Field(uint id, string kind, int number) => fields.Add(new ProjectCustomFieldIdentity(task, id, kind, number));
        for (int number = 1; number <= 30; number++) Field(number <= 10 ? (task ? taskText : resourceText)[number - 1] : (task ? 0x13du : 0xe1u) + (uint)number - 11, "Text", number);
        for (int number = 1; number <= 20; number++) {
            Field(number <= 5 ? (task ? 0x57u : 0x70u) + (uint)number - 1 : (task ? 0x12eu : 0xcdu) + (uint)number - 6, "Number", number);
            uint flag = number <= 10 ? task ? 0x48u + (uint)number - 1 : number == 10 ? 0x7eu : 0x7fu + (uint)number - 1 : (task ? 0x124u : 0xc3u) + (uint)number - 11;
            Field(flag, "Flag", number);
        }
        for (int number = 1; number <= 10; number++) {
            Field((task ? 0x109u : 0xadu) + (uint)number - 1, "Date", number);
            Field(number <= 3 ? (task ? 0x6au : 0x7bu) + (uint)number - 1 : (task ? 0x102u : 0xa6u) + (uint)number - 4, "Cost", number);
            Field(number <= 3 ? (task ? 0x67u : 0x75u) + (uint)number - 1 : (task ? 0x113u : 0xb7u) + (uint)number - 4, "Duration", number);
        }
        if (task) for (int number = 1; number <= 5; number++) {
            fields.Add(new ProjectCustomFieldIdentity(true, 0x34u + (uint)(number - 1) * 3, "Date", number, "Start"));
            fields.Add(new ProjectCustomFieldIdentity(true, 0x35u + (uint)(number - 1) * 3, "Date", number, "Finish"));
        }
        return fields;
    }
}
