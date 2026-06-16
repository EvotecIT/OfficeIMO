namespace OfficeIMO.Rtf;

/// <summary>
/// Word form-field metadata stored in an RTF <c>\ffdata</c> destination.
/// </summary>
public sealed class RtfFormFieldData {
    private readonly List<RtfFormFieldDataControl> _controls = new List<RtfFormFieldDataControl>();
    private readonly List<string> _dropDownItems = new List<string>();

    /// <summary>Raw controls in source or creation order.</summary>
    public IReadOnlyList<RtfFormFieldDataControl> Controls => _controls.AsReadOnly();

    /// <summary>Drop-down list items from repeated <c>\ffl</c> destinations.</summary>
    public IReadOnlyList<string> DropDownItems => _dropDownItems.AsReadOnly();

    /// <summary>Raw <c>\fftype</c> value.</summary>
    public int? TypeCode { get; set; }

    /// <summary>Known form-field kind when <see cref="TypeCode"/> is recognized.</summary>
    public RtfFormFieldKind? Kind {
        get {
            if (!TypeCode.HasValue) return null;
            return Enum.IsDefined(typeof(RtfFormFieldKind), TypeCode.Value) ? (RtfFormFieldKind)TypeCode.Value : null;
        }
        set => TypeCode = value.HasValue ? (int)value.Value : null;
    }

    /// <summary>Form field bookmark/name from <c>\ffname</c>.</summary>
    public string? Name { get; set; }

    /// <summary>Default text from <c>\ffdeftext</c>.</summary>
    public string? DefaultText { get; set; }

    /// <summary>Text formatting string from <c>\ffformat</c>.</summary>
    public string? Format { get; set; }

    /// <summary>Custom help text from <c>\ffhelptext</c>.</summary>
    public string? HelpText { get; set; }

    /// <summary>Custom status-bar text from <c>\ffstattext</c>.</summary>
    public string? StatusText { get; set; }

    /// <summary>Entry macro name from <c>\ffentrymcr</c>.</summary>
    public string? EntryMacro { get; set; }

    /// <summary>Exit macro name from <c>\ffexitmcr</c>.</summary>
    public string? ExitMacro { get; set; }

    /// <summary>Whether the field is enabled, from <c>\ffenabled</c>.</summary>
    public bool? Enabled { get; set; }

    /// <summary>Whether custom help text is used, from <c>\ffownhelp</c>.</summary>
    public bool? OwnHelp { get; set; }

    /// <summary>Whether custom status text is used, from <c>\ffownstat</c>.</summary>
    public bool? OwnStatus { get; set; }

    /// <summary>Whether the field is protected, from <c>\ffprot</c>.</summary>
    public bool? Protected { get; set; }

    /// <summary>Whether field recalculation is requested, from <c>\ffrecalc</c>.</summary>
    public bool? RecalculateOnExit { get; set; }

    /// <summary>Maximum text length from <c>\ffmaxlen</c>.</summary>
    public int? MaxLength { get; set; }

    /// <summary>Check-box size in half-points from <c>\ffhps</c>.</summary>
    public int? CheckBoxSizeHalfPoints { get; set; }

    /// <summary>Default numeric result from <c>\ffdefres</c>.</summary>
    public int? DefaultResult { get; set; }

    /// <summary>Selected numeric result from <c>\ffres</c>.</summary>
    public int? Result { get; set; }

    /// <summary>Adds a raw <c>\ffdata</c> control in write order and updates known convenience properties.</summary>
    public RtfFormFieldDataControl AddControl(string name, int? parameter = null, bool hasParameter = true) {
        var control = new RtfFormFieldDataControl(name, parameter, hasParameter);
        _controls.Add(control);
        ApplyKnownControl(control);
        return control;
    }

    /// <summary>Adds a drop-down list item.</summary>
    public void AddDropDownItem(string item) {
        _dropDownItems.Add(item ?? string.Empty);
    }

    internal void AddParsedControl(RtfFormFieldDataControl control) {
        _controls.Add(control ?? throw new ArgumentNullException(nameof(control)));
        ApplyKnownControl(control);
    }

    private void ApplyKnownControl(RtfFormFieldDataControl control) {
        switch (control.Name) {
            case "fftype":
                TypeCode = control.Parameter;
                break;
            case "ffenabled":
                Enabled = ReadToggle(control);
                break;
            case "ffownhelp":
                OwnHelp = ReadToggle(control);
                break;
            case "ffownstat":
                OwnStatus = ReadToggle(control);
                break;
            case "ffprot":
                Protected = ReadToggle(control);
                break;
            case "ffrecalc":
                RecalculateOnExit = ReadToggle(control);
                break;
            case "ffmaxlen":
                MaxLength = control.Parameter;
                break;
            case "ffhps":
                CheckBoxSizeHalfPoints = control.Parameter;
                break;
            case "ffdefres":
                DefaultResult = control.Parameter;
                break;
            case "ffres":
                Result = control.Parameter;
                break;
        }
    }

    private static bool ReadToggle(RtfFormFieldDataControl control) => !control.HasParameter || control.Parameter != 0;
}
