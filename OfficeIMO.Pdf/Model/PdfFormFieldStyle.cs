namespace OfficeIMO.Pdf;

/// <summary>
/// Visual style for generated simple AcroForm fields.
/// </summary>
public class PdfFormFieldStyle {
    private double _borderWidth = 1D;
    private string? _alternateName;
    private string? _mappingName;
    private PdfFormFieldTextAlignment? _textAlignment;
    private int? _maxLength;
    private double[]? _borderDashPattern;
    private PdfFormFieldBorderStyle _borderStyle = PdfFormFieldBorderStyle.Solid;

    /// <summary>Background fill color. Set to null for transparent field appearance streams.</summary>
    public PdfColor? BackgroundColor { get; set; } = PdfColor.White;

    /// <summary>Border stroke color. Set to null for no border stroke.</summary>
    public PdfColor? BorderColor { get; set; } = new PdfColor(0.75, 0.75, 0.75);

    /// <summary>Text color for generated text and choice field appearance streams.</summary>
    public PdfColor TextColor { get; set; } = PdfColor.Black;

    /// <summary>Check mark or radio dot color for generated button field appearance streams.</summary>
    public PdfColor MarkColor { get; set; } = PdfColor.Black;

    /// <summary>When true, generated AcroForm fields emit the common read-only field flag.</summary>
    public bool IsReadOnly { get; set; }

    /// <summary>When true, generated AcroForm fields emit the common required field flag.</summary>
    public bool IsRequired { get; set; }

    /// <summary>When true, generated AcroForm fields emit the common no-export field flag.</summary>
    public bool IsNoExport { get; set; }

    /// <summary>When true, generated text fields emit the multiline field flag.</summary>
    public bool IsMultiline { get; set; }

    /// <summary>When true, generated text fields emit the password field flag.</summary>
    public bool IsPassword { get; set; }

    /// <summary>When true, generated text fields emit the file-select field flag.</summary>
    public bool IsFileSelect { get; set; }

    /// <summary>When true, generated text and choice fields emit the do-not-spell-check field flag.</summary>
    public bool DoesNotSpellCheck { get; set; }

    /// <summary>When true, generated text fields emit the do-not-scroll field flag.</summary>
    public bool DoesNotScroll { get; set; }

    /// <summary>When true, generated text fields emit the comb field flag. Requires <see cref="MaxLength"/>.</summary>
    public bool IsComb { get; set; }

    /// <summary>When true, generated combo choice fields emit the editable-choice field flag.</summary>
    public bool IsEditableChoice { get; set; }

    /// <summary>When true, generated choice fields emit the sort field flag.</summary>
    public bool IsSortedChoice { get; set; }

    /// <summary>When true, generated choice fields emit the commit-on-selection-change field flag.</summary>
    public bool CommitsOnSelectionChange { get; set; }

    /// <summary>Optional maximum text length emitted as /MaxLen for generated text fields.</summary>
    public int? MaxLength {
        get => _maxLength;
        set {
            if (value.HasValue && value.Value < 1) {
                throw new ArgumentOutOfRangeException(nameof(value), value.Value, "PDF text field maximum length must be a positive integer.");
            }

            _maxLength = value;
        }
    }

    /// <summary>Optional text alignment for generated text and choice fields. When null, document-level AcroForm defaults can apply.</summary>
    public PdfFormFieldTextAlignment? TextAlignment {
        get => _textAlignment;
        set {
            if (value.HasValue) {
                Guard.FormFieldTextAlignment(value.Value, nameof(TextAlignment));
            }

            _textAlignment = value;
        }
    }

    /// <summary>Border stroke width in points. Set to 0 to suppress border drawing.</summary>
    public double BorderWidth {
        get => _borderWidth;
        set {
            if (value < 0 || double.IsNaN(value) || double.IsInfinity(value)) {
                throw new ArgumentOutOfRangeException(nameof(value), value, "PDF form field border width must be a non-negative finite number.");
            }

            _borderWidth = value;
        }
    }

    /// <summary>Border rendering style for generated field dictionaries and appearance streams.</summary>
    public PdfFormFieldBorderStyle BorderStyle {
        get => _borderStyle;
        set {
            switch (value) {
                case PdfFormFieldBorderStyle.Solid:
                case PdfFormFieldBorderStyle.Dashed:
                case PdfFormFieldBorderStyle.Underline:
                case PdfFormFieldBorderStyle.Beveled:
                case PdfFormFieldBorderStyle.Inset:
                    _borderStyle = value;
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, "PDF form field border style must be Solid, Dashed, Underline, Beveled, or Inset.");
            }
        }
    }

    /// <summary>Optional border dash pattern emitted into generated appearance streams. Null or empty output means a solid border.</summary>
    public IReadOnlyList<double>? BorderDashPattern {
        get => _borderDashPattern;
        set {
            if (value == null) {
                _borderDashPattern = null;
                return;
            }

            if (value.Count == 0) {
                throw new ArgumentException("PDF form field border dash pattern must contain at least one value.", nameof(value));
            }

            var copy = new double[value.Count];
            bool hasPositiveSegment = false;
            for (int i = 0; i < value.Count; i++) {
                double segment = value[i];
                if (segment < 0 || double.IsNaN(segment) || double.IsInfinity(segment)) {
                    throw new ArgumentOutOfRangeException(nameof(value), segment, "PDF form field border dash pattern values must be non-negative finite numbers.");
                }

                if (segment > 0D) {
                    hasPositiveSegment = true;
                }

                copy[i] = segment;
            }

            if (!hasPositiveSegment) {
                throw new ArgumentException("PDF form field border dash pattern must contain at least one positive value.", nameof(value));
            }

            _borderDashPattern = copy;
        }
    }

    /// <summary>Alternate field name emitted as AcroForm /TU metadata for accessibility-oriented field descriptions.</summary>
    public string? AlternateName {
        get => _alternateName;
        set {
            ValidateOptionalText(value, nameof(AlternateName));
            _alternateName = value;
        }
    }

    /// <summary>Mapping name emitted as AcroForm /TM metadata for export and assistive-processing workflows.</summary>
    public string? MappingName {
        get => _mappingName;
        set {
            ValidateOptionalText(value, nameof(MappingName));
            _mappingName = value;
        }
    }

    /// <summary>Creates a copy of this form field style.</summary>
    public PdfFormFieldStyle Clone() {
        return new PdfFormFieldStyle {
            BackgroundColor = BackgroundColor,
            BorderColor = BorderColor,
            BorderWidth = BorderWidth,
            BorderStyle = BorderStyle,
            TextColor = TextColor,
            MarkColor = MarkColor,
            BorderDashPattern = _borderDashPattern == null ? null : (double[])_borderDashPattern.Clone(),
            IsReadOnly = IsReadOnly,
            IsRequired = IsRequired,
            IsNoExport = IsNoExport,
            IsMultiline = IsMultiline,
            IsPassword = IsPassword,
            IsFileSelect = IsFileSelect,
            DoesNotSpellCheck = DoesNotSpellCheck,
            DoesNotScroll = DoesNotScroll,
            IsComb = IsComb,
            IsEditableChoice = IsEditableChoice,
            IsSortedChoice = IsSortedChoice,
            CommitsOnSelectionChange = CommitsOnSelectionChange,
            MaxLength = MaxLength,
            TextAlignment = TextAlignment,
            AlternateName = AlternateName,
            MappingName = MappingName
        };
    }

    private static void ValidateOptionalText(string? value, string paramName) {
        if (value != null && string.IsNullOrWhiteSpace(value)) {
            throw new ArgumentException("PDF form field metadata must be null or non-empty text.", paramName);
        }
    }
}
