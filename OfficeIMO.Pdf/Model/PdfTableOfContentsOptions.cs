namespace OfficeIMO.Pdf;

/// <summary>Configures generated section navigation.</summary>
public sealed class PdfTableOfContentsOptions {
    private int _minimumLevel = 1;
    private int _maximumLevel = 3;
    private double _indentPerLevel = 14D;

    /// <summary>Optional TOC heading; set null to omit it.</summary>
    public string? Title { get; set; } = "Contents";
    /// <summary>Lowest section level included.</summary>
    public int MinimumLevel {
        get => _minimumLevel;
        set {
            ValidateLevel(value, nameof(value));
            _minimumLevel = value;
        }
    }
    /// <summary>Highest section level included.</summary>
    public int MaximumLevel {
        get => _maximumLevel;
        set {
            ValidateLevel(value, nameof(value));
            _maximumLevel = value;
        }
    }
    /// <summary>Indent added for each hierarchy level after the minimum.</summary>
    public double IndentPerLevel {
        get => _indentPerLevel;
        set {
            Guard.NonNegative(value, nameof(value));
            _indentPerLevel = value;
        }
    }
    /// <summary>Leader rendered between the linked title and page number.</summary>
    public PdfTabLeaderStyle Leader { get; set; } = PdfTabLeaderStyle.Dots;
    /// <summary>Optional formatter for physical output page numbers.</summary>
    public Func<int, string>? PageNumberFormatter { get; set; }

    internal PdfTableOfContentsOptions Clone() {
        if (MinimumLevel > MaximumLevel) throw new ArgumentException("TOC minimum level cannot exceed maximum level.");
        return new PdfTableOfContentsOptions {
            Title = Title,
            MinimumLevel = MinimumLevel,
            MaximumLevel = MaximumLevel,
            IndentPerLevel = IndentPerLevel,
            Leader = Leader,
            PageNumberFormatter = PageNumberFormatter
        };
    }

    private static void ValidateLevel(int value, string paramName) {
        if (value < 1 || value > 9) throw new ArgumentOutOfRangeException(paramName, value, "TOC level must be between 1 and 9.");
    }
}
