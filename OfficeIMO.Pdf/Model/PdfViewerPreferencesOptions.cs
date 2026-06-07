namespace OfficeIMO.Pdf;

/// <summary>
/// Simple viewer preferences emitted in the generated PDF catalog.
/// </summary>
public sealed class PdfViewerPreferencesOptions {
    private PdfNonFullScreenPageMode? _nonFullScreenPageMode;
    private PdfViewerDirection? _direction;
    private PdfPrintScaling? _printScaling;
    private PdfDuplexMode? _duplex;
    private PdfPageBoundaryBox? _viewArea;
    private PdfPageBoundaryBox? _viewClip;
    private PdfPageBoundaryBox? _printArea;
    private PdfPageBoundaryBox? _printClip;
    private int? _numCopies;
    private List<PdfPrintPageRange>? _printPageRanges;

    /// <summary>Requests that viewers display the document title from metadata instead of the file name.</summary>
    public bool? DisplayDocTitle { get; set; }
    /// <summary>Requests that viewers hide the toolbar when the document is opened.</summary>
    public bool? HideToolbar { get; set; }
    /// <summary>Requests that viewers hide the menu bar when the document is opened.</summary>
    public bool? HideMenubar { get; set; }
    /// <summary>Requests that viewers hide user-interface elements when the document is opened.</summary>
    public bool? HideWindowUI { get; set; }
    /// <summary>Requests that viewers resize the window to fit the first displayed page.</summary>
    public bool? FitWindow { get; set; }
    /// <summary>Requests that viewers center the document window on screen.</summary>
    public bool? CenterWindow { get; set; }
    /// <summary>Requests that viewers pick the printer tray based on each page's PDF size.</summary>
    public bool? PickTrayByPdfSize { get; set; }
    /// <summary>Requests the page mode used when leaving full-screen display.</summary>
    public PdfNonFullScreenPageMode? NonFullScreenPageMode {
        get => _nonFullScreenPageMode;
        set {
            if (value.HasValue) {
                Guard.NonFullScreenPageMode(value.Value, nameof(NonFullScreenPageMode));
            }

            _nonFullScreenPageMode = value;
        }
    }
    /// <summary>Requests the viewer page progression direction.</summary>
    public PdfViewerDirection? Direction {
        get => _direction;
        set {
            if (value.HasValue) {
                Guard.ViewerDirection(value.Value, nameof(Direction));
            }

            _direction = value;
        }
    }
    /// <summary>Requests viewer print scaling behavior.</summary>
    public PdfPrintScaling? PrintScaling {
        get => _printScaling;
        set {
            if (value.HasValue) {
                Guard.PrintScaling(value.Value, nameof(PrintScaling));
            }

            _printScaling = value;
        }
    }
    /// <summary>Requests viewer duplex-printing behavior.</summary>
    public PdfDuplexMode? Duplex {
        get => _duplex;
        set {
            if (value.HasValue) {
                Guard.DuplexMode(value.Value, nameof(Duplex));
            }

            _duplex = value;
        }
    }
    /// <summary>Requests the page boundary box used when fitting pages in the viewer.</summary>
    public PdfPageBoundaryBox? ViewArea {
        get => _viewArea;
        set {
            if (value.HasValue) {
                Guard.PageBoundaryBox(value.Value, nameof(ViewArea));
            }

            _viewArea = value;
        }
    }
    /// <summary>Requests the page boundary box used when clipping pages in the viewer.</summary>
    public PdfPageBoundaryBox? ViewClip {
        get => _viewClip;
        set {
            if (value.HasValue) {
                Guard.PageBoundaryBox(value.Value, nameof(ViewClip));
            }

            _viewClip = value;
        }
    }
    /// <summary>Requests the page boundary box used when printing pages.</summary>
    public PdfPageBoundaryBox? PrintArea {
        get => _printArea;
        set {
            if (value.HasValue) {
                Guard.PageBoundaryBox(value.Value, nameof(PrintArea));
            }

            _printArea = value;
        }
    }
    /// <summary>Requests the page boundary box used when clipping printed pages.</summary>
    public PdfPageBoundaryBox? PrintClip {
        get => _printClip;
        set {
            if (value.HasValue) {
                Guard.PageBoundaryBox(value.Value, nameof(PrintClip));
            }

            _printClip = value;
        }
    }
    /// <summary>Requested number of copies used to initialize the viewer print dialog.</summary>
    public int? NumCopies {
        get => _numCopies;
        set {
            if (value.HasValue) {
                Guard.PositiveInteger(value.Value, nameof(NumCopies));
            }

            _numCopies = value;
        }
    }

    /// <summary>One-based inclusive page ranges used to initialize the viewer print dialog.</summary>
    public IReadOnlyList<PdfPrintPageRange> PrintPageRanges =>
        _printPageRanges == null || _printPageRanges.Count == 0
            ? Array.Empty<PdfPrintPageRange>()
            : _printPageRanges.Select(range => range.Clone()).ToList().AsReadOnly();

    internal bool HasAny =>
        DisplayDocTitle.HasValue ||
        HideToolbar.HasValue ||
        HideMenubar.HasValue ||
        HideWindowUI.HasValue ||
        FitWindow.HasValue ||
        CenterWindow.HasValue ||
        PickTrayByPdfSize.HasValue ||
        NonFullScreenPageMode.HasValue ||
        Direction.HasValue ||
        PrintScaling.HasValue ||
        Duplex.HasValue ||
        ViewArea.HasValue ||
        ViewClip.HasValue ||
        PrintArea.HasValue ||
        PrintClip.HasValue ||
        NumCopies.HasValue ||
        (_printPageRanges != null && _printPageRanges.Count > 0);

    /// <summary>Adds a one-based inclusive page range used to initialize the viewer print dialog.</summary>
    public PdfViewerPreferencesOptions AddPrintPageRange(int startPageNumber, int endPageNumber) {
        return AddPrintPageRange(new PdfPrintPageRange(startPageNumber, endPageNumber));
    }

    /// <summary>Adds a one-based inclusive page range used to initialize the viewer print dialog.</summary>
    public PdfViewerPreferencesOptions AddPrintPageRange(PdfPrintPageRange range) {
        Guard.NotNull(range, nameof(range));
        (_printPageRanges ??= new List<PdfPrintPageRange>()).Add(range.Clone());
        return this;
    }

    /// <summary>Clears generated print page ranges while leaving other viewer preferences unchanged.</summary>
    public PdfViewerPreferencesOptions ClearPrintPageRanges() {
        _printPageRanges?.Clear();
        return this;
    }

    internal PdfViewerPreferencesOptions Clone() {
        var clone = new PdfViewerPreferencesOptions {
            DisplayDocTitle = DisplayDocTitle,
            HideToolbar = HideToolbar,
            HideMenubar = HideMenubar,
            HideWindowUI = HideWindowUI,
            FitWindow = FitWindow,
            CenterWindow = CenterWindow,
            PickTrayByPdfSize = PickTrayByPdfSize,
            NonFullScreenPageMode = NonFullScreenPageMode,
            Direction = Direction,
            PrintScaling = PrintScaling,
            Duplex = Duplex,
            ViewArea = ViewArea,
            ViewClip = ViewClip,
            PrintArea = PrintArea,
            PrintClip = PrintClip,
            NumCopies = NumCopies
        };

        if (_printPageRanges != null) {
            foreach (PdfPrintPageRange range in _printPageRanges) {
                clone.AddPrintPageRange(range);
            }
        }

        return clone;
    }
}
