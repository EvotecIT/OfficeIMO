using Avalonia;
using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Features.Workspace;

namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class MainWindowViewModel {
    private PdfEditorGesture? _pendingRedaction;
    private PdfRedactionPlan? _pendingRedactionPlan;
    private PdfWorkspace? _pendingRedactionWorkspace;
    private long _pendingRedactionRevision;
    private long _redactionPlanGeneration;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(ActiveEditorTool))]
    [NotifyPropertyChangedFor(nameof(EditorInstruction))]
    private PdfEditorToolChoice _selectedEditorToolChoice = null!;

    [ObservableProperty]
    private string _editorText = "Review note";

    [ObservableProperty]
    private string _editorAuthor = Environment.UserName;

    [ObservableProperty]
    private string _editorColorHex = "#E5484D";

    [ObservableProperty]
    private string _editorStampName = "Approved";

    [ObservableProperty]
    private string _editorLinkUri = "https://";

    [ObservableProperty]
    private double _editorFontSize = 14D;

    [ObservableProperty]
    private string _redactionRemovedMarker = string.Empty;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasPendingRedaction))]
    private string? _pendingRedactionSummary;

    [ObservableProperty]
    private string _watermarkText = "CONFIDENTIAL";

    [ObservableProperty]
    private PdfFormFieldViewModel? _selectedFormField;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsCreatingTextField))]
    [NotifyPropertyChangedFor(nameof(IsCreatingCheckBox))]
    [NotifyPropertyChangedFor(nameof(IsCreatingChoiceField))]
    [NotifyPropertyChangedFor(nameof(IsCreatingChoiceList))]
    [NotifyPropertyChangedFor(nameof(IsCreatingButton))]
    private PdfFormFieldCreationChoice _selectedFormFieldCreationChoice = null!;

    [ObservableProperty]
    private string _newFormFieldName = "Field1";

    [ObservableProperty]
    private int _newFormFieldPageNumber = 1;

    [ObservableProperty]
    private double _newFormFieldX = 36D;

    [ObservableProperty]
    private double _newFormFieldY = 36D;

    [ObservableProperty]
    private double _newFormFieldWidth = 180D;

    [ObservableProperty]
    private double _newFormFieldHeight = 28D;

    [ObservableProperty]
    private string _newFormFieldValue = string.Empty;

    [ObservableProperty]
    private string _newFormFieldOptions = "Option 1\nOption 2";

    [ObservableProperty]
    private string _newFormFieldCaption = "Button";

    [ObservableProperty]
    private bool _newFormFieldIsMultiline;

    [ObservableProperty]
    private bool _newFormFieldIsPassword;

    [ObservableProperty]
    private bool _newFormFieldIsChecked;

    [ObservableProperty]
    private bool _newFormFieldIsComboBox = true;

    [ObservableProperty]
    private bool _newFormFieldAllowsCustomValue;

    [ObservableProperty]
    private bool _newFormFieldAllowsMultipleSelection;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasSelectedObject))]
    [NotifyPropertyChangedFor(nameof(HasSelectedAnnotation))]
    [NotifyPropertyChangedFor(nameof(HasSelectedImage))]
    [NotifyPropertyChangedFor(nameof(HasSelectedText))]
    [NotifyPropertyChangedFor(nameof(SelectedAnnotationObjectNumber))]
    [NotifyPropertyChangedFor(nameof(CanReplaceSelectedText))]
    [NotifyPropertyChangedFor(nameof(CanReplaceSelectedImage))]
    [NotifyPropertyChangedFor(nameof(CanResizeSelectedAnnotation))]
    private PdfEditorSelection? _selectedObject;

    [ObservableProperty]
    private string? _selectedObjectSummary;

    [ObservableProperty]
    private string? _selectedAnnotationSummary;

    [ObservableProperty]
    private string _selectedAnnotationContents = string.Empty;

    [ObservableProperty]
    private string _selectedAnnotationAuthor = string.Empty;

    [ObservableProperty]
    private string _annotationReplyText = string.Empty;

    public ObservableCollection<PdfEditorToolChoice> EditorTools { get; } = new();

    public ObservableCollection<PdfFormFieldViewModel> FormFields { get; } = new();

    public ObservableCollection<PdfFormFieldCreationChoice> FormFieldCreationChoices { get; } = new();

    public PdfEditorTool ActiveEditorTool => SelectedEditorToolChoice.Tool;

    public string EditorInstruction => SelectedEditorToolChoice.Hint;

    public bool HasPendingRedaction => !string.IsNullOrWhiteSpace(PendingRedactionSummary);

    public bool HasFormFields => FormFields.Count > 0;

    public bool HasSelectedObject => SelectedObject is not null;

    public bool HasSelectedAnnotation => SelectedObject?.Kind == PdfEditorSelectionKind.Annotation;

    public bool HasSelectedImage => SelectedObject?.Kind == PdfEditorSelectionKind.Image;

    public bool HasSelectedText => SelectedObject?.Kind == PdfEditorSelectionKind.Text;

    public int? SelectedAnnotationObjectNumber => HasSelectedAnnotation ? SelectedObject?.ObjectNumber : null;

    public bool CanEditAnnotations => _workspace?.CanEditAnnotations == true;

    public bool CanEditPageContent => _workspace?.CanEditPageContent == true;

    public bool CanRedact => _workspace?.CanRedact == true;

    public bool CanFillForms => _workspace?.CanFillForms == true && SelectedFormField?.CanFill == true;

    public bool CanFlattenForms => _workspace?.CanFlattenForms == true;

    public bool CanFillAndFlattenForms => CanFillForms && CanFlattenForms;

    public bool CanFlattenSelectedFormField => CanFlattenForms && SelectedFormField is not null;

    public bool CanAuthorForms => _workspace?.CanAuthorForms == true;

    public bool IsCreatingTextField => SelectedFormFieldCreationChoice.Kind == PdfFormFieldCreationKind.Text;

    public bool IsCreatingCheckBox => SelectedFormFieldCreationChoice.Kind == PdfFormFieldCreationKind.CheckBox;

    public bool IsCreatingChoiceField => SelectedFormFieldCreationChoice.Kind is PdfFormFieldCreationKind.Choice or PdfFormFieldCreationKind.RadioButtonGroup;

    public bool IsCreatingChoiceList => SelectedFormFieldCreationChoice.Kind == PdfFormFieldCreationKind.Choice;

    public bool IsCreatingButton => SelectedFormFieldCreationChoice.Kind == PdfFormFieldCreationKind.PushButton;

    partial void OnSelectedEditorToolChoiceChanged(PdfEditorToolChoice value) {
        foreach (PdfPageViewModel page in Pages) page.EditorTool = value.Tool;
        if (value.Tool != PdfEditorTool.Redact) CancelPendingRedaction();
    }

    partial void OnSelectedFormFieldChanged(PdfFormFieldViewModel? value) {
        OnPropertyChanged(nameof(CanFillForms));
        OnPropertyChanged(nameof(CanFillAndFlattenForms));
        OnPropertyChanged(nameof(CanFlattenSelectedFormField));
    }

    [RelayCommand]
    private void SelectEditorTool(string? toolId) {
        if (!Enum.TryParse(toolId, ignoreCase: true, out PdfEditorTool tool)) return;
        PdfEditorToolChoice? choice = EditorTools.FirstOrDefault(candidate => candidate.Tool == tool);
        if (choice is not null) SelectedEditorToolChoice = choice;
    }

    private void InitializeLocalizedEditorChoices() {
        EditorTools.Add(new(PdfEditorTool.Select, LocalizedEditorText("Select", "Label", "Select"), LocalizedEditorText("Select", "Hint", "Select and copy text or open links")));
        EditorTools.Add(new(PdfEditorTool.Note, LocalizedEditorText("Note", "Label", "Note"), LocalizedEditorText("Note", "Hint", "Click to add a comment note")));
        EditorTools.Add(new(PdfEditorTool.FreeText, LocalizedEditorText("FreeText", "Label", "Text box"), LocalizedEditorText("FreeText", "Hint", "Draw a free-text annotation")));
        EditorTools.Add(new(PdfEditorTool.Highlight, LocalizedEditorText("Highlight", "Label", "Highlight"), LocalizedEditorText("Highlight", "Hint", "Drag across text or an area")));
        EditorTools.Add(new(PdfEditorTool.Underline, LocalizedEditorText("Underline", "Label", "Underline"), LocalizedEditorText("Underline", "Hint", "Drag across text or an area")));
        EditorTools.Add(new(PdfEditorTool.StrikeOut, LocalizedEditorText("StrikeOut", "Label", "Strikeout"), LocalizedEditorText("StrikeOut", "Hint", "Drag across text or an area")));
        EditorTools.Add(new(PdfEditorTool.Rectangle, LocalizedEditorText("Rectangle", "Label", "Rectangle"), LocalizedEditorText("Rectangle", "Hint", "Draw a rectangle annotation")));
        EditorTools.Add(new(PdfEditorTool.Ellipse, LocalizedEditorText("Ellipse", "Label", "Ellipse"), LocalizedEditorText("Ellipse", "Hint", "Draw an ellipse annotation")));
        EditorTools.Add(new(PdfEditorTool.Line, LocalizedEditorText("Line", "Label", "Line"), LocalizedEditorText("Line", "Hint", "Drag a review line")));
        EditorTools.Add(new(PdfEditorTool.Ink, LocalizedEditorText("Ink", "Label", "Ink"), LocalizedEditorText("Ink", "Hint", "Draw a freehand ink path")));
        EditorTools.Add(new(PdfEditorTool.Stamp, LocalizedEditorText("Stamp", "Label", "Stamp"), LocalizedEditorText("Stamp", "Hint", "Place an annotation stamp")));
        EditorTools.Add(new(PdfEditorTool.AddText, LocalizedEditorText("AddText", "Label", "Add text"), LocalizedEditorText("AddText", "Hint", "Add permanent page text without reflowing existing content")));
        EditorTools.Add(new(PdfEditorTool.AddImage, LocalizedEditorText("AddImage", "Label", "Add image"), LocalizedEditorText("AddImage", "Hint", "Choose and place a PNG or JPEG image")));
        EditorTools.Add(new(PdfEditorTool.Link, LocalizedEditorText("Link", "Label", "Link"), LocalizedEditorText("Link", "Hint", "Draw a URI link hotspot")));
        EditorTools.Add(new(PdfEditorTool.SignatureAppearance, LocalizedEditorText("SignatureAppearance", "Label", "Signature appearance"), LocalizedEditorText("SignatureAppearance", "Hint", "Draw a visual-only signature label; this does not cryptographically sign the PDF")));
        EditorTools.Add(new(PdfEditorTool.Redact, LocalizedEditorText("Redact", "Label", "Redact"), LocalizedEditorText("Redact", "Hint", "Draw an area, review it, then permanently remove intersecting content")));
        SelectedEditorToolChoice = EditorTools[0];

        foreach ((PdfFormFieldCreationKind kind, string label) in new[] {
            (PdfFormFieldCreationKind.Text, "Text field"),
            (PdfFormFieldCreationKind.CheckBox, "Check box"),
            (PdfFormFieldCreationKind.Choice, "Choice field"),
            (PdfFormFieldCreationKind.RadioButtonGroup, "Radio group"),
            (PdfFormFieldCreationKind.Signature, "Signature field"),
            (PdfFormFieldCreationKind.PushButton, "Button")
        }) {
            FormFieldCreationChoices.Add(new(kind, _localizer.GetOrDefault($"Editor.FormKind.{kind}", label)));
        }
        SelectedFormFieldCreationChoice = FormFieldCreationChoices[0];
    }

    private string LocalizedEditorText(string tool, string property, string fallback) =>
        _localizer.GetOrDefault($"Editor.Tool.{tool}.{property}", fallback);

    private async void OnPageEditorGestureCompleted(PdfEditorGesture gesture) {
        bool acceptsEditorGesture = DocumentMode is StudioDocumentMode.Annotate or StudioDocumentMode.Edit ||
                                    DocumentMode == StudioDocumentMode.Protect && ActiveEditorTool == PdfEditorTool.Redact;
        if (_workspace is null ||
            !acceptsEditorGesture ||
            ActiveEditorTool == PdfEditorTool.Select ||
            IsWorkspaceBusy) return;
        PdfWorkspace workspace = _workspace;
        long revision = workspace.Revision;
        PdfEditorTool tool = ActiveEditorTool;
        PdfEditorProperties properties = CreateEditorProperties();
        ErrorMessage = null;
        if (tool == PdfEditorTool.Redact) {
            if (!CanRedact) {
                ErrorMessage = UiText("Editor.RedactionUnavailable");
                return;
            }
            CancelPendingRedaction();
            long generation = _redactionPlanGeneration;
            PdfRedactionPlan? plan = null;
            bool succeeded = await RunStandaloneAsync(async token => {
                OperationStatus = UiText("Editor.PlanningRedaction");
                plan = await workspace.PlanRedactionAsync(gesture, properties, token).ConfigureAwait(true);
                token.ThrowIfCancellationRequested();
            }, CancellationToken.None).ConfigureAwait(true);
            if (!succeeded || plan is null) return;
            if (generation != _redactionPlanGeneration ||
                !ReferenceEquals(_workspace, workspace) ||
                workspace.Revision != revision ||
                ActiveEditorTool != PdfEditorTool.Redact) {
                OperationStatus = UiText("Editor.RedactionPreviewStale");
                return;
            }
            int textMatches = plan.Matches.Count(static match => match.Kind == PdfRedactionMatchKind.TextBlock);
            int imageMatches = plan.Matches.Count(static match => match.Kind == PdfRedactionMatchKind.ImagePlacement);
            int annotationMatches = plan.Matches.Count(static match => match.Kind == PdfRedactionMatchKind.Annotation);
            _pendingRedaction = gesture;
            _pendingRedactionPlan = plan;
            _pendingRedactionWorkspace = workspace;
            _pendingRedactionRevision = revision;
            SetPendingRedactionArea(gesture);
            PendingRedactionSummary = UiFormat(
                "Editor.RedactionSummary",
                gesture.PageNumber,
                textMatches,
                imageMatches,
                annotationMatches);
            return;
        }

        try {
            byte[]? imageBytes = null;
            if (tool == PdfEditorTool.AddImage) {
                string? path = await _pickImage(CancellationToken.None).ConfigureAwait(true);
                if (string.IsNullOrWhiteSpace(path)) return;
                imageBytes = await File.ReadAllBytesAsync(path).ConfigureAwait(true);
                if (!ReferenceEquals(_workspace, workspace) || workspace.Revision != revision) {
                    OperationStatus = UiText("Editor.ImageSelectionStale");
                    return;
                }
            }
            properties = properties with { ImageBytes = imageBytes };
            bool succeeded = await RunMutationAsync(
                token => workspace.ApplyEditorGestureAsync(tool, gesture, properties, token, CreateProgress()),
                CancellationToken.None).ConfigureAwait(true);
            if (succeeded) OperationStatus = UiText("Editor.EditAdded");
        } catch (Exception ex) {
            ErrorMessage = ex.Message;
        }
    }

    private void OnPageObjectSelected(PdfEditorSelection? selection) {
        if (_workspace is null || selection is null) {
            ClearObjectSelection();
            return;
        }

        if (selection.Kind == PdfEditorSelectionKind.Annotation) {
            PdfAnnotation? annotation = _workspace.DocumentInfo.Annotations.FirstOrDefault(candidate =>
                candidate.ObjectNumber == selection.ObjectNumber && candidate.PageNumber == selection.PageNumber);
            if (annotation is null) {
                ClearObjectSelection();
                return;
            }
            SelectedAnnotationSummary = UiFormat(
                "Editor.AnnotationSummary",
                selection.Subtype ?? UiText("Editor.Annotation"),
                selection.PageNumber,
                selection.ObjectNumber);
            SelectedAnnotationContents = annotation.Contents ?? string.Empty;
            SelectedAnnotationAuthor = annotation.Title ?? string.Empty;
            SelectedAnnotationX = annotation.X1;
            SelectedAnnotationY = annotation.Y1;
            SelectedAnnotationWidth = annotation.Width;
            SelectedAnnotationHeight = annotation.Height;
            if (annotation.Color.Count >= 3) EditorColorHex = FormatColor(annotation.Color);
        } else {
            SelectedAnnotationSummary = null;
            SelectedAnnotationContents = string.Empty;
            SelectedAnnotationAuthor = string.Empty;
        }

        SelectedObjectText = selection.Text ?? string.Empty;

        SelectedObject = selection;
        SelectedObjectSummary = selection.Kind switch {
            PdfEditorSelectionKind.Text => UiFormat("Editor.TextSelectionSummary", selection.PageNumber, selection.Text?.Length ?? 0),
            PdfEditorSelectionKind.Image => UiFormat(
                "Editor.ImageSelectionSummary",
                selection.PageNumber,
                selection.ImagePlacement?.Width,
                selection.ImagePlacement?.Height),
            _ => SelectedAnnotationSummary
        };
        foreach (PdfPageViewModel page in Pages) {
            page.SelectedObject = page.PageNumber == selection.PageNumber ? selection : null;
        }
    }

    [RelayCommand]
    private async Task ApplyPendingRedactionAsync(CancellationToken cancellationToken) {
        if (_workspace is null ||
            _pendingRedaction is null ||
            _pendingRedactionPlan is null ||
            _pendingRedactionWorkspace is null) return;
        PdfWorkspace workspace = _pendingRedactionWorkspace;
        PdfRedactionPlan plan = _pendingRedactionPlan;
        long revision = _pendingRedactionRevision;
        if (!ReferenceEquals(_workspace, workspace) || workspace.Revision != revision) {
            CancelPendingRedaction();
            ErrorMessage = UiText("Editor.RedactionReviewStale");
            return;
        }
        PdfVerifiedRedactionResult? proof = null;
        bool succeeded = await RunMutationAsync(async token => {
            proof = await workspace.ApplyVerifiedRedactionAsync(
                plan,
                revision,
                RedactionRemovedMarker,
                token,
                CreateProgress()).ConfigureAwait(true);
        }, cancellationToken).ConfigureAwait(true);
        if (!succeeded || proof is null) return;
        OperationStatus = proof.Evidence.IsVerified
            ? UiFormat(
                "Editor.RedactionVerified",
                proof.Evidence.VerifiedAbsentCount,
                proof.Evidence.AffectedPageNumbers.Count)
            : UiFormat(
                "Editor.RedactionIncomplete",
                proof.Evidence.ResidualCount,
                proof.Evidence.InconclusiveCount);
    }

    [RelayCommand]
    private void CancelPendingRedaction() {
        _redactionPlanGeneration++;
        _pendingRedaction = null;
        _pendingRedactionPlan = null;
        _pendingRedactionWorkspace = null;
        _pendingRedactionRevision = 0;
        PendingRedactionSummary = null;
        SetPendingRedactionArea(null);
    }

    private void SetPendingRedactionArea(PdfEditorGesture? gesture) {
        foreach (PdfPageViewModel page in Pages) {
            page.PendingRedactionArea = gesture is not null && page.PageNumber == gesture.PageNumber
                ? new Rect(gesture.Left, gesture.Top, gesture.Right - gesture.Left, gesture.Bottom - gesture.Top)
                : null;
        }
    }

    [RelayCommand]
    private async Task FillFormFieldAsync(CancellationToken cancellationToken) {
        if (_workspace is null || SelectedFormField is null) return;
        await RunMutationAsync(
            token => _workspace.FillFormFieldAsync(SelectedFormField.Name, SelectedFormField.CreateValue(), flatten: false, token, CreateProgress()),
            cancellationToken).ConfigureAwait(true);
    }

    [RelayCommand]
    private async Task FillAndFlattenFormFieldAsync(CancellationToken cancellationToken) {
        if (_workspace is null || SelectedFormField is null) return;
        await RunMutationAsync(
            token => _workspace.FillFormFieldAsync(SelectedFormField.Name, SelectedFormField.CreateValue(), flatten: true, token, CreateProgress()),
            cancellationToken).ConfigureAwait(true);
    }

    [RelayCommand]
    private async Task FlattenSelectedFormFieldAsync(CancellationToken cancellationToken) {
        if (_workspace is null || SelectedFormField is null) return;
        await RunMutationAsync(
            token => _workspace.FlattenFormFieldAsync(SelectedFormField.Name, token, CreateProgress()),
            cancellationToken).ConfigureAwait(true);
    }

    [RelayCommand]
    private async Task FlattenAllFormFieldsAsync(CancellationToken cancellationToken) {
        if (_workspace is null) return;
        await RunMutationAsync(
            token => _workspace.FlattenFormFieldsAsync(token, CreateProgress()),
            cancellationToken).ConfigureAwait(true);
    }

    [RelayCommand]
    private async Task CreateFormFieldAsync(CancellationToken cancellationToken) {
        if (_workspace is null) return;
        string[] options = ParseFormFieldOptions(NewFormFieldOptions);
        PdfFormFieldCreationKind kind = SelectedFormFieldCreationChoice.Kind;
        bool allowsMultipleSelection = kind == PdfFormFieldCreationKind.Choice && NewFormFieldAllowsMultipleSelection;
        bool isComboBox = kind == PdfFormFieldCreationKind.Choice && NewFormFieldIsComboBox && !allowsMultipleSelection;
        double height = kind == PdfFormFieldCreationKind.RadioButtonGroup
            ? Math.Max(NewFormFieldHeight, GetRequiredRadioGroupHeight(options.Length))
            : NewFormFieldHeight;
        var createOptions = new PdfFormFieldCreateOptions {
            Name = NewFormFieldName?.Trim() ?? string.Empty,
            Kind = kind,
            PageNumber = NewFormFieldPageNumber,
            X = NewFormFieldX,
            Y = NewFormFieldY,
            Width = NewFormFieldWidth,
            Height = height,
            Value = kind == PdfFormFieldCreationKind.CheckBox
                ? NewFormFieldIsChecked ? "Yes" : "Off"
                : NewFormFieldValue ?? string.Empty,
            ChoiceOptions = options,
            IsComboBox = isComboBox,
            Caption = string.IsNullOrWhiteSpace(NewFormFieldCaption) ? "Button" : NewFormFieldCaption.Trim(),
            FieldFlags = allowsMultipleSelection ? 2097152 : 0,
            Style = new PdfFormFieldStyle {
                IsMultiline = kind == PdfFormFieldCreationKind.Text && NewFormFieldIsMultiline,
                IsPassword = kind == PdfFormFieldCreationKind.Text && NewFormFieldIsPassword,
                IsEditableChoice = isComboBox && NewFormFieldAllowsCustomValue
            }
        };
        bool succeeded = await RunMutationAsync(
            token => _workspace.CreateFormFieldAsync(createOptions, token, CreateProgress()),
            cancellationToken).ConfigureAwait(true);
        if (succeeded) {
            SelectedFormField = FormFields.FirstOrDefault(field => string.Equals(field.Name, createOptions.Name, StringComparison.Ordinal));
            NewFormFieldHeight = createOptions.Height;
            NewFormFieldName = GetNextFormFieldName();
        }
    }

    [RelayCommand]
    private async Task ApplyWatermarkAsync(CancellationToken cancellationToken) {
        if (_workspace is null) return;
        await RunMutationAsync(
            token => _workspace.ApplyWatermarkAsync(WatermarkText, token, CreateProgress()),
            cancellationToken).ConfigureAwait(true);
    }

    [RelayCommand]
    private async Task ApplyPageNumbersAsync(CancellationToken cancellationToken) {
        if (_workspace is null) return;
        await RunMutationAsync(
            token => _workspace.ApplyPageNumbersAsync(token, CreateProgress()),
            cancellationToken).ConfigureAwait(true);
    }

    [RelayCommand]
    private async Task UpdateSelectedAnnotationAsync(CancellationToken cancellationToken) {
        if (_workspace is null || SelectedAnnotationObjectNumber is not int objectNumber) return;
        PdfColor color = ParseColor(EditorColorHex);
        string contents = SelectedAnnotationContents;
        string author = SelectedAnnotationAuthor;
        ClearObjectSelection();
        await RunMutationAsync(
            token => _workspace.UpdateAnnotationAsync(objectNumber, contents, author, color, token, CreateProgress()),
            cancellationToken).ConfigureAwait(true);
    }

    [RelayCommand]
    private async Task ReplyToSelectedAnnotationAsync(CancellationToken cancellationToken) {
        if (_workspace is null || SelectedAnnotationObjectNumber is not int objectNumber) return;
        string reply = AnnotationReplyText;
        PdfColor color = ParseColor(EditorColorHex);
        ClearObjectSelection();
        await RunMutationAsync(
            token => _workspace.AddAnnotationReplyAsync(objectNumber, reply, EditorAuthor, color, token, CreateProgress()),
            cancellationToken).ConfigureAwait(true);
        AnnotationReplyText = string.Empty;
    }

    [RelayCommand]
    private async Task FlattenSelectedAnnotationAsync(CancellationToken cancellationToken) {
        if (_workspace is null || SelectedAnnotationObjectNumber is not int objectNumber) return;
        ClearObjectSelection();
        await RunMutationAsync(
            token => _workspace.FlattenAnnotationAsync(objectNumber, token, CreateProgress()),
            cancellationToken).ConfigureAwait(true);
    }

    [RelayCommand]
    private async Task FlattenAllAnnotationsAsync(CancellationToken cancellationToken) {
        if (_workspace is null) return;
        ClearObjectSelection();
        await RunMutationAsync(
            token => _workspace.FlattenAllAnnotationsAsync(token, CreateProgress()),
            cancellationToken).ConfigureAwait(true);
    }

    [RelayCommand]
    private async Task RemoveSelectedAnnotationAsync(CancellationToken cancellationToken) {
        if (_workspace is null || SelectedAnnotationObjectNumber is not int objectNumber) return;
        ClearObjectSelection();
        await RunMutationAsync(
            token => _workspace.RemoveAnnotationAsync(objectNumber, token, CreateProgress()),
            cancellationToken).ConfigureAwait(true);
    }

    private PdfEditorProperties CreateEditorProperties() => new(
        EditorText ?? string.Empty,
        EditorAuthor ?? string.Empty,
        ParseColor(EditorColorHex),
        string.IsNullOrWhiteSpace(EditorStampName) ? "Approved" : EditorStampName.Trim(),
        EditorLinkUri ?? string.Empty,
        Math.Clamp(EditorFontSize, 4D, 144D));

    private void RebuildFormFields() {
        string? selectedName = SelectedFormField?.Name;
        FormFields.Clear();
        if (_workspace is not null) {
            foreach (PdfFormField field in _workspace.DocumentInfo.FormFields.Where(static field => !string.IsNullOrWhiteSpace(field.Name))) {
                FormFields.Add(new PdfFormFieldViewModel(field, _localizer));
            }
        }
        SelectedFormField = FormFields.FirstOrDefault(field => string.Equals(field.Name, selectedName, StringComparison.Ordinal))
            ?? FormFields.FirstOrDefault();
        OnPropertyChanged(nameof(HasFormFields));
        OnPropertyChanged(nameof(CanFillForms));
        OnPropertyChanged(nameof(CanFlattenForms));
        OnPropertyChanged(nameof(CanFillAndFlattenForms));
        OnPropertyChanged(nameof(CanFlattenSelectedFormField));
        OnPropertyChanged(nameof(CanAuthorForms));
    }

    private static string[] ParseFormFieldOptions(string? value) => (value ?? string.Empty)
        .Split(['\r', '\n', ',', ';'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
        .Distinct(StringComparer.Ordinal)
        .ToArray();

    private static double GetRequiredRadioGroupHeight(int optionCount) {
        int count = Math.Max(optionCount, 1);
        const double buttonSize = 14D;
        const double buttonGap = 6D;
        return count * buttonSize + (count - 1) * buttonGap;
    }

    private string GetNextFormFieldName() {
        HashSet<string> existing = FormFields.Select(static field => field.Name).ToHashSet(StringComparer.Ordinal);
        int suffix = 1;
        while (existing.Contains("Field" + suffix.ToString(System.Globalization.CultureInfo.InvariantCulture))) suffix++;
        return "Field" + suffix.ToString(System.Globalization.CultureInfo.InvariantCulture);
    }

    private void ClearObjectSelection() {
        SelectedObject = null;
        SelectedObjectSummary = null;
        SelectedAnnotationSummary = null;
        SelectedAnnotationContents = string.Empty;
        SelectedAnnotationAuthor = string.Empty;
        SelectedObjectText = string.Empty;
        foreach (PdfPageViewModel page in Pages) page.SelectedObject = null;
    }

    private static PdfColor ParseColor(string? value) {
        string hex = (value ?? string.Empty).Trim().TrimStart('#');
        if (hex.Length != 6 ||
            !byte.TryParse(hex.AsSpan(0, 2), System.Globalization.NumberStyles.HexNumber, null, out byte red) ||
            !byte.TryParse(hex.AsSpan(2, 2), System.Globalization.NumberStyles.HexNumber, null, out byte green) ||
            !byte.TryParse(hex.AsSpan(4, 2), System.Globalization.NumberStyles.HexNumber, null, out byte blue)) {
            throw new FormatException("Editor color must be a six-digit hex value such as #E5484D.");
        }
        return PdfColor.FromRgb(red, green, blue);
    }

    private static string FormatColor(IReadOnlyList<double> color) => string.Create(
        System.Globalization.CultureInfo.InvariantCulture,
        $"#{ToColorByte(color[0]):X2}{ToColorByte(color[1]):X2}{ToColorByte(color[2]):X2}");

    private static byte ToColorByte(double value) => (byte)Math.Round(Math.Clamp(value, 0D, 1D) * 255D);
}
