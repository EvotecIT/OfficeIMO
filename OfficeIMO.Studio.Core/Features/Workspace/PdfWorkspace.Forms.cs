using OfficeIMO.Pdf;

namespace OfficeIMO.Studio.Features.Workspace;

internal sealed partial class PdfWorkspace {
    internal Task MoveFormWidgetAsync(string name, int page, double x, double y, double width, double height,
        CancellationToken token, IProgress<PdfWorkspaceProgress>? progress = null) =>
        MutateBytesAsync(PdfWorkspaceOperationKind.FormAuthor, "Moved form field " + name, [page],
            bytes => LoadDocument(bytes).Forms.Edit(edit => edit.Move(name, page, x, y, width, height)).ToBytes(), token, progress);

    internal Task RemoveFormDefinitionAsync(string name, CancellationToken token, IProgress<PdfWorkspaceProgress>? progress = null) =>
        MutateBytesAsync(PdfWorkspaceOperationKind.FormAuthor, "Removed form field " + name, [],
            bytes => LoadDocument(bytes).Forms.Edit(edit => edit.Remove(name)).ToBytes(), token, progress);

    internal Task SetFormDefaultAsync(string name, string? value, CancellationToken token, IProgress<PdfWorkspaceProgress>? progress = null) =>
        MutateBytesAsync(PdfWorkspaceOperationKind.FormAuthor, "Updated form default " + name, [],
            bytes => LoadDocument(bytes).Forms.Edit(edit => edit.SetDefaultValue(name, value)).ToBytes(), token, progress);

    internal Task UpdateFormDefinitionAsync(string name, string newName, bool required, bool readOnly,
        CancellationToken token, IProgress<PdfWorkspaceProgress>? progress = null) =>
        MutateBytesAsync(PdfWorkspaceOperationKind.FormAuthor, "Updated form field " + name, [], bytes => {
            var document = LoadDocument(bytes);
            var field = document.Inspect().FormFields.Single(field => field.Name == name);
            int flags = ((field.Flags ?? 0) & ~3) | (readOnly ? 1 : 0) | (required ? 2 : 0);
            return document.Forms.Edit(edit => {
                edit.SetFlags(name, flags);
                if (!string.Equals(name, newName, StringComparison.Ordinal)) edit.Rename(name, newName);
            }).ToBytes();
        }, token, progress);

    internal Task SetFormTabOrderAsync(int pageNumber, PdfPageTabOrder order,
        CancellationToken token, IProgress<PdfWorkspaceProgress>? progress = null) =>
        MutateBytesAsync(PdfWorkspaceOperationKind.FormAuthor, "Updated form tab order", [pageNumber],
            bytes => LoadDocument(bytes).Forms.Edit(edit => edit.SetTabOrder(pageNumber, order)).ToBytes(), token, progress);

    internal Task FillFormFieldsAsync(IReadOnlyDictionary<string, PdfFormFieldValue> values,
        CancellationToken cancellationToken, IProgress<PdfWorkspaceProgress>? progress = null) {
        var snapshot = values.ToDictionary(pair => pair.Key, pair => PdfFormFieldValue.FromValues(pair.Value.Values), StringComparer.Ordinal);
        return MutateBytesAsync(PdfWorkspaceOperationKind.FormFill, "Applied form field edits", [], bytes => {
            var document = LoadDocument(bytes);
            var plan = document.PlanMutation(PdfMutationOperation.FillFormFields, snapshot.Keys);
            return (plan.ExecutionMode == PdfMutationExecutionMode.AppendOnly
                ? document.Forms.AppendRevision(snapshot) : document.Forms.Fill(snapshot)).ToBytes();
        }, cancellationToken, progress);
    }

    internal Task FillFormFieldAsync(
        string fieldName,
        string value,
        bool flatten,
        CancellationToken cancellationToken,
        IProgress<PdfWorkspaceProgress>? progress = null) =>
        FillFormFieldAsync(fieldName, PdfFormFieldValue.From(value ?? string.Empty), flatten, cancellationToken, progress);

    internal Task FillFormFieldAsync(
        string fieldName,
        PdfFormFieldValue value,
        bool flatten,
        CancellationToken cancellationToken,
        IProgress<PdfWorkspaceProgress>? progress = null) {
        if (string.IsNullOrWhiteSpace(fieldName)) throw new ArgumentException("Choose a form field.", nameof(fieldName));
        ArgumentNullException.ThrowIfNull(value);
        string normalizedName = fieldName.Trim();
        IReadOnlyDictionary<string, PdfFormFieldValue> values = new Dictionary<string, PdfFormFieldValue>(StringComparer.Ordinal) {
            [normalizedName] = value
        };
        return MutateBytesAsync(
            flatten ? PdfWorkspaceOperationKind.FormFlatten : PdfWorkspaceOperationKind.FormFill,
            flatten ? "Filled and flattened form field " + normalizedName : "Filled form field " + normalizedName,
            Array.Empty<int>(),
            bytes => {
                PdfDocument document = LoadDocument(bytes);
                if (flatten) return document.Forms.FillAndFlatten(values).ToBytes();
                PdfMutationPlan plan = document.PlanMutation(PdfMutationOperation.FillFormFields, values.Keys);
                return (plan.ExecutionMode == PdfMutationExecutionMode.AppendOnly
                    ? document.Forms.AppendRevision(values)
                    : document.Forms.Fill(values)).ToBytes();
            },
            cancellationToken,
            progress);
    }

    internal Task CreateFormFieldAsync(
        PdfFormFieldCreateOptions options,
        CancellationToken cancellationToken,
        IProgress<PdfWorkspaceProgress>? progress = null) {
        ArgumentNullException.ThrowIfNull(options);
        if (!CanAuthorForms) throw new InvalidOperationException("This document cannot safely author form fields.");
        if (options.PageNumber < 1 || options.PageNumber > Pages.Count) {
            throw new ArgumentOutOfRangeException(nameof(options), $"Form field page must be between 1 and {Pages.Count}.");
        }

        PdfFormFieldCreateOptions snapshot = new() {
            Name = options.Name,
            Kind = options.Kind,
            PageNumber = options.PageNumber,
            X = options.X,
            Y = options.Y,
            Width = options.Width,
            Height = options.Height,
            Value = options.Value,
            DefaultValue = options.DefaultValue,
            FieldFlags = options.FieldFlags,
            WidgetFlags = options.WidgetFlags,
            ChoiceOptions = options.ChoiceOptions.ToArray(),
            CheckedValueName = options.CheckedValueName,
            IsComboBox = options.IsComboBox,
            Caption = options.Caption,
            FontSize = options.FontSize,
            RadioButtonSize = options.RadioButtonSize,
            RadioButtonGap = options.RadioButtonGap,
            Style = options.Style?.Clone()
        };
        string normalizedName = snapshot.Name?.Trim() ?? string.Empty;
        snapshot.Name = normalizedName;
        return MutateBytesAsync(
            PdfWorkspaceOperationKind.FormAuthor,
            "Created " + GetFormKindDescription(snapshot.Kind) + " " + normalizedName,
            new[] { snapshot.PageNumber },
            bytes => LoadDocument(bytes).Forms.Edit(edit => edit.Create(snapshot)).ToBytes(),
            cancellationToken,
            progress);
    }

    internal Task FlattenFormFieldAsync(
        string fieldName,
        CancellationToken cancellationToken,
        IProgress<PdfWorkspaceProgress>? progress = null) {
        if (string.IsNullOrWhiteSpace(fieldName)) throw new ArgumentException("Choose a form field.", nameof(fieldName));
        string normalizedName = fieldName.Trim();
        return MutateBytesAsync(
            PdfWorkspaceOperationKind.FormFlatten,
            "Flattened form field " + normalizedName,
            Array.Empty<int>(),
            bytes => LoadDocument(bytes).Forms.Flatten(normalizedName).ToBytes(),
            cancellationToken,
            progress);
    }

    internal Task FlattenFormFieldsAsync(
        CancellationToken cancellationToken,
        IProgress<PdfWorkspaceProgress>? progress = null) =>
        MutateBytesAsync(
            PdfWorkspaceOperationKind.FormFlatten,
            "Flattened all form fields",
            Array.Empty<int>(),
            bytes => LoadDocument(bytes).Forms.Flatten().ToBytes(),
            cancellationToken,
            progress);

    private static string GetFormKindDescription(PdfFormFieldCreationKind kind) => kind switch {
        PdfFormFieldCreationKind.Text => "text field",
        PdfFormFieldCreationKind.CheckBox => "check box",
        PdfFormFieldCreationKind.Choice => "choice field",
        PdfFormFieldCreationKind.Signature => "signature field",
        PdfFormFieldCreationKind.RadioButtonGroup => "radio group",
        PdfFormFieldCreationKind.PushButton => "button",
        _ => throw new ArgumentOutOfRangeException(nameof(kind), kind, "Unsupported form field kind.")
    };
}
