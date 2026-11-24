using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Editor;

namespace OfficeIMO.Studio.Native;

/// <summary>
/// Form authoring, as the desktop Studio's Forms mode does it: adding text, check box, choice and signature fields where
/// they are drawn, changing a field's name, flags and box, and removing a field. Each is one undoable change. Boxes are
/// in points of the page as displayed, which the editor maps into each page's own space through its box and rotation.
/// </summary>
internal sealed partial class StudioNativeDocument {
    private Task AddFormField(StudioNativeFormFieldAdd field, CancellationToken cancellationToken) {
        string name = field.Name?.Trim() is { Length: > 0 } given ? given : throw new ArgumentException("Name the field.", nameof(field));
        PdfFormFieldCreationKind kind = field.Kind switch {
            "text" => PdfFormFieldCreationKind.Text, "checkbox" => PdfFormFieldCreationKind.CheckBox,
            "choice" => PdfFormFieldCreationKind.Choice, "signature" => PdfFormFieldCreationKind.Signature,
            _ => throw new ArgumentException($"“{field.Kind}” isn't a kind of field Studio adds.", nameof(field))
        };
        RequireFormAuthoring();
        if (FormFieldNames().Contains(name)) throw new ArgumentException($"The form already has a field named “{name}”.", nameof(field));
        string[] options = (field.Options ?? []).Select(static option => option?.Trim() ?? string.Empty).Where(static option => option.Length > 0)
            .Distinct(StringComparer.Ordinal).ToArray();
        if (kind == PdfFormFieldCreationKind.Choice && options.Length == 0) throw new ArgumentException("A choice field needs at least one option.", nameof(field));
        PdfPageRectangle box = PageBox(field.Page, field.Bounds ?? throw new ArgumentException("Draw the field's box.", nameof(field)));
        var create = new PdfFormFieldCreateOptions {
            Name = name, Kind = kind, PageNumber = field.Page, X = box.Left, Y = box.Bottom, Width = box.Width, Height = box.Height,
            Value = kind == PdfFormFieldCreationKind.CheckBox ? "Off" : string.Empty,
            ChoiceOptions = options, IsComboBox = kind == PdfFormFieldCreationKind.Choice && field.Combo != false,
            Style = new PdfFormFieldStyle {
                AlternateName = string.IsNullOrWhiteSpace(field.Label) ? null : field.Label.Trim(),
                IsRequired = field.Required, IsReadOnly = field.ReadOnly,
                IsMultiline = kind == PdfFormFieldCreationKind.Text && field.Multiline
            }
        };
        return _workspace.CreateFormFieldAsync(create, cancellationToken);
    }

    private async Task UpdateFormField(StudioNativeFormFieldUpdate update, CancellationToken cancellationToken) {
        string name = update.Name ?? throw new ArgumentException("Name the field to change.", nameof(update));
        RequireFormAuthoring();
        PdfFormField field = _workspace.CreateDocumentSnapshot().Forms.GetForInteraction().Fields.FirstOrDefault(candidate => candidate.Name == name)
            ?? throw new ArgumentException($"The form has no field “{name}”.", nameof(update));
        bool changesDefinition = update.NewName is not null || update.Required is not null || update.ReadOnly is not null;
        if (changesDefinition == update.Bounds is not null)
            throw new ArgumentException("An update changes a field's name and flags, or moves its box, not both.", nameof(update));
        if (update.Bounds is { } bounds) {
            PdfFormWidget widget = field.Widgets.Count == 1 ? field.Widgets[0]
                : throw new ArgumentException("Only a field with one box can be moved.", nameof(update));
            int page = update.Page ?? widget.PageNumber ?? throw new ArgumentException("The field isn't on a page.", nameof(update));
            PdfPageRectangle box = PageBox(page, bounds);
            await _workspace.MoveFormWidgetAsync(name, page, box.Left, box.Bottom, box.Width, box.Height, cancellationToken).ConfigureAwait(false);
            return;
        }
        string newName = update.NewName?.Trim() ?? name;
        if (newName.Length == 0) throw new ArgumentException("A field needs a name.", nameof(update));
        if (newName != name && FormFieldNames().Contains(newName)) throw new ArgumentException($"The form already has a field named “{newName}”.", nameof(update));
        await _workspace.UpdateFormDefinitionAsync(name, newName, update.Required ?? field.IsRequired, update.ReadOnly ?? field.IsReadOnly,
            cancellationToken).ConfigureAwait(false);
    }

    private Task RemoveFormField(string name, CancellationToken cancellationToken) {
        RequireFormAuthoring();
        if (!FormFieldNames().Contains(name)) throw new ArgumentException($"The form has no field “{name}”.", nameof(name));
        return _workspace.RemoveFormDefinitionAsync(name, cancellationToken);
    }

    private void RequireFormAuthoring() {
        if (!_workspace.CanAuthorForms) throw new NotSupportedException("This document's form can't be changed.");
    }

    private HashSet<string> FormFieldNames() =>
        _workspace.CreateDocumentSnapshot().Forms.GetForInteraction().Fields.Select(static field => field.Name)
            .Where(static name => !string.IsNullOrEmpty(name)).Select(static name => name!).ToHashSet(StringComparer.Ordinal);

    /// <summary>A box on the page as displayed, in the page's own space.</summary>
    private PdfPageRectangle PageBox(int page, StudioNativeRectangle bounds) {
        if (page < 1 || page > _workspace.Pages.Count) throw new ArgumentOutOfRangeException(nameof(page), "The page does not exist.");
        if (!(Finite(bounds.Width) >= 4D && Finite(bounds.Height) >= 4D)) throw new ArgumentException("A field needs a width and a height of at least 4 points.", nameof(bounds));
        var gesture = new PdfEditorGesture(page, Finite(bounds.X), Finite(bounds.Y), bounds.X + bounds.Width, bounds.Y + bounds.Height, []);
        return PdfEditorCommandFactory.Create(_workspace.CreateDocumentSnapshot(), PdfEditorTool.Rectangle, gesture,
            new PdfEditorProperties(string.Empty, string.Empty, global::OfficeIMO.OfficeColor.Black, string.Empty, string.Empty, 12D)).Bounds;
    }
}

/// <summary>
/// A field to add: <c>kind</c> is text, checkbox, choice or signature; <c>bounds</c> are in points of the page as
/// displayed. A choice lists its <c>options</c> and is a drop-down unless <c>combo</c> is false.
/// </summary>
internal sealed record StudioNativeFormFieldAdd(string? Kind, string? Name, int Page, StudioNativeRectangle? Bounds,
    IReadOnlyList<string>? Options = null, bool? Combo = null, bool Required = false, bool ReadOnly = false, bool Multiline = false,
    string? Label = null);

/// <summary>A field's new name and flags, or alone its new box, on <c>page</c> or the page it is on.</summary>
internal sealed record StudioNativeFormFieldUpdate(string? Name, string? NewName = null, bool? Required = null, bool? ReadOnly = null,
    int? Page = null, StudioNativeRectangle? Bounds = null);
