using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Features.Shell;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.VisualTree;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioFormsJourneyTests {
    [Fact]
    public async Task DefinitionEditingProtectsDraftsPreservesFlagsAndSavesPageTabOrder() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-form-definition-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "form.pdf"), output = Path.Combine(root, "edited.pdf");
        try {
            byte[] original = PdfDocument.Create(compose => compose.Page(page => page.Size(300, 300)))
                .Forms.Edit(edit => edit.Create(new() { Name = "First", Value = "Original", Style = new() { IsMultiline = true } })
                    .Create(new() { Name = "Second", Value = "Other" })).ToBytes();
            File.WriteAllBytes(source, original);
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null),
                pickSavePdf: _ => Task.FromResult<string?>(output));
            await model.OpenDocumentAsync(source);
            model.SelectedFormField = model.FormFields.Single(field => field.Name == "First");
            model.FormDefinitionName = "Second";
            await model.ApplyFormDefinitionCommand.ExecuteAsync(null);
            Assert.NotNull(model.ErrorMessage);
            Assert.False(model.CanUndo);
            model.SelectedFormField!.TextValue = "Draft";
            model.FormDefinitionName = "Renamed";
            Assert.False(model.CanEditFormDefinition);
            await model.ApplyFormDefinitionCommand.ExecuteAsync(null);
            Assert.False(model.CanUndo);
            Assert.Equal("Draft", model.SelectedFormField.TextValue);
            model.SelectedFormField.ResetDraftCommand.Execute(null);
            model.FormDefinitionRequired = true;
            model.FormDefinitionReadOnly = true;
            await model.ApplyFormDefinitionCommand.ExecuteAsync(null);
            Assert.Null(model.ErrorMessage);
            Assert.Equal("Renamed", model.SelectedFormField!.Name);
            model.FormTabOrder = PdfPageTabOrder.Column;
            await model.ApplyFormTabOrderCommand.ExecuteAsync(null);
            await model.SaveAsCommand.ExecuteAsync(null);
            var info = PdfDocument.Load(output).Inspect();
            var renamed = info.FormFields.Single(field => field.Name == "Renamed");
            Assert.True(renamed.IsMultiline);
            Assert.True(renamed.IsRequired);
            Assert.True(renamed.IsReadOnly);
            Assert.Equal("Original", renamed.Value);
            Assert.Equal("C", Assert.Single(info.Pages).TabOrder);
            Assert.Equal(original, File.ReadAllBytes(source));
            await model.UndoCommand.ExecuteAsync(null);
            await model.UndoCommand.ExecuteAsync(null);
            Assert.Contains(model.FormFields, field => field.Name == "First" && !field.IsReadOnly);
        } finally { Directory.Delete(root, recursive: true); }
    }

    [Fact]
    public async Task AuthoringPersistsFieldConstraintsAndUndoRemovesTheCreatedField() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-form-author-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "form.pdf");
        string output = Path.Combine(root, "authored.pdf");
        try {
            byte[] original = PdfDocument.Create(compose => compose.Page(page => page.Size(300, 300))).ToBytes();
            File.WriteAllBytes(source, original);
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null),
                pickSavePdf: _ => Task.FromResult<string?>(output));
            await model.OpenDocumentAsync(source);
            model.NewFormFieldName = "Account.Reference";
            model.NewFormFieldDisplayName = "Account reference";
            model.NewFormFieldIsRequired = true;
            model.NewFormFieldIsReadOnly = true;
            model.NewFormFieldMaxLength = 12;
            model.NewFormFieldValue = "Assigned";
            await model.CreateFormFieldCommand.ExecuteAsync(null);
            Assert.Null(model.ErrorMessage);
            Assert.Equal("Account reference", Assert.Single(model.FormFields).DisplayName);
            await model.SaveAsCommand.ExecuteAsync(null);
            var field = Assert.Single(PdfDocument.Load(output).Inspect().FormFields);
            Assert.True(field.IsRequired);
            Assert.True(field.IsReadOnly);
            Assert.Equal(12, field.MaxLength);
            Assert.Equal("Assigned", field.Value);
            Assert.Equal(original, File.ReadAllBytes(source));
            await model.UndoCommand.ExecuteAsync(null);
            Assert.Empty(model.FormFields);
        } finally { Directory.Delete(root, recursive: true); }
    }

    [Fact]
    public async Task ClosingWithSaveRejectsInvalidDraftAndSavesCorrectedValue() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-form-close-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "form.pdf");
        try {
            File.WriteAllBytes(source, PdfDocument.Create(compose => compose.Page(page => page.Size(300, 300)))
                .Forms.Edit(edit => edit.Create(new() { Name = "Code", Value = "OK", Style = new() { MaxLength = 2 } })).ToBytes());
            byte[] original = File.ReadAllBytes(source);
            int prompts = 0;
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null),
                confirmUnsavedChanges: () => { prompts++; return Task.FromResult(UnsavedChangesDecision.Save); });
            await model.OpenDocumentAsync(source);
            model.SelectedFormField!.TextValue = "TOO LONG";
            Assert.False(await model.PrepareCloseDocumentAsync());
            Assert.Equal(1, prompts);
            Assert.NotNull(model.ErrorMessage);
            Assert.Equal(original, File.ReadAllBytes(source));
            model.SelectedFormField.TextValue = "AB";
            Assert.True(await model.PrepareCloseDocumentAsync());
            Assert.Equal(2, prompts);
            Assert.Equal("AB", Assert.Single(PdfDocument.Load(source).Inspect().FormFields).Value);
            Assert.False(model.IsDirty);
        } finally { Directory.Delete(root, recursive: true); }
    }

    [Fact]
    public async Task SaveIncludesAllValidDraftsAndRetainsConflictingDraftsAfterUndo() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-form-save-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "form.pdf");
        string output = Path.Combine(root, "saved.pdf");
        try {
            PdfDocument.Create(compose => compose.Page(page => page.Content(content => {
                content.Item(item => item.TextField("First", value: "Original first"));
                content.Item(item => item.TextField("Second", value: "Original second"));
            }))).Save(source);
            byte[] original = File.ReadAllBytes(source);
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null),
                pickSavePdf: _ => Task.FromResult<string?>(output));
            await model.OpenDocumentAsync(source);
            model.FormFields[0].TextValue = "Edited first";
            model.FormFields[1].TextValue = "Edited second";
            Assert.True(model.IsDirty);
            await model.SaveAsCommand.ExecuteAsync(null);
            Assert.Null(model.ErrorMessage);
            Assert.False(model.HasFormDrafts);
            Assert.False(model.IsDirty);
            var saved = PdfDocument.Load(output).Inspect().FormFieldsByName;
            Assert.Equal("Edited first", saved["First"].Value);
            Assert.Equal("Edited second", saved["Second"].Value);
            Assert.Equal(original, File.ReadAllBytes(source));
            model.FormFields.Single(item => item.Name == "First").TextValue = "New draft";
            await model.UndoCommand.ExecuteAsync(null);
            var draft = model.FormFields.Single(item => item.Name == "First");
            Assert.Equal("New draft", draft.TextValue);
            Assert.True(draft.HasDraftConflict);
            byte[] beforeBlockedSave = File.ReadAllBytes(output);
            await model.SaveCommand.ExecuteAsync(null);
            Assert.Equal(beforeBlockedSave, File.ReadAllBytes(output));
            Assert.NotNull(model.ErrorMessage);
            draft.KeepDraftCommand.Execute(null);
            await model.SaveCommand.ExecuteAsync(null);
            Assert.Null(model.ErrorMessage);
            Assert.Equal("New draft", PdfDocument.Load(output).Inspect().FormFieldsByName["First"].Value);
            Assert.False(model.HasFormDrafts);
        } finally { Directory.Delete(root, recursive: true); }
    }

    [Theory]
    [InlineData(null)]
    [InlineData("Newer draft")]
    [InlineData("Original")]
    public async Task FillAndFlattenConsumesOnlyTheAppliedDraft(string? newerValue) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-form-consumed-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "form.pdf");
        try {
            PdfDocument.Create(compose => compose.Page(page => page.Content(content =>
                content.Item(item => item.TextField("Name", value: "Original"))))).Save(source);
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null));
            await model.OpenDocumentAsync(source);
            var field = model.SelectedFormField!;
            field.TextValue = "Applied value";
            model.PropertyChanged += (_, args) => {
                if (newerValue is not null && args.PropertyName == nameof(model.IsWorkspaceBusy) && model.IsWorkspaceBusy)
                    field.TextValue = newerValue;
            };
            await model.FillAndFlattenFormFieldCommand.ExecuteAsync(null);
            Assert.Null(model.ErrorMessage);
            Assert.Empty(model.FormFields);
            if (newerValue is not null) {
                Assert.Equal(newerValue, Assert.Single(model.UnassignedFormDrafts).TextValue);
                model.DiscardUnassignedFormDraftCommand.Execute(model.UnassignedFormDrafts[0]);
            } else Assert.False(model.HasFormDrafts);
            await model.SaveCommand.ExecuteAsync(null);
            Assert.Null(model.ErrorMessage);
            Assert.Empty(PdfDocument.Load(source).Inspect().FormFields);
            Assert.Contains("Applied value", PdfDocument.Load(source).Read().Text, StringComparison.Ordinal);
            await model.UndoCommand.ExecuteAsync(null);
            Assert.Equal("Original", Assert.Single(model.FormFields).TextValue);
        } finally { Directory.Delete(root, recursive: true); }
    }

    [Theory]
    [InlineData("fill")]
    [InlineData("bulk")]
    [InlineData("save")]
    public async Task ApplyingValuesPreservesConcurrentRevertToOriginal(string operation) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-form-revert-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "form.pdf");
        try {
            PdfDocument.Create(compose => compose.Page(page => page.Content(content =>
                content.Item(item => item.TextField("Name", value: "Original"))))).Save(source);
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null));
            await model.OpenDocumentAsync(source);
            var field = model.SelectedFormField!;
            field.TextValue = "Applied value";
            bool reverted = false;
            model.PropertyChanged += (_, args) => {
                if (!reverted && args.PropertyName == nameof(model.IsWorkspaceBusy) && model.IsWorkspaceBusy) {
                    reverted = true;
                    field.ResetDraftCommand.Execute(null);
                }
            };
            await (operation switch {
                "fill" => model.FillFormFieldCommand.ExecuteAsync(null),
                "bulk" => model.ApplyFormDraftsCommand.ExecuteAsync(null),
                _ => model.SaveCommand.ExecuteAsync(null)
            });
            Assert.Null(model.ErrorMessage);
            Assert.Equal("Original", model.SelectedFormField!.TextValue);
            Assert.True(model.HasFormDrafts);
            Assert.False(model.SelectedFormField.HasDraftConflict);
            if (operation == "save") Assert.Equal("Applied value", PdfDocument.Load(source).Inspect().FormFieldsByName["Name"].Value);
            await model.SaveCommand.ExecuteAsync(null);
            Assert.Null(model.ErrorMessage);
            Assert.Equal("Original", PdfDocument.Load(source).Inspect().FormFieldsByName["Name"].Value);
            Assert.False(model.HasFormDrafts);
        } finally { Directory.Delete(root, recursive: true); }
    }

    [Fact]
    public async Task FlattenRetainsAnUnappliedDraftAndUndoRestoresItsField() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-form-orphan-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "form.pdf");
        try {
            PdfDocument.Create(compose => compose.Page(page => page.Content(content =>
                content.Item(item => item.TextField("Name", value: "Original"))))).Save(source);
            byte[] original = File.ReadAllBytes(source);
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null));
            await model.OpenDocumentAsync(source);
            model.SelectedFormField!.TextValue = "Retained draft";
            await model.FlattenSelectedFormFieldCommand.ExecuteAsync(null);
            Assert.Empty(model.FormFields);
            Assert.Equal("Retained draft", Assert.Single(model.UnassignedFormDrafts).TextValue);
            await model.SaveCommand.ExecuteAsync(null);
            Assert.NotNull(model.ErrorMessage);
            Assert.Equal(original, File.ReadAllBytes(source));
            await model.UndoCommand.ExecuteAsync(null);
            Assert.Empty(model.UnassignedFormDrafts);
            Assert.Equal("Retained draft", Assert.Single(model.FormFields).TextValue);
            model.SelectedFormField!.ResetDraftCommand.Execute(null);
            Assert.False(model.HasFormDrafts);
            Assert.False(model.IsDirty);
        } finally { Directory.Delete(root, recursive: true); }
    }

    [Theory]
    [InlineData(960, 620, false)]
    [InlineData(1280, 800, true)]
    public async Task InspectorExplainsConstraintsAndDisablesReadOnlyEditing(int width, int height, bool dark) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            services.Preferences.Update(current => current with { Theme = dark ? StudioThemePreference.Dark : StudioThemePreference.Light });
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "constraints.pdf");
            File.WriteAllBytes(source, PdfDocument.Create(compose => compose.Page(page => page.Size(300, 400)))
                .Forms.Edit(edit => edit
                    .Create(new() { Name = "Account.Code", Value = "", X = 30, Y = 300,
                        Style = new() { AlternateName = "Account reference", IsRequired = true, MaxLength = 3 } })
                    .Create(new() { Name = "Account.Fixed", Value = "Locked", X = 30, Y = 240,
                        Style = new() { AlternateName = "Assigned account", IsReadOnly = true } })).ToBytes());
            string? evidence = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
            if (!string.IsNullOrWhiteSpace(evidence)) {
                Directory.CreateDirectory(evidence);
                File.Copy(source, Path.Combine(evidence, "form-source.pdf"), overwrite: true);
                File.WriteAllBytes(Path.Combine(evidence, "form-flattened.pdf"), PdfDocument.Load(source).Forms.Flatten("Account.Code").ToBytes());
            }
            var window = new MainWindow(services) { Width = width, Height = height };
            try {
                window.Show();
                var model = window.ViewModel;
                await model.OpenDocumentAsync(source);
                model.DocumentMode = StudioDocumentMode.Forms;
                model.SelectedFormField = model.FormFields.Single(field => field.Name == "Account.Code");
                Assert.Equal("Account reference", model.SelectedFormField.DisplayName);
                Assert.True(model.SelectedFormField.HasValidationMessage);
                Assert.True(model.CanFillForms);
                model.SelectedFormField.TextValue = "TOO LONG";
                Assert.False(model.CanFillForms);
                window.UpdateLayout();
                await model.SelectedPage!.EnsureRenderedAsync();
                Capture(window, $"forms-invalid-{width}-{dark}");
                Assert.Contains(window.GetVisualDescendants().OfType<TextBlock>(), text => text.IsEffectivelyVisible && text.Text?.Contains("at most 3") == true);
                model.SelectedFormField.TextValue = "ABC";
                Assert.True(model.CanPreviewFormAppearance);
                await model.PreviewFormAppearanceCommand.ExecuteAsync(null);
                Assert.Null(model.FormPreviewError);
                Assert.True(model.HasFormPreview);
                Assert.False(model.CanUndo);
                Assert.Equal("", PdfDocument.Load(source).Inspect().FormFieldsByName["Account.Code"].Value);
                window.UpdateLayout();
                var preview = window.GetVisualDescendants().OfType<Image>().Single(image => image.Name == "FormAppearancePreview");
                Assert.True(preview.IsEffectivelyVisible);
                ((Control)preview.Parent!).BringIntoView();
                Capture(window, $"forms-appearance-preview-{width}-{dark}");
                model.SelectedFormField.TextValue = "A🚀";
                Assert.False(model.HasFormPreview);
                Assert.True(model.CanFillForms);
                model.SelectedFormField = model.FormFields.Single(field => field.Name == "Account.Fixed");
                window.UpdateLayout();
                var editor = window.GetVisualDescendants().OfType<TextBox>().Single(box => box.IsEffectivelyVisible && box.Text == "Locked");
                Assert.False(editor.IsEffectivelyEnabled);
                Assert.False(model.CanFillForms);
                Capture(window, $"forms-readonly-{width}-{dark}");
                model.SelectedFormField.TextValue = "Blocked change";
                await model.FillFormFieldCommand.ExecuteAsync(null);
                Assert.False(model.CanUndo);
                foreach (var entry in model.FormFields) entry.ResetDraftCommand.Execute(null);
                model.SelectedFormField = model.FormFields.Single(entry => entry.Name == "Account.Code");
                model.SelectedFormField.TextValue = "Retained draft";
                await model.FlattenSelectedFormFieldCommand.ExecuteAsync(null);
                Assert.Single(model.UnassignedFormDrafts);
                window.UpdateLayout();
                var retainedEditor = window.GetVisualDescendants().OfType<TextBox>().Single(box => box.Text == "Retained draft");
                retainedEditor.BringIntoView();
                var discard = window.GetVisualDescendants().OfType<Button>().Single(button => Equals(button.Content, services.Localizer.Get("Forms.DiscardDraft")));
                await model.SelectedPage!.EnsureRenderedAsync();
                ((Control)discard.Parent!).BringIntoView();
                Capture(window, $"forms-retained-draft-{width}-{dark}");
                Assert.NotNull(discard.Command);
                discard.Command.Execute(discard.CommandParameter);
                Assert.Empty(model.UnassignedFormDrafts);
                await model.UndoCommand.ExecuteAsync(null);
                Assert.False(model.IsDirty);
                var inspector = window.GetVisualDescendants().OfType<FormsInspectorView>().Single();
                inspector.FindControl<Expander>("DefinitionEditor")!.IsExpanded = true;
                inspector.FindControl<TextBox>("DefinitionName")!.Text = "Account.Renamed";
                inspector.FindControl<CheckBox>("DefinitionRequired")!.IsChecked = true;
                window.UpdateLayout();
                await model.SelectedPage!.EnsureRenderedAsync();
                var applyDefinition = inspector.GetVisualDescendants().OfType<Button>()
                    .Single(button => Equals(button.Content, services.Localizer.Get("Forms.ApplyDefinition")));
                ((Control)applyDefinition.Parent!).BringIntoView();
                Capture(window, $"forms-definition-{width}-{dark}");
                Assert.True(applyDefinition.IsEffectivelyEnabled);
                await model.ApplyFormDefinitionCommand.ExecuteAsync(null);
                Assert.Null(model.ErrorMessage);
                Assert.Equal("Account.Renamed", model.SelectedFormField!.Name);
                await model.UndoCommand.ExecuteAsync(null);
                inspector.FindControl<Expander>("DefinitionEditor")!.IsExpanded = false;
                inspector.FindControl<Expander>("TabOrderEditor")!.IsExpanded = true;
                inspector.FindControl<ComboBox>("TabOrderChoice")!.SelectedItem = PdfPageTabOrder.Column;
                window.UpdateLayout();
                var applyTabOrder = inspector.GetVisualDescendants().OfType<Button>()
                    .Single(button => Equals(button.Content, services.Localizer.Get("Forms.ApplyTabOrder")));
                ((Control)applyTabOrder.Parent!).BringIntoView();
                await model.SelectedPage!.EnsureRenderedAsync();
                Capture(window, $"forms-tab-order-{width}-{dark}");
                await model.ApplyFormTabOrderCommand.ExecuteAsync(null);
                Assert.Null(model.ErrorMessage);
                await model.UndoCommand.ExecuteAsync(null);
                inspector.FindControl<Expander>("TabOrderEditor")!.IsExpanded = false;
                inspector.FindControl<TextBox>("NewFieldLabel")!.Text = "Customer reference";
                inspector.FindControl<CheckBox>("NewFieldRequired")!.IsChecked = true;
                inspector.FindControl<CheckBox>("NewFieldReadOnly")!.IsChecked = true;
                inspector.FindControl<NumericUpDown>("NewFieldMaximumLength")!.Value = 20;
                model.NewFormFieldValue = "Assigned";
                window.UpdateLayout();
                await model.SelectedPage!.EnsureRenderedAsync();
                inspector.FindControl<CheckBox>("NewFieldReadOnly")!.BringIntoView();
                Capture(window, $"forms-authoring-{width}-{dark}");
                inspector.FindControl<NumericUpDown>("NewFieldMaximumLength")!.BringIntoView();
                Capture(window, $"forms-authoring-limit-{width}-{dark}");
                var create = inspector.GetVisualDescendants().OfType<Button>()
                    .Single(button => Equals(button.Content, services.Localizer.Get("DocumentWorkspace.CreateField")));
                Assert.NotNull(create.Command);
                await model.CreateFormFieldCommand.ExecuteAsync(create.CommandParameter);
                Assert.Null(model.ErrorMessage);
                Assert.Equal("Customer reference", model.SelectedFormField!.DisplayName);
                Assert.False(model.SelectedFormField.CanFill);
                Assert.Contains("20", model.SelectedFormField.FieldState);
                await model.UndoCommand.ExecuteAsync(null);
                Assert.False(model.IsDirty);
                model.SelectedFormField = model.FormFields.Single(field => field.Name == "Account.Code");
                inspector.FindControl<Expander>("DefinitionEditor")!.IsExpanded = true;
                inspector.FindControl<Expander>("DefaultValueEditor")!.IsExpanded = true;
                inspector.FindControl<CheckBox>("HasFieldDefault")!.IsChecked = true;
                inspector.FindControl<TextBox>("FieldDefaultValue")!.Text = "ZZ";
                await model.SetFormDefaultCommand.ExecuteAsync(null);
                Assert.Null(model.ErrorMessage);
                Assert.Equal("ZZ", model.SelectedFormField!.SavedDefaultValue);
                inspector.FindControl<Expander>("DefaultValueEditor")!.IsExpanded = false;
                inspector.FindControl<Expander>("MoveFieldEditor")!.IsExpanded = true;
                inspector.FindControl<NumericUpDown>("MoveFieldX")!.Value = 60;
                inspector.FindControl<NumericUpDown>("MoveFieldY")!.Value = 180;
                window.UpdateLayout();
                await model.SelectedPage!.EnsureRenderedAsync();
                await model.PreviewFormPlacementCommand.ExecuteAsync(null);
                Assert.Null(model.FormPreviewError);
                Assert.True(model.HasFormPreview);
                await Avalonia.Threading.Dispatcher.UIThread.InvokeAsync(() => window.UpdateLayout(), Avalonia.Threading.DispatcherPriority.Background);
                Capture(window, $"forms-placement-preview-{width}-{dark}");
                await model.MoveFormWidgetCommand.ExecuteAsync(null);
                Assert.Null(model.ErrorMessage);
                Assert.Equal(60, model.SelectedFormField!.SingleWidget!.X1);
                Assert.Equal(180, model.SelectedFormField.SingleWidget.Y1);
                window.UpdateLayout();
                await model.SelectedPage!.EnsureRenderedAsync();
                var move = inspector.GetVisualDescendants().OfType<Button>()
                    .Single(button => Equals(button.Content, services.Localizer.Get("Forms.ApplyPlacement")));
                move.BringIntoView();
                Capture(window, $"forms-moved-{width}-{dark}");
                await model.RemoveFormDefinitionCommand.ExecuteAsync(null);
                Assert.DoesNotContain(model.FormFields, field => field.Name == "Account.Code");
                await model.UndoCommand.ExecuteAsync(null);
                await model.UndoCommand.ExecuteAsync(null);
                await model.UndoCommand.ExecuteAsync(null);
                Assert.False(model.IsDirty);
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }

    [Theory]
    [InlineData(PdfFormFieldCreationKind.Choice)]
    [InlineData(PdfFormFieldCreationKind.RadioButtonGroup)]
    public void UnselectedFieldDoesNotInventAValue(PdfFormFieldCreationKind kind) {
        var document = PdfDocument.Create(compose => compose.Page(page => page.Size(300, 300)))
            .Forms.Edit(edit => edit.Create(new() {
                Name = "Decision", Kind = kind, ChoiceOptions = ["Accept", "Decline"],
                X = 20, Y = 30, Width = 180, Height = 50
            })).ToDocument().Forms.Fill(new Dictionary<string, string> { ["Decision"] = "" });
        var model = new PdfFormFieldViewModel(Assert.Single(document.Inspect().FormFields));
        Assert.Null(model.SelectedChoice);
        Assert.All(model.Choices, choice => Assert.False(choice.IsSelected));
        Assert.Contains(model.CreateValue().Values[0], new[] { "", "Off" });
    }

    [Fact]
    public async Task ApplyingOneFieldPreservesAnotherFieldsDraft() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-form-drafts-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "form.pdf");
        try {
            PdfDocument.Create(compose => compose.Page(page => page.Content(content => {
                content.Item(item => item.TextField("First", value: "Original first"));
                content.Item(item => item.TextField("Second", value: "Original second"));
            }))).Save(source);
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null));
            await model.OpenDocumentAsync(source);
            model.FormFields.Single(field => field.Name == "First").TextValue = "Unapplied draft";
            model.SelectedFormField = model.FormFields.Single(field => field.Name == "Second");
            model.SelectedFormField.TextValue = "Applied second";
            await model.FillFormFieldCommand.ExecuteAsync(null);
            Assert.Null(model.ErrorMessage);
            Assert.Equal("Unapplied draft", model.FormFields.Single(field => field.Name == "First").TextValue);
            Assert.Equal("Applied second", model.FormFields.Single(field => field.Name == "Second").TextValue);
            await model.UndoCommand.ExecuteAsync(null);
            Assert.Equal("Unapplied draft", model.FormFields.Single(field => field.Name == "First").TextValue);
            Assert.Equal("Original second", model.FormFields.Single(field => field.Name == "Second").TextValue);
            string other = Path.Combine(root, "other.pdf");
            File.Copy(source, other);
            model.FormFields.Single(field => field.Name == "First").ResetDraftCommand.Execute(null);
            await model.OpenDocumentAsync(other);
            Assert.Equal("Original first", model.FormFields.Single(field => field.Name == "First").TextValue);
        } finally { Directory.Delete(root, recursive: true); }
    }
    private static void Capture(Window window, string name) {
        window.UpdateLayout();
        using var bitmap = window.CaptureRenderedFrame();
        Assert.NotNull(bitmap);
        string? folder = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
        if (string.IsNullOrWhiteSpace(folder)) return;
        Directory.CreateDirectory(folder);
        bitmap.Save(Path.Combine(folder, name + ".png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }
}
