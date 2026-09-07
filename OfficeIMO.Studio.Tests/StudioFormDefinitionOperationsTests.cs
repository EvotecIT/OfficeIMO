using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Shell;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioFormDefinitionOperationsTests {
    [Fact]
    public async Task MovingDefaultsAndRemovalPreserveValuesAndRemainUndoable() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-form-operations-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "source.pdf"), output = Path.Combine(root, "edited.pdf");
        try {
            byte[] original = PdfDocument.Create(compose => {
                compose.Page(page => page.Size(300, 400));
                compose.Page(page => page.Size(300, 400));
            }).Forms.Edit(edit => edit.Create(new() { Name = "Code", Value = "Now", X = 30, Y = 300, Style = new() { MaxLength = 3 } }))
                .ToBytes();
            File.WriteAllBytes(source, original);
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), pickSavePdf: _ => Task.FromResult<string?>(output));
            await model.OpenDocumentAsync(source);
            model.FormHasDefaultValue = true;
            model.FormDefaultValue = "Long default";
            await model.SetFormDefaultCommand.ExecuteAsync(null);
            Assert.NotNull(model.ErrorMessage);
            Assert.False(model.CanUndo);
            Assert.True(model.FormHasDefaultValue);
            Assert.Equal("Long default", model.FormDefaultValue);
            model.FormHasDefaultValue = true;
            model.FormDefaultValue = "ABC";
            await model.SetFormDefaultCommand.ExecuteAsync(null);
            Assert.Null(model.ErrorMessage);
            Assert.Equal("Now", model.SelectedFormField!.TextValue);
            model.FormWidgetPage = 2;
            model.FormWidgetX = 40;
            model.FormWidgetY = 200;
            model.FormWidgetWidth = 210;
            model.FormWidgetHeight = 32;
            await model.MoveFormWidgetCommand.ExecuteAsync(null);
            Assert.Null(model.ErrorMessage);
            Assert.Equal(2, model.SelectedPage!.PageNumber);
            await model.SaveAsCommand.ExecuteAsync(null);
            var field = Assert.Single(PdfDocument.Load(output).Inspect().FormFields);
            Assert.Equal("ABC", field.DefaultValue);
            Assert.Equal("Now", field.Value);
            var widget = Assert.Single(field.Widgets);
            Assert.Equal(2, widget.PageNumber);
            Assert.Equal(40, widget.X1);
            Assert.Equal(200, widget.Y1);
            Assert.Equal(210, widget.Width);
            Assert.Equal(32, widget.Height);
            Assert.Equal(original, File.ReadAllBytes(source));
            await model.RemoveFormDefinitionCommand.ExecuteAsync(null);
            Assert.Empty(model.FormFields);
            await model.UndoCommand.ExecuteAsync(null);
            Assert.Equal("Now", Assert.Single(model.FormFields).TextValue);
            model.FormHasDefaultValue = false;
            await model.SetFormDefaultCommand.ExecuteAsync(null);
            await model.SaveCommand.ExecuteAsync(null);
            Assert.Null(Assert.Single(PdfDocument.Load(output).Inspect().FormFields).DefaultValue);
        } finally { Directory.Delete(root, recursive: true); }
    }
}
