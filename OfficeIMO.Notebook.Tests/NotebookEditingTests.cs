namespace OfficeIMO.Notebook.Tests;

public sealed class NotebookEditingTests {
    [Fact]
    public void MoveAndRemovePreserveRetainedCellPayloadsAtTheirNewIndices() {
        const string content = """
            {
              "nbformat":4,"nbformat_minor":99,"metadata":{"vendor":{"schema":"future"}},
              "cells":[
                {"id":"remove","cell_type":"raw","source":"discard","metadata":{}},
                {"id":"attachment","cell_type":"markdown","source":["![asset](attachment:figure.png)"],
                 "metadata":{"caption":"figure"},"attachments":{"figure.png":{"image/png":"iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGMQTJsJAAGcARGmCxx6AAAAAElFTkSuQmCC"}}},
                {"id":"opaque","cell_type":"future_cell","source":"opaque","metadata":{},"future":{"items":[12,34]}},
                {"id":"code","cell_type":"code","source":"print(3)","metadata":{"custom":"retained"},"execution_count":3,
                 "outputs":[{"output_type":"display_data","data":{"application/x.custom":{"value":3}},"metadata":{}}]}
              ]
            }
            """;
        using JsonDocument original = JsonDocument.Parse(content);
        using NotebookDocument notebook = NotebookDocument.Load(Encoding.UTF8.GetBytes(content));
        NotebookCell opaque = notebook.Cells[2];

        Assert.Throws<ArgumentOutOfRangeException>("toIndex", () => notebook.MoveCell(2, 4));
        Assert.Same(opaque, notebook.Cells[2]);
        notebook.MoveCell(3, 1);
        notebook.RemoveCell(0);
        notebook.MoveCell(1, 2);
        Assert.Same(opaque, notebook.Cells[1]);

        byte[] saved = notebook.ToBytes();
        using JsonDocument written = JsonDocument.Parse(saved);
        JsonElement originalCells = original.RootElement.GetProperty("cells");
        JsonElement writtenCells = written.RootElement.GetProperty("cells");
        Assert.True(JsonElement.DeepEquals(originalCells[3], writtenCells[0]));
        Assert.True(JsonElement.DeepEquals(originalCells[2], writtenCells[1]));
        Assert.True(JsonElement.DeepEquals(originalCells[1], writtenCells[2]));
        Assert.True(JsonElement.DeepEquals(original.RootElement.GetProperty("metadata"), written.RootElement.GetProperty("metadata")));

        using NotebookDocument reopened = NotebookDocument.Load(saved);
        Assert.Equal(new[] { "code", "opaque", "attachment" }, reopened.Cells.Select(cell => cell.Id));
        Assert.Equal(NotebookCellKind.Unknown, reopened.Cells[1].Kind);
    }

    [Fact]
    public void ClearingOutputsAndClearingExecutionCountHaveIndependentSaveEffects() {
        const string content = """
            {"nbformat":4,"nbformat_minor":5,"metadata":{},"cells":[
              {"id":"keep-count","cell_type":"code","source":"1","metadata":{},"execution_count":1,
               "outputs":[{"output_type":"stream","name":"stdout","text":"removed"}]},
              {"id":"keep-output","cell_type":"code","source":"2","metadata":{},"execution_count":2,
               "outputs":[{"output_type":"stream","name":"stdout","text":"retained"}]},
              {"id":"old-output-shape","cell_type":"code","source":"3","metadata":{},"execution_count":3,"outputs":null}
            ]}
            """;
        using NotebookDocument notebook = NotebookDocument.Load(Encoding.UTF8.GetBytes(content));
        notebook.Cells[0].ClearOutputs();
        notebook.Cells[1].ExecutionCount = null;
        notebook.Cells[2].ClearOutputs();

        using JsonDocument saved = JsonDocument.Parse(notebook.ToBytes());
        JsonElement cells = saved.RootElement.GetProperty("cells");
        Assert.Equal(0, cells[0].GetProperty("outputs").GetArrayLength());
        Assert.Equal(1, cells[0].GetProperty("execution_count").GetInt32());
        Assert.Equal(JsonValueKind.Null, cells[1].GetProperty("execution_count").ValueKind);
        Assert.Equal("retained", cells[1].GetProperty("outputs")[0].GetProperty("text").GetString());
        Assert.Equal(0, cells[2].GetProperty("outputs").GetArrayLength());
    }

    [Fact]
    public void AuthoredCodeLanguageSurvivesSavingToTheDefaultPath() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-notebook-editing-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string path = Path.Combine(directory, "notebook.ipynb");
        try {
            using NotebookDocument notebook = NotebookDocument.Create(path);
            Assert.Equal(path, notebook.FilePath);
            NotebookCell cell = notebook.AddCodeCell("Write-Output 'hello'", "pwsh");
            Assert.Equal("pwsh", cell.Language);
            notebook.Save();

            using NotebookDocument reopened = NotebookDocument.Load(path);
            Assert.Equal(path, reopened.FilePath);
            Assert.Equal("pwsh", Assert.Single(reopened.Cells).Language);
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }
}
