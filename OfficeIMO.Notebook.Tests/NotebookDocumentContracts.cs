namespace OfficeIMO.Notebook.Tests;

public sealed class NotebookDocumentContracts {
    [Fact]
    public void SaveEditedFutureNotebookPreservesUnknownCellAndOutputPayloads() {
        string path = Fixture("test4plus.ipynb");
        using JsonDocument sourceJson = JsonDocument.Parse(File.ReadAllBytes(path));
        using NotebookDocument notebook = NotebookDocument.Load(path);
        Assert.Equal(99, notebook.NbFormatMinor);
        Assert.Equal("future-1", notebook.Cells[0].Id);
        Assert.Equal(NotebookCellKind.Unknown, notebook.Cells[9].Kind);

        notebook.Cells[0].Source = "# A revised overview";
        using JsonDocument saved = JsonDocument.Parse(notebook.ToBytes());
        JsonElement savedCells = saved.RootElement.GetProperty("cells");
        JsonElement sourceCells = sourceJson.RootElement.GetProperty("cells");

        Assert.Equal(99, saved.RootElement.GetProperty("nbformat_minor").GetInt32());
        Assert.True(JsonElement.DeepEquals(sourceJson.RootElement.GetProperty("extra"), saved.RootElement.GetProperty("extra")));
        Assert.Equal("future-1", savedCells[0].GetProperty("id").GetString());
        Assert.Equal("# A revised overview", savedCells[0].GetProperty("source").GetString());
        Assert.True(JsonElement.DeepEquals(sourceCells[9], savedCells[9]));
        Assert.True(JsonElement.DeepEquals(sourceCells[3].GetProperty("outputs"), savedCells[3].GetProperty("outputs")));
        Assert.Contains(savedCells.EnumerateArray(), cell => cell.TryGetProperty("outputs", out JsonElement outputs) &&
            outputs.EnumerateArray().Any(output => output.GetProperty("output_type").GetString() == "future output" &&
                output.GetProperty("some key")[0].GetString() == "some data"));
    }

    [Fact]
    public void LegacyNotebookRequiresExplicitNormalizationAndKeepsHeadingStreamAndMimeAlternatives() {
        using NotebookDocument notebook = NotebookDocument.Load(Fixture("test3.ipynb"));
        Assert.Equal(3, notebook.NbFormat);
        Assert.Equal(9, notebook.Cells.Count);
        Assert.Equal(NotebookCellKind.Markdown, notebook.Cells[0].Kind);
        Assert.Equal("# nbconvert latex test", notebook.Cells[0].Source);
        NotebookOutput stream = Assert.Single(notebook.Cells[3].Outputs);
        Assert.Equal(NotebookOutputKind.Stream, stream.Kind);
        Assert.Equal("stdout", stream.Name);
        Assert.Equal("hello\n", stream.Text);
        NotebookOutput html = Assert.Single(notebook.Cells[5].Outputs);
        Assert.Equal(NotebookOutputKind.ExecuteResult, html.Kind);
        Assert.Contains("text/html", html.Data.Keys);
        Assert.Contains("text/plain", html.Data.Keys);
        Assert.Throws<OfficeConversionException>(() => notebook.ToBytes());

        using JsonDocument output = JsonDocument.Parse(notebook.ToBytes(new NotebookSaveOptions { AllowLoss = true }));
        JsonElement root = output.RootElement;
        Assert.Equal(4, root.GetProperty("nbformat").GetInt32());
        Assert.Equal(5, root.GetProperty("nbformat_minor").GetInt32());
        Assert.Equal("markdown", root.GetProperty("cells")[0].GetProperty("cell_type").GetString());
        JsonElement image = root.GetProperty("cells")[8].GetProperty("outputs")[0];
        Assert.Equal("execute_result", image.GetProperty("output_type").GetString());
        Assert.StartsWith("iVBOR", image.GetProperty("data").GetProperty("image/png").GetString(), StringComparison.Ordinal);
    }

    [Fact]
    public async Task NewNotebookCanBePreparedAndReadThroughCallerOwnedStreams() {
        using NotebookDocument notebook = NotebookDocument.Create();
        notebook.AddMarkdownCell("# Lab results");
        NotebookCell code = notebook.AddCodeCell("print(42)");
        code.ExecutionCount = 1;
        code.AddOutput(NotebookOutput.CreateStream("stdout", "42\n"));

        using var destination = new MemoryStream();
        await notebook.SaveAsync(destination);
        Assert.True(destination.CanWrite);
        byte[] bytes = destination.ToArray();
        using JsonDocument json = JsonDocument.Parse(bytes);
        JsonElement root = json.RootElement;
        Assert.Equal(4, root.GetProperty("nbformat").GetInt32());
        JsonElement cell = root.GetProperty("cells")[1];
        Assert.Equal(code.Id, cell.GetProperty("id").GetString());
        Assert.Equal(1, cell.GetProperty("execution_count").GetInt32());
        Assert.Equal("42\n", cell.GetProperty("outputs")[0].GetProperty("text").GetString());

        using var stream = new ForwardOnlyStream(bytes);
        using NotebookDocument loaded = await NotebookDocument.LoadAsync(stream, new NotebookLoadOptions { MaxInputBytes = bytes.Length });
        Assert.Equal("# Lab results", loaded.Cells[0].Source);
        Assert.False(stream.Closed);
    }

    [Fact]
    public void LoadMemorySnapshotsBytesBeforeLazilyReadingCellSource() {
        byte[] bytes = Encoding.UTF8.GetBytes("""
            {"nbformat":4,"nbformat_minor":5,"metadata":{},"cells":[
              {"id":"memo","cell_type":"markdown","source":["Hello ","world"],"metadata":{}}
            ]}
            """);
        using NotebookDocument notebook = NotebookDocument.Load(bytes);
        Array.Clear(bytes, 0, bytes.Length);
        Assert.Equal("Hello world", Assert.Single(notebook.Cells).Source);
    }

    [Fact]
    public async Task SavePathPreservesExistingDestinationOnNormalizationLossAndCancellation() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-notebook-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string destination = Path.Combine(directory, "published.ipynb");
            File.WriteAllText(destination, "existing file");
            const UnixFileMode ownerOnly = UnixFileMode.UserRead | UnixFileMode.UserWrite;
            if (!OperatingSystem.IsWindows()) File.SetUnixFileMode(destination, ownerOnly);
            using NotebookDocument legacy = NotebookDocument.Load(Fixture("test3.ipynb"));
            await Assert.ThrowsAsync<OfficeConversionException>(() => legacy.SaveAsync(destination));
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
                legacy.SaveAsync(destination, new NotebookSaveOptions { AllowLoss = true }, cancellation.Token));
            Assert.Equal("existing file", File.ReadAllText(destination));
            Assert.Single(Directory.EnumerateFiles(directory));

            await legacy.SaveAsync(destination, new NotebookSaveOptions { AllowLoss = true });
            using JsonDocument saved = JsonDocument.Parse(File.ReadAllBytes(destination));
            Assert.Equal(4, saved.RootElement.GetProperty("nbformat").GetInt32());
            Assert.Single(Directory.EnumerateFiles(directory));
            if (!OperatingSystem.IsWindows()) {
                Assert.Equal(ownerOnly, File.GetUnixFileMode(destination));
                legacy.Save(destination, new NotebookSaveOptions { AllowLoss = true });
                Assert.Equal(ownerOnly, File.GetUnixFileMode(destination));
            }
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Theory]
    [InlineData("cells", "MaxCells")]
    [InlineData("totalOutputs", "MaxTotalOutputs")]
    [InlineData("outputsPerCell", "MaxOutputsPerCell")]
    [InlineData("bytes", "MaxInputBytes")]
    public async Task NonseekableLoadRejectsInputsExceedingConfiguredLimits(string limit, string expectedOption) {
        const string content = """
            {"nbformat":4,"nbformat_minor":5,"metadata":{},"cells":[
              {"cell_type":"code","id":"one","source":"","metadata":{},"execution_count":null,
               "outputs":[{"output_type":"stream","name":"stdout","text":"1"}]},
              {"cell_type":"code","id":"two","source":"","metadata":{},"execution_count":null,
               "outputs":[{"output_type":"stream","name":"stdout","text":"2"}]}
            ]}
            """;
        byte[] bytes = Encoding.UTF8.GetBytes(content);
        var options = new NotebookLoadOptions {
            MaxInputBytes = limit == "bytes" ? bytes.Length - 1 : bytes.Length,
            MaxCells = limit == "cells" ? 1 : 2,
            MaxOutputsPerCell = limit == "outputsPerCell" ? 0 : 1,
            MaxTotalOutputs = limit == "totalOutputs" ? 1 : 2
        };
        using var stream = new ForwardOnlyStream(bytes);
        InvalidDataException error = await Assert.ThrowsAsync<InvalidDataException>(() => NotebookDocument.LoadAsync(stream, options));
        Assert.Contains(expectedOption, error.Message, StringComparison.Ordinal);
        Assert.False(stream.Closed);
    }

    [Fact]
    public void OlderNotebookGeneratesDeterministicIdsWithoutDisplacingValidOriginals() {
        const string content = """
            {"nbformat":4,"nbformat_minor":2,"metadata":{},"cells":[
              {"cell_type":"markdown","source":"one"},
              {"cell_type":"markdown","id":"cell-1","source":"two"},
              {"cell_type":"markdown","id":"cell-1","source":"three"}
            ]}
            """;
        byte[] bytes = Encoding.UTF8.GetBytes(content);
        using NotebookDocument notebook = NotebookDocument.Load(bytes);
        using JsonDocument json = JsonDocument.Parse(notebook.ToBytes());
        Assert.Equal(5, json.RootElement.GetProperty("nbformat_minor").GetInt32());
        Assert.Equal(new[] { "cell-1-2", "cell-1", "cell-3" }, json.RootElement.GetProperty("cells").EnumerateArray()
            .Select(static cell => cell.GetProperty("id").GetString()));
    }

    private static string Fixture(string fileName) => Path.Combine(AppContext.BaseDirectory, "Fixtures", fileName);

    private sealed class ForwardOnlyStream(byte[] bytes) : Stream {
        private readonly MemoryStream _inner = new(bytes, writable: false);
        internal bool Closed { get; private set; }
        public override bool CanRead => !Closed;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }
        public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);
        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) =>
            _inner.ReadAsync(buffer, offset, count, cancellationToken);
        public override void Flush() => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        protected override void Dispose(bool disposing) {
            if (disposing) _inner.Dispose();
            Closed = true;
            base.Dispose(disposing);
        }
    }
}
