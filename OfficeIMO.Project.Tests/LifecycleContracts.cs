namespace OfficeIMO.Project.Tests;

public class LifecycleContracts {
    [Fact]
    public void ReadOnlyDisposeAndSaveOnDisposeRespectCallerStreamOwnership() {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(XmlContracts.Wrap(XmlContracts.TaskXml())));
        using (var readOnly = ProjectDocument.Load(stream, new ProjectLoadOptions { AccessMode = DocumentAccessMode.ReadOnly })) {
            Assert.Throws<InvalidOperationException>(() => readOnly.Name = "change");
            Assert.Throws<InvalidOperationException>(() => readOnly.Save(new MemoryStream()));
            Assert.Contains("Original", readOnly.ToXml());
        }
        Assert.True(stream.CanRead);
        using var destination = new MemoryStream();
        var document = ProjectDocument.Create(destination, new DocumentCreateOptions { PersistenceMode = DocumentPersistenceMode.SaveOnDispose });
        var task = document.Tasks.Add("Persisted");
        Assert.Equal(0, destination.Length);
        document.Dispose();
        Assert.True(destination.CanRead);
        Assert.Throws<ObjectDisposedException>(() => task.Name = "After dispose");
        using var copy = ProjectDocument.Load(destination);
        Assert.Equal("Persisted", copy.Tasks[0].Name);
    }

    [Fact]
    public async Task FileSavesAreAtomicAndReassessCurrentValidation() {
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO.Project.Tests-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string path = Path.Combine(directory, "project.xml");
        try {
            File.WriteAllText(path, "original bytes");
            using var document = ProjectDocument.Create();
            document.Tasks.Add("Valid");
            Assert.Throws<IOException>(() => document.Save(path, new ProjectSaveOptions { FileConflictPolicy = OfficeConversionFileConflictPolicy.FailIfExists }));
            Assert.Equal("original bytes", File.ReadAllText(path));
            Assert.Throws<OperationCanceledException>(() => document.Save(path, cancellationToken: new CancellationToken(true)));
            Assert.Equal("original bytes", File.ReadAllText(path));
            var assessment = document.AssessSave();
            Assert.False(assessment.HasErrors);
            document.Tasks[0].PercentComplete = 101;
            Assert.Throws<InvalidDataException>(() => document.Save(path));
            Assert.Equal("original bytes", File.ReadAllText(path));
            document.Tasks[0].PercentComplete = 0;
            await document.SaveAsync(path);
            using var read = await ProjectDocument.LoadAsync(path);
            Assert.Equal("Valid", read.Tasks[0].Name);
            Assert.Single(Directory.GetFiles(directory));
        } finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void ExplicitStreamSaveLeavesTheAssociatedPathIntact() {
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO.Project.Tests-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string path = Path.Combine(directory, "project.xml");
        try {
            File.WriteAllText(path, XmlContracts.Wrap(XmlContracts.TaskXml()));
            using var document = ProjectDocument.Load(path);
            document.Tasks[0].Name = "Independent copy";
            using var copy = new MemoryStream(); document.Save(copy);
            Assert.True(document.IsModified);
            document.Tasks[0].Name = "Associated path";
            document.Save();
            using var copied = ProjectDocument.Load(new MemoryStream(copy.ToArray()));
            using var associated = ProjectDocument.Load(path);
            Assert.Equal("Independent copy", copied.Tasks[0].Name);
            Assert.Equal("Associated path", associated.Tasks[0].Name);
        } finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public async Task ExplicitAsyncStreamSaveLeavesTheAssociatedStreamIntact() {
        using var associatedStream = new MemoryStream();
        byte[] source = Encoding.UTF8.GetBytes(XmlContracts.Wrap(XmlContracts.TaskXml()));
        associatedStream.Write(source, 0, source.Length); associatedStream.Position = 0;
        using var document = ProjectDocument.Load(associatedStream);
        document.Tasks[0].Name = "Independent copy";
        using var copy = new MemoryStream(); await document.SaveAsync(copy, new ProjectSaveOptions {
            Format = ProjectFileFormat.Mpx4, LossPolicy = OfficeConversionLossPolicy.Allow
        });
        Assert.True(document.IsModified);
        document.Tasks[0].Name = "Associated stream";
        document.Save();
        using var copied = ProjectDocument.Load(new MemoryStream(copy.ToArray()));
        using var associated = ProjectDocument.Load(new MemoryStream(associatedStream.ToArray()));
        Assert.Equal("Independent copy", copied.Tasks[0].Name);
        Assert.Equal("Associated stream", associated.Tasks[0].Name);
    }

    [Fact]
    public void IndependentStreamSaveDoesNotSuppressSaveOnDispose() {
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO.Project.Tests-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string path = Path.Combine(directory, "project.xml");
        try {
            File.WriteAllText(path, XmlContracts.Wrap(XmlContracts.TaskXml()));
            var document = ProjectDocument.Load(path, new ProjectLoadOptions { PersistenceMode = DocumentPersistenceMode.SaveOnDispose });
            document.Tasks[0].Name = "Persist on dispose";
            using var copy = new MemoryStream(); document.Save(copy);
            Assert.True(document.IsModified);
            document.Dispose();
            using var associated = ProjectDocument.Load(path);
            Assert.Equal("Persist on dispose", associated.Tasks[0].Name);
        } finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void NonSeekableInputIsBoundedAndOutputRemainsCallerOwned() {
        using var input = new ForwardStream(Encoding.UTF8.GetBytes(XmlContracts.Wrap(XmlContracts.TaskXml())));
        using var document = ProjectDocument.Load(input);
        Assert.True(input.CanRead);
        using var output = new ForwardStream();
        document.Save(output);
        Assert.True(output.CanWrite);
        using var copy = ProjectDocument.Parse(Encoding.UTF8.GetString(output.Bytes));
        Assert.Equal("Original", copy.Tasks[0].Name);
        using var limited = new ForwardStream(new byte[100]);
        Assert.Throws<InvalidDataException>(() => ProjectDocument.Load(limited, new ProjectLoadOptions { MaxInputBytes = 16 }));
    }

    private sealed class ForwardStream : Stream {
        private readonly MemoryStream _inner;
        internal ForwardStream(byte[]? input = null) { _inner = input == null ? new MemoryStream() : new MemoryStream(input); }
        internal byte[] Bytes => _inner.ToArray();
        public override bool CanRead => _inner.CanRead;
        public override bool CanSeek => false;
        public override bool CanWrite => _inner.CanWrite;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override void Flush() => _inner.Flush();
        public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);
        public override void Write(byte[] buffer, int offset, int count) => _inner.Write(buffer, offset, count);
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        protected override void Dispose(bool disposing) { if (disposing) _inner.Dispose(); base.Dispose(disposing); }
    }
}
