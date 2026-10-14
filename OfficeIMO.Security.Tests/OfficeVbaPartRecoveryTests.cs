using System.IO;
using System.IO.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Security.Tests;

public sealed class OfficeVbaPartRecoveryTests {
    [Theory]
    [InlineData("excel")]
    [InlineData("word")]
    [InlineData("powerpoint")]
    public void FailedFirstPayloadWriteRemovesItsPartAndAllowsRetry(string host) {
        using var storage = new FailingPackage();
        using OpenXmlPackage document = host == "excel"
            ? SpreadsheetDocument.Create(storage, SpreadsheetDocumentType.MacroEnabledWorkbook)
            : host == "word" ? WordprocessingDocument.Create(storage, WordprocessingDocumentType.MacroEnabledDocument)
            : PresentationDocument.Create(storage, PresentationDocumentType.MacroEnabledPresentation);
        OpenXmlPart owner = document is SpreadsheetDocument workbook ? workbook.AddWorkbookPart()
            : document is WordprocessingDocument word ? word.AddMainDocumentPart()
            : ((PresentationDocument)document).AddPresentationPart();
        var project = OfficeVbaProject.Create(); project.AddModule("Helpers", "'source");
        byte[] bytes = project.Write().GetBytes();
        Assert.Throws<IOException>(() => OfficeVbaProjectPartEditor.Apply(owner, null, bytes, new OfficeVbaWriteOptions(), out _));
        Assert.Empty(owner.GetPartsOfType<VbaProjectPart>());
        storage.FailWrites = false;
        OfficeVbaProjectPartEditor.Apply(owner, null, bytes, new OfficeVbaWriteOptions(), out VbaProjectPart retry);
        Assert.Equal(bytes, OfficeVbaProjectPartEditor.Read(retry, bytes.Length));
    }

    // Failure is injected at PackagePart.GetStreamCore, the actual SDK storage boundary.
    private sealed class FailingPackage : Package {
        private readonly Package _storage;
        internal bool FailWrites = true;
        internal FailingPackage() : base(FileAccess.ReadWrite) => _storage = Open(new MemoryStream(), FileMode.Create, FileAccess.ReadWrite);
        protected override PackagePart CreatePartCore(Uri uri, string contentType, CompressionOption compression) => Wrap(_storage.CreatePart(uri, contentType, compression));
        protected override PackagePart GetPartCore(Uri uri) => Wrap(_storage.GetPart(uri));
        protected override PackagePart[] GetPartsCore() => _storage.GetParts().Select(Wrap).ToArray();
        protected override void DeletePartCore(Uri uri) => _storage.DeletePart(uri);
        protected override void FlushCore() => _storage.Flush();
        protected override void Dispose(bool disposing) { if (disposing) ((IDisposable)_storage).Dispose(); base.Dispose(disposing); }
        private PackagePart Wrap(PackagePart part) => new FailingPart(this, part);
    }

    private sealed class FailingPart : PackagePart {
        private readonly FailingPackage _owner;
        private readonly PackagePart _storage;
        internal FailingPart(FailingPackage owner, PackagePart storage)
            : base(owner, storage.Uri, storage.ContentType, storage.CompressionOption) { _owner = owner; _storage = storage; }
        protected override Stream GetStreamCore(FileMode mode, FileAccess access) =>
            _owner.FailWrites && ContentType == "application/vnd.ms-office.vbaProject" && (access & FileAccess.Write) != 0
                ? throw new IOException("Synthetic VBA storage write failure.") : _storage.GetStream(mode, access);
    }
}
