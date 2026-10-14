using System.IO;
using System.IO.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.Word;
using OfficeIMO.Excel;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Security.Tests {
    public sealed partial class OfficeVbaCarrierRecoveryTests {
        [Theory]
        [InlineData("excel", false)]
        [InlineData("word", false)]
        [InlineData("powerpoint", false)]
        [InlineData("excel", true)]
        [InlineData("word", true)]
        [InlineData("powerpoint", true)]
        public void FailedExistingPayloadUpdateRestoresProjectAndSignaturesAndAllowsRetry(string host, bool failSignatureDelete) {
            using RecoveryPackage storage = new RecoveryPackage();
            using OpenXmlPackage document = host == "excel"
                ? SpreadsheetDocument.Create(storage, SpreadsheetDocumentType.MacroEnabledWorkbook)
                : host == "word" ? WordprocessingDocument.Create(storage, WordprocessingDocumentType.MacroEnabledDocument)
                : PresentationDocument.Create(storage, PresentationDocumentType.MacroEnabledPresentation);
            OpenXmlPart owner = document is SpreadsheetDocument workbook ? workbook.AddWorkbookPart()
                : document is WordprocessingDocument word ? word.AddMainDocumentPart()
                : ((PresentationDocument)document).AddPresentationPart();
            OfficeVbaProject project = OfficeVbaProject.Create(); project.AddModule("Helpers", "'original");
            byte[] original = project.Write().GetBytes();
            OfficeVbaProjectPartEditor.Apply(owner, null, original, new OfficeVbaWriteOptions(), out VbaProjectPart part);
            ExtendedPart signature = part.AddExtendedPart("http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature",
                "application/vnd.ms-office.vbaProjectSignature", "bin", "signature");
            byte[] signatureBytes = { 1, 2, 3, 4 };
            using (MemoryStream input = new MemoryStream(signatureBytes)) signature.FeedData(input);
            ExtendedPart agile = part.AddExtendedPart("http://schemas.microsoft.com/office/2014/relationships/vbaProjectSignatureAgile",
                "application/vnd.ms-office.vbaProjectSignatureAgile", "bin", "agile");
            using (MemoryStream input = new MemoryStream(new byte[] { 5, 6 })) agile.FeedData(input);
            project.SetModuleSource("Helpers", "'replacement"); byte[] replacement = project.Write().GetBytes();
            storage.FailPayloadWrite = !failSignatureDelete;
            storage.FailSignatureDelete = failSignatureDelete;

            Assert.Throws<IOException>(() => OfficeVbaProjectPartEditor.Apply(owner, part, replacement,
                new OfficeVbaWriteOptions { AllowSignatureRemoval = true }, out _));

            Assert.Equal(original, OfficeVbaProjectPartEditor.Read(part, original.Length));
            Assert.Equal(2, part.GetPartsOfType<ExtendedPart>().Count());
            ExtendedPart restored = (ExtendedPart)part.GetPartById("signature");
            Assert.Equal("signature", part.GetIdOfPart(restored));
            using (Stream input = restored.GetStream(FileMode.Open, FileAccess.Read)) {
                using MemoryStream output = new MemoryStream(); input.CopyTo(output); Assert.Equal(signatureBytes, output.ToArray());
            }
            using (Stream input = part.GetPartById("agile").GetStream(FileMode.Open, FileAccess.Read)) {
                using MemoryStream output = new MemoryStream(); input.CopyTo(output); Assert.Equal(new byte[] { 5, 6 }, output.ToArray());
            }
            Assert.True(OfficeVbaProjectPartEditor.Apply(owner, part, replacement,
                new OfficeVbaWriteOptions { AllowSignatureRemoval = true }, out _));
            Assert.Equal(replacement, OfficeVbaProjectPartEditor.Read(part, replacement.Length));
            Assert.Empty(part.GetPartsOfType<ExtendedPart>());
        }

        [Fact]
        public void ApplyingIdenticalWordSourceCompletesMissingSupplementalData() {
            string path = Path.Combine(Path.GetTempPath(), "OfficeIMO-vba-data-" + Guid.NewGuid().ToString("N") + ".docm");
            try {
                OfficeVbaProject project = OfficeVbaProject.Create();
                project.AddDocumentModuleWithIdentity("ThisDocument", "", "1Normal.ThisDocument");
                byte[] bytes = project.Write().GetBytes();
                using (WordDocument word = WordDocument.Create(path)) {
                    word.AddMacro(bytes);
                    word.SetVbaProject(OfficeVbaProject.Load(bytes));
                    Assert.Equal(bytes, word.ExtractMacros());
                    word.Save();
                }
                using WordprocessingDocument package = WordprocessingDocument.Open(path, false);
                Assert.NotNull(package.MainDocumentPart!.VbaProjectPart!.VbaDataPart?.VbaSuppData);
            } finally { if (File.Exists(path)) File.Delete(path); }
        }

        [Fact]
        public void ApplyingIdenticalExcelSourceCompletesMissingWorkbookCodeName() {
            string path = Path.Combine(Path.GetTempPath(), "OfficeIMO-vba-code-" + Guid.NewGuid().ToString("N") + ".xlsm");
            try {
                OfficeVbaProject project = OfficeVbaProject.Create();
                project.AddDocumentModule("OwnedWorkbook", "", new Guid("00020819-0000-0000-C000-000000000046"));
                byte[] bytes = project.Write().GetBytes();
                using (ExcelDocument excel = ExcelDocument.Create(path)) {
                    excel.AddWorksheet("Data"); excel.AddMacro(bytes);
                    excel.SetVbaProject(OfficeVbaProject.Load(bytes)); excel.Save();
                }
                using SpreadsheetDocument package = SpreadsheetDocument.Open(path, false);
                Assert.Equal("OwnedWorkbook", package.WorkbookPart!.Workbook!.GetFirstChild<WorkbookProperties>()?.CodeName?.Value);
            } finally { if (File.Exists(path)) File.Delete(path); }
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void SupplementalWriteFailureRestoresExistingOrRemovesNewProject(bool existing) {
            using RecoveryPackage storage = new RecoveryPackage();
            using WordprocessingDocument document = WordprocessingDocument.Create(storage, WordprocessingDocumentType.MacroEnabledDocument);
            MainDocumentPart owner = document.AddMainDocumentPart();
            OfficeVbaProject project = OfficeVbaProject.Create(); project.AddModule("Helpers", "'original");
            byte[] original = project.Write().GetBytes(); VbaProjectPart? previous = null;
            if (existing) OfficeVbaProjectPartEditor.Apply(owner, null, original, new OfficeVbaWriteOptions(), out previous);
            project.SetModuleSource("Helpers", "'replacement"); byte[] replacement = project.Write().GetBytes();
            storage.FailDataWrite = true;
            Action<VbaProjectPart> initialize = part => {
                VbaDataPart data = part.AddNewPart<VbaDataPart>(); data.VbaSuppData = new DocumentFormat.OpenXml.Office.Word.VbaSuppData(); data.VbaSuppData.Save();
            };
            Assert.Throws<IOException>(() => OfficeVbaProjectPartEditor.Apply(owner, previous, replacement, new OfficeVbaWriteOptions(), out _, initialize));
            if (existing) { Assert.Equal(original, OfficeVbaProjectPartEditor.Read(owner.VbaProjectPart!, original.Length)); Assert.Null(owner.VbaProjectPart!.VbaDataPart); }
            else Assert.Null(owner.VbaProjectPart);
            Assert.True(OfficeVbaProjectPartEditor.Apply(owner, owner.VbaProjectPart, replacement, new OfficeVbaWriteOptions(), out VbaProjectPart retry, initialize));
            Assert.Equal(replacement, OfficeVbaProjectPartEditor.Read(retry, replacement.Length)); Assert.NotNull(retry.VbaDataPart!.VbaSuppData);
        }

        [Fact]
        public void FailedRestorationReportsBothStorageFailures() {
            using RecoveryPackage storage = new RecoveryPackage();
            using WordprocessingDocument document = WordprocessingDocument.Create(storage, WordprocessingDocumentType.MacroEnabledDocument);
            MainDocumentPart owner = document.AddMainDocumentPart();
            OfficeVbaProject project = OfficeVbaProject.Create(); project.AddModule("Helpers", "'original");
            OfficeVbaProjectPartEditor.Apply(owner, null, project.Write().GetBytes(), new OfficeVbaWriteOptions(), out VbaProjectPart part);
            project.SetModuleSource("Helpers", "'replacement"); storage.PayloadWriteFailures = 2;
            AggregateException failure = Assert.Throws<AggregateException>(() => OfficeVbaProjectPartEditor.Apply(owner, part,
                project.Write().GetBytes(), new OfficeVbaWriteOptions(), out _));
            Assert.Equal(2, failure.InnerExceptions.Count); Assert.All(failure.InnerExceptions, exception => Assert.IsType<IOException>(exception));
        }

        [Fact]
        public void RecoveryBudgetRejectsAnOversizedSubgraphBeforeChangingPayloadOrSignature() {
            using RecoveryPackage storage = new RecoveryPackage();
            using WordprocessingDocument document = WordprocessingDocument.Create(storage, WordprocessingDocumentType.MacroEnabledDocument);
            MainDocumentPart owner = document.AddMainDocumentPart();
            OfficeVbaProject project = OfficeVbaProject.Create(); project.AddModule("Helpers", "'old");
            byte[] original = project.Write().GetBytes();
            OfficeVbaProjectPartEditor.Apply(owner, null, original, new OfficeVbaWriteOptions(), out VbaProjectPart part);
            ExtendedPart signature = part.AddExtendedPart("http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature",
                "application/vnd.ms-office.vbaProjectSignature", "bin", "signature");
            using (MemoryStream input = new MemoryStream(new byte[] { 1 })) signature.FeedData(input);
            project.SetModuleSource("Helpers", "'new");
            Assert.Throws<InvalidDataException>(() => OfficeVbaProjectPartEditor.Apply(owner, part, project.Write().GetBytes(),
                new OfficeVbaWriteOptions { AllowSignatureRemoval = true, MaximumRecoveryBytes = original.Length }, out _));
            Assert.Equal(original, OfficeVbaProjectPartEditor.Read(part, original.Length)); Assert.Same(signature, part.GetPartById("signature"));
        }

        // Fail once at the SDK package boundary, then leave storage usable for restoration.
        private sealed class RecoveryPackage : Package {
            private readonly Package _storage = Open(new MemoryStream(), FileMode.Create, FileAccess.ReadWrite);
            internal bool FailPayloadWrite;
            internal bool FailSignatureDelete;
            internal bool FailDataWrite;
            internal bool FailProjectDelete;
            internal bool FailDataDelete;
            internal int PayloadWriteFailures;
            internal RecoveryPackage() : base(FileAccess.ReadWrite) { }
            protected override PackagePart CreatePartCore(Uri uri, string contentType, CompressionOption compression) {
                return Wrap(_storage.CreatePart(uri, contentType, compression));
            }
            protected override PackagePart GetPartCore(Uri uri) => Wrap(_storage.GetPart(uri));
            protected override PackagePart[] GetPartsCore() => _storage.GetParts().Select(Wrap).ToArray();
            protected override void DeletePartCore(Uri uri) {
                if (FailDataDelete && _storage.PartExists(uri) && _storage.GetPart(uri).ContentType == "application/vnd.ms-word.vbaData+xml") {
                    FailDataDelete = false; throw new IOException("Synthetic supplemental data deletion failure.");
                }
                if (FailProjectDelete && _storage.PartExists(uri) && _storage.GetPart(uri).ContentType == "application/vnd.ms-office.vbaProject") {
                    FailProjectDelete = false; throw new IOException("Synthetic project deletion failure.");
                }
                if (FailSignatureDelete && _storage.PartExists(uri) && _storage.GetPart(uri).ContentType.Contains("vbaProjectSignatureAgile")) {
                    FailSignatureDelete = false; throw new IOException("Synthetic signature deletion failure.");
                }
                _storage.DeletePart(uri);
            }
            protected override void FlushCore() => _storage.Flush();
            protected override void Dispose(bool disposing) { if (disposing) ((IDisposable)_storage).Dispose(); base.Dispose(disposing); }
            private PackagePart Wrap(PackagePart part) => new RecoveryPart(this, part);
        }
        private sealed class RecoveryPart : PackagePart {
            private readonly RecoveryPackage _owner; private readonly PackagePart _storage;
            internal RecoveryPart(RecoveryPackage owner, PackagePart storage) : base(owner, storage.Uri, storage.ContentType, storage.CompressionOption) { _owner=owner; _storage=storage; }
            protected override Stream GetStreamCore(FileMode mode, FileAccess access) {
                Stream stream = _storage.GetStream(mode, access);
                if ((_owner.FailPayloadWrite || _owner.PayloadWriteFailures > 0) && ContentType == "application/vnd.ms-office.vbaProject" && (access & FileAccess.Write) != 0) {
                    return new PartialWriteStream(stream, _owner, false);
                }
                if (_owner.FailDataWrite && ContentType == "application/vnd.ms-word.vbaData+xml" && (access & FileAccess.Write) != 0) return new PartialWriteStream(stream, _owner, true);
                return stream;
            }
        }
        private sealed class PartialWriteStream : Stream {
            private readonly Stream _storage;
            private readonly RecoveryPackage _owner;
            private readonly bool _isData;
            internal PartialWriteStream(Stream storage, RecoveryPackage owner, bool isData) { _storage=storage; _owner=owner; _isData=isData; }
            public override void Write(byte[] buffer, int offset, int count) {
                if (_isData ? _owner.FailDataWrite : _owner.FailPayloadWrite || _owner.PayloadWriteFailures > 0) {
                    if (_isData) _owner.FailDataWrite=false;
                    if (_owner.PayloadWriteFailures > 0) _owner.PayloadWriteFailures--;
                    _owner.FailPayloadWrite=false; _storage.Write(buffer, offset, Math.Min(count, 16)); throw new IOException("Synthetic partial payload write failure.");
                }
                _storage.Write(buffer, offset, count);
            }
            public override bool CanRead => _storage.CanRead;
            public override bool CanSeek => _storage.CanSeek;
            public override bool CanWrite => _storage.CanWrite;
            public override long Length => _storage.Length;
            public override long Position { get => _storage.Position; set => _storage.Position=value; }
            public override void Flush() => _storage.Flush();
            public override int Read(byte[] buffer,int offset,int count) => _storage.Read(buffer,offset,count);
            public override long Seek(long offset,SeekOrigin origin) => _storage.Seek(offset,origin);
            public override void SetLength(long value) => _storage.SetLength(value);
            protected override void Dispose(bool disposing) { if (disposing) _storage.Dispose(); base.Dispose(disposing); }
        }
    }
}
