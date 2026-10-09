using System.IO;
using System.IO.Packaging;
using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Core.Internal;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.Word;

namespace OfficeIMO.Security.Tests {
    public sealed partial class OfficeVbaCarrierRecoveryTests {
        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void RecoveryRejectsLoadedMediaRelationshipsBeforeCloningOrMutation(bool nested) {
            OfficeVbaProject project = OfficeVbaProject.Create(); project.AddModule("Helpers", "'original");
            byte[] original = project.Write().GetBytes();
            using MemoryStream storage = new MemoryStream();
            Uri sourceUri;
            using (WordprocessingDocument document = WordprocessingDocument.Create(storage, WordprocessingDocumentType.MacroEnabledDocument)) {
                MainDocumentPart main = document.AddMainDocumentPart();
                main.Document = new DocumentFormat.OpenXml.Wordprocessing.Document(new DocumentFormat.OpenXml.Wordprocessing.Body());
                VbaProjectPart part = main.AddNewPart<VbaProjectPart>();
                using (MemoryStream input = new MemoryStream(original)) part.FeedData(input);
                OpenXmlPart source = part;
                if (nested) {
                    source = part.AddExtendedPart("urn:officeimo:opaque", "application/octet-stream", "bin", "opaque");
                    using MemoryStream input = new MemoryStream(new byte[] { 1 }); source.FeedData(input);
                }
                sourceUri = source.Uri;
            }
            storage.Position = 0;
            using (Package package = Package.Open(storage, FileMode.Open, FileAccess.ReadWrite)) {
                Uri mediaUri = new Uri("/word/media/recovery.bin", UriKind.Relative);
                PackagePart media = package.CreatePart(mediaUri, "application/octet-stream");
                using (Stream output = media.GetStream()) output.Write(new byte[128], 0, 128);
                package.GetPart(sourceUri).CreateRelationship(PackUriHelper.GetRelativeUri(sourceUri, mediaUri), TargetMode.Internal,
                    "http://schemas.microsoft.com/office/2007/relationships/media", "media");
            }
            storage.Position = 0;
            using WordprocessingDocument loaded = WordprocessingDocument.Open(storage, true);
            VbaProjectPart previous = loaded.MainDocumentPart!.VbaProjectPart!;
            OpenXmlPart referenced = nested ? previous.GetPartById("opaque") : previous;
            Assert.Single(referenced.DataPartReferenceRelationships);
            project.SetModuleSource("Helpers", "'replacement");
            Assert.Throws<InvalidDataException>(() => OfficeVbaProjectPartEditor.Apply(loaded.MainDocumentPart, previous, project.Write().GetBytes(),
                new OfficeVbaWriteOptions { MaximumRecoveryBytes = original.Length + (nested ? 1 : 0) }, out _));
            Assert.Equal(original, OfficeVbaProjectPartEditor.Read(previous, original.Length));
            Assert.Single(referenced.DataPartReferenceRelationships);
        }

        [Theory]
        [InlineData(false, false)]
        [InlineData(true, false)]
        [InlineData(false, true)]
        [InlineData(true, true)]
        public void FailedFinalNamedModuleDeletionRetainsSourceAndReportsInvalidatedStorage(bool opaque, bool failChild) {
            using RecoveryPackage storage = new RecoveryPackage();
            using WordprocessingDocument package = WordprocessingDocument.Create(storage, WordprocessingDocumentType.MacroEnabledDocument);
            MainDocumentPart main = package.AddMainDocumentPart();
            OfficeVbaProject project = OfficeVbaProject.Create(); project.AddModule("Helpers", "'original");
            byte[] original = project.Write().GetBytes();
            if (opaque) {
                Assert.True(OfficeCompoundFileReader.TryRead(original, out OfficeCompoundFile? compound, out _));
                original = OfficeCompoundFileWriter.Rewrite(compound!, new Dictionary<string, byte[]> { ["VBA/dir"] = new byte[] { 2 } });
            }
            OfficeVbaProjectPartEditor.Apply(main, null, original, new OfficeVbaWriteOptions(), out VbaProjectPart part);
            string relationshipId = main.GetIdOfPart(part);
            VbaDataPart data = part.AddNewPart<VbaDataPart>("data");
            data.VbaSuppData = new DocumentFormat.OpenXml.Office.Word.VbaSuppData(); data.VbaSuppData.Save();
            byte[] supplemental;
            using (Stream input = data.GetStream(FileMode.Open, FileAccess.Read)) {
                using MemoryStream output = new MemoryStream(); input.CopyTo(output); supplemental = output.ToArray();
            }
            // Substitute the real faulting SDK package at the existing document boundary, without a production test hook.
            using WordDocument word = WordDocument.Create();
            word.OpenXmlDocument.Dispose();
            typeof(WordDocument).GetField("_wordprocessingDocument", BindingFlags.Instance | BindingFlags.NonPublic)!.SetValue(word, package);
            storage.FailProjectDelete = !failChild; storage.FailDataDelete = failChild;
            // System.IO.Packaging marks the failed physical part deleted before invoking
            // DeletePartCore. Restore the reachable carrier, but require disposal when that
            // invalidated cache also prevents cleanup of the detached physical part.
            AggregateException failure = Assert.Throws<AggregateException>(() => word.RemoveMacro("Helpers"));
            Assert.IsType<IOException>(failure.InnerExceptions[0]);
            Assert.IsType<InvalidOperationException>(failure.InnerExceptions[1]);
            Assert.Contains("Discard this document instance", failure.Message);
            Assert.True(word.HasMacros);
            Assert.Equal(original, word.ExtractMacros());
            Assert.Equal(relationshipId, main.GetIdOfPart(main.VbaProjectPart!));
            using (Stream input = main.VbaProjectPart!.GetPartById("data").GetStream(FileMode.Open, FileAccess.Read)) {
                using MemoryStream output = new MemoryStream(); input.CopyTo(output); Assert.Equal(supplemental, output.ToArray());
            }
        }

        [Fact]
        public void NewCarrierCleanupFailureReportsBothStorageFailures() {
            using RecoveryPackage storage = new RecoveryPackage();
            using WordprocessingDocument document = WordprocessingDocument.Create(storage, WordprocessingDocumentType.MacroEnabledDocument);
            MainDocumentPart main = document.AddMainDocumentPart();
            OfficeVbaProject project = OfficeVbaProject.Create(); project.AddModule("Helpers", "'source");
            storage.FailPayloadWrite = true; storage.FailProjectDelete = true;
            AggregateException failure = Assert.Throws<AggregateException>(() => OfficeVbaProjectPartEditor.Apply(main, null,
                project.Write().GetBytes(), new OfficeVbaWriteOptions(), out _));
            Assert.Equal(2, failure.InnerExceptions.Count);
            Assert.All(failure.InnerExceptions, exception => Assert.IsType<IOException>(exception));
            Assert.Contains("Discard this document instance", failure.Message);
        }
    }
}
