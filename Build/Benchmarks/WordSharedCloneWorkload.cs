using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;

/// <summary>One public assessment over a package whose inert payload has several incoming relationships.</summary>
public sealed class WordSharedCloneWorkload {
    private readonly byte[] _input;
    private readonly byte[] _payload = new byte[65536];
    private readonly int _parents;
    private LegacyDocWriteAssessment _result;

    public WordSharedCloneWorkload(int parents) {
        _parents = parents;
        new Random(323).NextBytes(_payload);
        using (var input = new MemoryStream()) {
            using (var package = WordprocessingDocument.Create(input, WordprocessingDocumentType.Document)) {
                ExtendedPart shared = package.AddExtendedPart("urn:officeimo:benchmark:shared", "application/octet-stream", ".bin", "shared");
                using (var content = new MemoryStream(_payload)) shared.FeedData(content);
                MainDocumentPart main = package.AddMainDocumentPart();
                main.Document = new W.Document(new W.Body(new W.Paragraph(new W.Run(new W.Text("Shared-graph assessment")))));
                main.AddPart(shared, "shared");
                for (int index = 0; index < parents; index++) {
                    ExtendedPart parent = main.AddExtendedPart("urn:officeimo:benchmark:parent", "application/xml", ".xml");
                    using (var content = new MemoryStream(new byte[] { 60, 112, 47, 62 })) parent.FeedData(content);
                    parent.AddPart(shared, "shared");
                }
            }
            _input = input.ToArray();
        }
        using (var input = new MemoryStream(_input)) {
            using (var package = WordprocessingDocument.Open(input, false)) ValidateInput(package);
        }
    }

    public long AllocatedBytes { get; private set; }

    public void Assess() {
        long before = GC.GetAllocatedBytesForCurrentThread();
        try {
            using (var input = new MemoryStream(_input)) {
                using (WordDocument document = WordDocument.Load(input)) _result = document.AssessLegacyDocWrite();
            }
        } finally {
            AllocatedBytes = GC.GetAllocatedBytesForCurrentThread() - before;
        }
    }

    public void Validate() {
        if (_result == null || !_result.IsSupported || !(_result.EncodedByteCount > 0)) {
            throw new InvalidDataException("Shared-graph assessment did not encode the document: " + _result?.Message);
        }
    }

    private void ValidateInput(WordprocessingDocument package) {
        OpenXmlPart shared = package.GetPartById("shared");
        if (!ReferenceEquals(shared, package.MainDocumentPart.GetPartById("shared"))) {
            throw new InvalidDataException("The shared root/main relationship changed.");
        }
        OpenXmlPart[] parents = package.MainDocumentPart.Parts.Select(pair => pair.OpenXmlPart)
            .Where(part => part.RelationshipType == "urn:officeimo:benchmark:parent").ToArray();
        if (parents.Length != _parents || parents.Any(part => !ReferenceEquals(shared, part.GetPartById("shared")))) {
            throw new InvalidDataException("The shared parent relationships changed.");
        }
        using (Stream content = shared.GetStream()) {
            using (var copy = new MemoryStream()) {
                content.CopyTo(copy);
                if (!_payload.SequenceEqual(copy.ToArray())) throw new InvalidDataException("Shared bytes changed.");
            }
        }
    }
}
