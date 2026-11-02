using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.Pdf;

namespace OfficeIMO.Pdf.Benchmarks;

public sealed partial class PdfRuntimeWorkload {
    private string? _destinationPath;
    private string? _expectedDestinationSha256;

    /// <summary>Fingerprint of the equivalent complete forward PDF prepared outside measurement.</summary>
    public string ExpectedDestinationSha256 => _expectedDestinationSha256 ?? throw new InvalidOperationException("Destination reference was not prepared.");
    /// <summary>Fingerprint of the complete persisted artifact, computed after measurement.</summary>
    public string DestinationSha256 { get; private set; } = string.Empty;
    /// <summary>Whether the operation materialized a final PDF byte array before writing the file.</summary>
    public bool FinalByteArrayMaterialized => _operation == "DestinationCopied";
    /// <summary>Writer-reported completed-page store budget.</summary>
    public long PageContentMemoryLimitBytes => _serialization?.PageContentMemoryLimitBytes ?? 0L;
    /// <summary>Writer-reported completed-object store budget.</summary>
    public long ObjectBufferMemoryLimitBytes => _serialization?.ObjectBufferMemoryLimitBytes ?? 0L;
    /// <summary>Whether the writer spilled its completed-page store.</summary>
    public bool PageContentSpilled => _serialization?.PageContentSpilled ?? false;
    /// <summary>Whether the writer spilled its completed-object store.</summary>
    public bool ObjectBufferSpilled => _serialization?.ObjectBufferSpilled ?? false;

    /// <summary>Creates the complete forward artifact fingerprint before any measured call.</summary>
    public void PrepareDestination() {
        if (CorpusProfile != PdfRuntimeCorpus.BusinessProfile) throw new InvalidOperationException("Destination generation requires the business corpus.");
        if (_expectedDestinationSha256 is not null) return;
        using var stream = new MemoryStream();
        PdfRuntimeCorpus.Create(Pages, PdfObjectSerializationMode.ForwardOnly).Save(stream);
        _expectedDestinationSha256 = Convert.ToHexString(SHA256.HashData(stream.ToArray()));
    }

    /// <summary>
    /// Composes and saves one complete PDF file using identical PDF 1.7/forward options.
    /// DestinationDirect saves to FileStream; DestinationCopied saves to MemoryStream, copies and writes the resulting bytes.
    /// Destination directory creation, reference preparation and readback are outside this operation.
    /// </summary>
    public void ExecuteDestination(string operation, string destination, CancellationToken cancellationToken = default) {
        if (operation is not ("DestinationDirect" or "DestinationCopied")) throw new ArgumentException("Unknown destination operation.", nameof(operation));
        if (_expectedDestinationSha256 is null) throw new InvalidOperationException("Destination reference was not prepared.");
        cancellationToken.ThrowIfCancellationRequested();
        _operation = operation;
        _destinationPath = destination;
        PdfDocument document = PdfRuntimeCorpus.Create(Pages, PdfObjectSerializationMode.ForwardOnly);
        PdfSaveResult saved;
        if (operation == "DestinationDirect") {
            using var stream = new FileStream(destination, FileMode.Create, FileAccess.Write, FileShare.None);
            saved = document.SaveWithEvidence(stream, cancellationToken);
        } else {
            using var stream = new MemoryStream();
            saved = document.SaveWithEvidence(stream, cancellationToken);
            byte[] bytes = stream.ToArray();
            cancellationToken.ThrowIfCancellationRequested();
            File.WriteAllBytes(destination, bytes);
        }
        _serialization = saved.Serialization ?? throw new InvalidDataException("Generation returned no serialization report.");
    }

    /// <summary>Reads and qualifies every persisted page and the complete output digest outside measurement.</summary>
    public void ValidateDestination(string directory) {
        if (_destinationPath is null) throw new InvalidOperationException("No destination was written.");
        _output = File.ReadAllBytes(_destinationPath);
        DestinationSha256 = Convert.ToHexString(SHA256.HashData(_output));
        PdfSerializationReport serialization = _serialization ?? throw new InvalidDataException("No destination serialization evidence.");
        if (DestinationSha256 != ExpectedDestinationSha256 || serialization.BytesWritten != _output.LongLength ||
            !serialization.IsForwardOnlyObjectSerialization || serialization.PeakRetainedObjectBytes != 0L) {
            throw new InvalidDataException("Persisted forward PDF digest or writer contract changed.");
        }
        Validate();
        File.WriteAllText(Path.Combine(directory, "destination-evidence.json"), JsonSerializer.Serialize(new {
            Contract = "pdf-forward-destination-v1", Operation = _operation, CorpusProfile, Pages, SourceSha256,
            ExpectedDestinationSha256, DestinationSha256, InputBytes, OutputBytes, VerifiedPages,
            FinalByteArrayMaterialized, Serialization = _serialization
        }, new JsonSerializerOptions { WriteIndented = true }));
    }
}
