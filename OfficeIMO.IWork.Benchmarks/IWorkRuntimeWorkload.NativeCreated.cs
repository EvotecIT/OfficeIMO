using System.Security.Cryptography;

namespace OfficeIMO.IWork.Benchmarks;

public sealed partial class IWorkRuntimeWorkload {
    private readonly IWorkKeynoteDocument? _createdNativeDocument;

    /// <summary>Creates a scalable native package with the OfficeIMO writer outside measurement.</summary>
    /// <remarks>This is owned-producer scale evidence, separate from independent Apple fixtures.</remarks>
    public static IWorkRuntimeWorkload CreateNativeKeynote(int slides) => new(slides);

    private IWorkRuntimeWorkload(int slides) {
        if (slides < 1) throw new ArgumentOutOfRangeException(nameof(slides));
        _kind = IWorkDocumentKind.Keynote;
        _units = slides;
        _createdNativeDocument = IWorkKeynoteDocument.Create();
        for (int index = 1; index <= slides; index++)
            _createdNativeDocument.AddSlide().AddText(IWorkScaleInput.Text(index), 20, 20, 920, 100);
        _input = _createdNativeDocument.SaveBytes();
        InputSha256 = Convert.ToHexString(SHA256.HashData(_input));
    }
}
