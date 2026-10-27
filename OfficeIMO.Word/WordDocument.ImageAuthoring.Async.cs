using OfficeIMO.Drawing;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Word;

public partial class WordDocument {
    /// <summary>Downloads a bounded remote image and embeds it in a new paragraph.</summary>
    /// <remarks>The remote loader enforces its network, redirect, time and response-size policy.</remarks>
    public async Task<WordImage> AddImageAsync(Uri uri, OfficeLength? width = null, OfficeLength? height = null,
        WordImageOptions? options = null, OfficeRemoteImageLoadOptions? remoteImageOptions = null,
        CancellationToken cancellationToken = default) {
        if (uri == null) throw new ArgumentNullException(nameof(uri));
        var resolved = options?.Clone() ?? new WordImageOptions();
        OfficeRemoteImage image = await OfficeRemoteImageLoader.LoadAsync(uri.AbsoluteUri, remoteImageOptions, cancellationToken).ConfigureAwait(false);
        using var stream = new MemoryStream(image.ToBytes());
        return AddParagraph().AddImage(stream, image.FileName, width, height, resolved);
    }
}
