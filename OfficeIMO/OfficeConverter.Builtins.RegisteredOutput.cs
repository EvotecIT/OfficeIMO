namespace OfficeIMO;

// Optional format owners reopen their own artifacts without adding their engines to the front-door package.
internal interface IOfficeConversionArtifactValidation {
    Task<bool> TryValidateArtifactAsync(Stream output, OfficeConversionOptions options, CancellationToken cancellationToken);
}

public static partial class OfficeConverter {
    private static async Task ValidateRegisteredArtifactAsync(IOfficeConversionRoute route, Stream output,
        OfficeConversionOptions options, CancellationToken cancellationToken) {
        if (route is IOfficeConversionArtifactValidation owner &&
            await owner.TryValidateArtifactAsync(output, options, cancellationToken).ConfigureAwait(false)) return;
        await ValidateRegisteredOutputAsync(output, options, cancellationToken).ConfigureAwait(false);
    }
}
