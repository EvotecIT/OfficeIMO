using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using OfficeIMO.Internal;
using OfficeIMO.Markup;

namespace OfficeIMO.Workflows;

internal static partial class MarkdownPowerPointProjection {
    internal static async Task PrepareImagesAsync(IList<OfficeMarkupBlock> blocks, MarkdownToPowerPointOptions options,
        string temporaryRoot, IDictionary<string, string> downloads, List<OfficeDiagnostic> findings, CancellationToken token) {
        for (int index = 0; index < blocks.Count; index++) {
            token.ThrowIfCancellationRequested();
            OfficeMarkupBlock block = blocks[index];
            if (block is OfficeMarkupImageBlock image) {
                try {
                    string path = image.Source;
                    if (path.StartsWith("data:", StringComparison.OrdinalIgnoreCase)) {
                        if (!options.AllowDataUriImages) throw new InvalidOperationException("Data URI images are disabled.");
                        continue; // The native renderer enforces the decoded-byte and encoding limits.
                    }
                    Uri.TryCreate(path, UriKind.Absolute, out var uri);
                    if (uri == null && Uri.TryCreate(options.ReaderOptions?.BaseUri, UriKind.Absolute, out var baseUri) &&
                        baseUri.Scheme is "http" or "https") Uri.TryCreate(baseUri, path, out uri);
                    if (uri != null && !uri.IsFile) {
                        if (!options.AllowRemoteImages || uri.Scheme is not ("http" or "https"))
                            throw new InvalidOperationException("Remote image retrieval is disabled or its scheme is unsupported.");
                        if (!downloads.TryGetValue(path, out string? captured)) {
                            OfficeRemoteImage loaded = await OfficeRemoteImageLoader.LoadAsync(uri, options.RemoteImageLoadOptions, token).ConfigureAwait(false);
                            captured = Path.Combine(temporaryRoot, Guid.NewGuid().ToString("N") + Path.GetExtension(loaded.FileName));
                            using (var output = OfficeTemporaryFile.CreateAtPath(captured, 81920, FileOptions.Asynchronous))
                                await OfficeStreamWriter.WriteAllBytesAsync(output, loaded.ToBytes(), token).ConfigureAwait(false);
                            downloads.Add(path, captured);
                        }
                        path = captured!;
                    } else {
                        if (!options.AllowLocalImages) throw new InvalidOperationException("Local images are disabled.");
                        path = uri?.IsFile == true ? uri.LocalPath : path;
                        path = Path.GetFullPath(Path.IsPathRooted(path) ? path : Path.Combine(options.BaseDirectory
                            ?? throw new InvalidOperationException("Relative images require a base directory."), path));
                        IEnumerable<string> roots = options.AllowedImageDirectories.Count > 0 ? options.AllowedImageDirectories
                            : options.BaseDirectory == null ? Array.Empty<string>() : new[] { options.BaseDirectory };
                        if (!roots.Any(root => OfficePathIdentity.IsSameOrDescendant(path, root)))
                            throw new InvalidOperationException("The image is outside the allowed directories.");
                        using var snapshot = await OfficeStreamFileSnapshot.CaptureAsync(_ => Task.FromResult<Stream>(
                            OfficePathIdentity.OpenRegularFileForRead(path, Path.GetDirectoryName(OfficePathIdentity.ResolvePhysicalPath(path))!, 81920)),
                            ".data", 64L * 1024 * 1024, null, token).ConfigureAwait(false);
                        string captured = Path.Combine(temporaryRoot, Guid.NewGuid().ToString("N") + Path.GetExtension(path));
                        File.Copy(snapshot.FilePath, captured);
                        path = captured;
                    }
                    var replacement = new OfficeMarkupImageBlock(path, image.Alt, image.Title, image.Width, image.Height) {
                        Caption = image.Caption, Placement = image.Placement, SourceText = image.SourceText
                    };
                    foreach (var attribute in image.Attributes) replacement.Attributes.Add(attribute.Key, attribute.Value);
                    blocks[index] = replacement;
                } catch (Exception error) when (error is not OperationCanceledException) {
                    findings.Add(new OfficeDiagnostic(OfficeDiagnosticCodes.ImageSkippedByPolicy, OfficeDiagnosticSeverity.Warning,
                        error.Message, image.Source, "Markdown", OfficeConversionLossKind.Omission));
                    blocks[index] = new OfficeMarkupParagraphBlock(image.Alt ?? image.Title ?? "Image omitted");
                }
            } else {
                IEnumerable<IList<OfficeMarkupBlock>> children = block switch {
                    OfficeMarkupSlideBlock slide => new[] { slide.Blocks },
                    OfficeMarkupSectionBlock section => new[] { section.Blocks },
                    OfficeMarkupColumnBlock column => new[] { column.Blocks },
                    OfficeMarkupListBlock list => list.Items.Select(item => item.Blocks),
                    _ => Array.Empty<IList<OfficeMarkupBlock>>()
                };
                foreach (IList<OfficeMarkupBlock> child in children)
                    await PrepareImagesAsync(child, options, temporaryRoot, downloads, findings, token).ConfigureAwait(false);
            }
        }
    }
}
