using System.IO;
using System.Text;
using System.Threading;

namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    internal List<string> GetUnsupportedContentStreamFilters(CancellationToken cancellationToken = default) =>
        InspectContentStreams(recoverPrefix: false, cancellationToken, out _);

    internal List<string> InspectContentStreams(bool recoverPrefix, CancellationToken cancellationToken,
        out PdfReadPage? inspectionPage) {
        cancellationToken.ThrowIfCancellationRequested();
        inspectionPage = this;
        var unsupported = new List<string>();
        var pageResources = ResolveDictionary(GetInheritedValue("Resources"));
        var pageContentBudget = new PageContentBudget(this, cancellationToken);
        bool canInspectFormInvocations = MayContainFormXObjects(pageResources);
        StringBuilder? content = canInspectFormInvocations ? new StringBuilder() : null;
        PdfArray? recoveryArray = recoverPrefix && _pageDict.Items.TryGetValue("Contents", out PdfObject? contents)
            ? ResolveArray(contents) : null;
        List<PageContentStreamEntry>? streams = recoveryArray is null ? GetContentStreamObjects() : null;
        int streamCount = recoveryArray?.Items.Count ?? streams!.Count;
        bool resolvedPrefix = true;
        for (int index = 0; index < streamCount; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            PageContentStreamEntry entry;
            if (recoveryArray is null) {
                entry = streams![index];
            } else if (!TryResolvePageContentStream(recoveryArray.Items[index], out entry)) {
                resolvedPrefix = false;
                continue;
            }
            PdfStream stream = entry.Stream;
            if (AddUnsupportedFilters(stream, unsupported)) {
                canInspectFormInvocations = false;
                continue;
            }

            try {
                if (content is null) {
                    pageContentBudget.Validate(stream);
                } else {
                    byte[] decoded = pageContentBudget.Decode(stream);
                    if (content.Length > 0) content.Append(PdfContentStreamSequenceDecoder.Separator);
                    content.Append(PdfEncoding.Latin1GetString(decoded));
                }
            } catch (InvalidDataException) when (recoveryArray is not null && resolvedPrefix &&
                !stream.DecodingFailed) {
                cancellationToken.ThrowIfCancellationRequested();
                if (!Filters.StreamDecoder.TryGetDeclaredFilterNames(stream.Dictionary, _objects, out var filterNames) ||
                    filterNames.Count == 0) throw;
                // A content array is one operator sequence. Later streams can depend on state lost here.
                inspectionPage = index == 0 ? null : WithInspectionContentPrefix(recoveryArray, index, cancellationToken);
                break;
            }
        }

        if (canInspectFormInvocations && content is { Length: > 0 }) {
            CollectUnsupportedFormFilters(content.ToString(), pageResources, unsupported, new HashSet<PdfStream>(), pageContentBudget);
        }
        return unsupported;
    }

    private PdfReadPage WithInspectionContentPrefix(PdfArray contents, int count, CancellationToken cancellationToken) {
        var page = new PdfDictionary();
        foreach (KeyValuePair<string, PdfObject> item in _pageDict.Items) {
            cancellationToken.ThrowIfCancellationRequested();
            page.Items.Add(item.Key, item.Value);
        }
        var prefix = new PdfArray(count);
        for (int index = 0; index < count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            prefix.Items.Add(contents.Items[index]);
        }
        page.Items["Contents"] = prefix;
        return new PdfReadPage(ObjectNumber, page, _objects, _limits, _fontResourceCache,
            _demandTextExtraction, _demandContentExtraction, _includeArtifactText,
            _outputIntentColorTransform, _optionalContentVisibilityState, _preferToUnicode,
            _useWinAnsiFallback, _adjustKerningFromTJ);
    }
}
