using System.IO.Packaging;
using DocumentFormat.OpenXml.Features;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.OpenXml.Internal {
    /// <summary>
    /// Opens Word, Excel and PowerPoint packages for reading over the caller's stream.
    /// </summary>
    /// <remarks>
    /// The Open XML SDK copies every stream it opens for reading to a temporary file and reopens the package from that
    /// copy, so that it can rewrite external relationship targets that are not valid URIs. A package whose
    /// relationship targets all parse needs no rewriting and reads the same without the copy. This opener opens the
    /// package in place and reads every relationship straight away; a package that does not open or read cleanly
    /// that way goes through the SDK's stream overload, which repairs or reports it as before.
    /// </remarks>
    internal static class OfficeOpenXmlReadOnlyPackage {
        internal static SpreadsheetDocument OpenSpreadsheet(Stream stream, OpenSettings settings) =>
            Open(stream, settings, static (source, openSettings) => SpreadsheetDocument.Open(source, false, openSettings),
                SpreadsheetDocument.Open);

        internal static WordprocessingDocument OpenWordprocessing(Stream stream, OpenSettings settings) =>
            Open(stream, settings, static (source, openSettings) => WordprocessingDocument.Open(source, false, openSettings),
                WordprocessingDocument.Open);

        internal static PresentationDocument OpenPresentation(Stream stream, OpenSettings settings) =>
            Open(stream, settings, static (source, openSettings) => PresentationDocument.Open(source, false, openSettings),
                PresentationDocument.Open);

        private static TDocument Open<TDocument>(Stream stream, OpenSettings settings,
            Func<Stream, OpenSettings, TDocument> openStream, Func<Package, OpenSettings, TDocument> openPackage)
            where TDocument : OpenXmlPackage {
            if (!stream.CanRead || !stream.CanSeek) return openStream(stream, settings);
            long start = stream.Position;
            Package? package = null;
            TDocument? document = null;
            try {
                package = Package.Open(stream, FileMode.Open, FileAccess.Read);
                document = openPackage(package, settings);
                // Loads every reachable part and its relationships, as the SDK does before it loads the first part
                // root anyway. A target System.IO.Packaging cannot parse throws here, while the repairing route is
                // still available, instead of on some later read.
                _ = document.StrictRelationshipFound;
                if (document.Features.Get<IDisposableFeature>() is { } disposables) {
                    // The SDK leaves a package it was handed open; closing the document closes it too.
                    disposables.Register(package);
                    return document;
                }
            } catch (Exception exception) when (exception is not OutOfMemoryException) {
                // The SDK's stream route decides how to open, repair or report this package, as it always has.
            }

            document?.Dispose();
            package?.Close();
            stream.Position = start;
            return openStream(stream, settings);
        }
    }
}
