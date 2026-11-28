using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        internal MainDocumentPart MainDocumentPartRoot =>
            _wordprocessingDocument?.MainDocumentPart ?? throw new InvalidOperationException("MainDocumentPart is missing.");

        internal Document DocumentRoot {
            get => MainDocumentPartRoot.Document ?? throw new InvalidOperationException("Document is missing.");
            set => MainDocumentPartRoot.Document = value;
        }

        internal Body BodyRoot =>
            TryGetBodyRoot() ?? DocumentRoot.Body ?? throw new InvalidOperationException("Document body is missing.");

        // Finding the main part and its body takes several lookups, which paragraph-by-paragraph authoring repeats.
        // The cached body is used only while it is still the body of the main part's loaded root.
        private WordprocessingDocument? _bodyCachePackage;
        private MainDocumentPart? _bodyCachePart;
        private Document? _bodyCacheDocument;
        private Body? _bodyCache;

        /// <summary>The main document body, or <see langword="null"/> when the package has none.</summary>
        internal Body? TryGetBodyRoot() {
            Body? body = _bodyCache;
            if (body != null
                && ReferenceEquals(_bodyCachePackage, _wordprocessingDocument)
                && _bodyCachePart is { IsRootElementLoaded: true } cachedPart
                && ReferenceEquals(cachedPart.RootElement, _bodyCacheDocument)
                && ReferenceEquals(body.Parent, _bodyCacheDocument)) {
                return body;
            }

            WordprocessingDocument? package = _wordprocessingDocument;
            MainDocumentPart? part = package?.MainDocumentPart;
            Document? document = part?.Document;
            body = document?.Body;
            if (body == null) {
                _bodyCache = null;
                return null;
            }

            _bodyCachePackage = package;
            _bodyCachePart = part;
            _bodyCacheDocument = document;
            _bodyCache = body;
            return body;
        }
    }

    public partial class WordList {
        private MainDocumentPart MainDocumentPartRoot => _document.MainDocumentPartRoot;

        private NumberingDefinitionsPart EnsureNumberingDefinitionsPartRoot() =>
            MainDocumentPartRoot.NumberingDefinitionsPart ?? MainDocumentPartRoot.AddNewPart<NumberingDefinitionsPart>();

        private Numbering EnsureNumberingRoot() =>
            EnsureNumberingDefinitionsPartRoot().Numbering ??= new Numbering();

        private Numbering? TryGetNumberingRoot() =>
            _document._wordprocessingDocument?.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
    }

    public partial class WordTable {
        private Body BodyRoot => _document.BodyRoot;
    }
}
