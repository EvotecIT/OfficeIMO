using DocumentFormat.OpenXml.Drawing.Charts;
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
            DocumentRoot.Body ?? throw new InvalidOperationException("Document body is missing.");
    }

    public partial class WordChart {
        private ChartPart ChartPartRoot =>
            _chartPart ?? throw new InvalidOperationException("ChartPart is missing.");

        private ChartSpace ChartSpaceRoot =>
            ChartPartRoot.ChartSpace ?? throw new InvalidOperationException("ChartSpace is missing.");
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
