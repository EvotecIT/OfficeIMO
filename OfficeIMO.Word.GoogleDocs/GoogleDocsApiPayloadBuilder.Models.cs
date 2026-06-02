using OfficeIMO.GoogleWorkspace;
using System.Globalization;
using System.IO;
using System.Text;

namespace OfficeIMO.Word.GoogleDocs {
    internal static partial class GoogleDocsApiPayloadBuilder {
        private sealed class MaterializedParagraph {
            public string PrefixText { get; set; } = string.Empty;
            public StringBuilder TextBuilder { get; } = new StringBuilder();
            public List<MaterializedRun> Runs { get; } = new List<MaterializedRun>();
            public List<MaterializedFootnote> Footnotes { get; } = new List<MaterializedFootnote>();
            public List<MaterializedImage> Images { get; } = new List<MaterializedImage>();
            public string InsertedText => PrefixText + TextBuilder.ToString();
        }

        private sealed class MaterializedRun {
            public int StartOffset { get; set; }
            public int EndOffset { get; set; }
            public GoogleDocsParagraphRun Source { get; set; } = new GoogleDocsParagraphRun();
        }

        private sealed class MaterializedImage {
            public int InsertOffset { get; set; }
            public string Uri { get; set; } = string.Empty;
            public GoogleDocsInlineImage Source { get; set; } = new GoogleDocsInlineImage();
        }

        internal sealed class PreparedInitialBatchUpdate {
            public GoogleDocsApiBatchUpdatePayload Payload { get; set; } = new GoogleDocsApiBatchUpdatePayload();
            public List<GoogleDocsFootnote> Footnotes { get; } = new List<GoogleDocsFootnote>();
        }

        internal sealed class PreparedTableContentBatchUpdate {
            public GoogleDocsApiBatchUpdatePayload Payload { get; set; } = new GoogleDocsApiBatchUpdatePayload();
            public List<GoogleDocsFootnote> Footnotes { get; } = new List<GoogleDocsFootnote>();
        }

        private sealed class MaterializedFootnote {
            public int InsertOffset { get; set; }
            public GoogleDocsFootnote Source { get; set; } = new GoogleDocsFootnote();
        }

        private sealed class LiveTableContext {
            public int StartIndex { get; set; }
            public GoogleDocsApiTableResponse Table { get; set; } = new GoogleDocsApiTableResponse();
        }
    }
}
