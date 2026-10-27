using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Omd = OfficeIMO.Markdown;

namespace OfficeIMO.Word.Markdown {
    internal partial class MarkdownToWordConverter {
        private int _activeBlockRenderDepth;

        private void EnterBlockRender(WordMarkdownImportOptions options) {
            if (options.MaxNestingDepth < 1) {
                throw new ArgumentOutOfRangeException(
                    nameof(options.MaxNestingDepth),
                    options.MaxNestingDepth,
                    "MaxNestingDepth must be greater than zero.");
            }
            if (_activeBlockRenderDepth >= options.MaxNestingDepth) {
                throw new System.IO.InvalidDataException(
                    $"Markdown-to-Word block nesting exceeds MaxNestingDepth ({options.MaxNestingDepth}).");
            }

            _activeBlockRenderDepth++;
        }

        private void ExitBlockRender() {
            _activeBlockRenderDepth--;
        }

        private void RenderSharedBlockOmd(
            Omd.IMarkdownBlock block,
            IWordBlockRenderHost host,
            WordMarkdownImportOptions options,
            WordDocument document,
            WordList? currentList = null,
            int listLevel = 0,
            int quoteDepth = 0,
            double pageContentWidthPixels = 0,
            Omd.MarkdownColumnAlignment alignment = Omd.MarkdownColumnAlignment.None) {
            new BlockRenderer(this, host, options, document, listLevel, quoteDepth, pageContentWidthPixels, alignment)
                .Render(block);
        }

        private void ProcessBlockOmd(
            Omd.IMarkdownBlock block,
            WordDocument document,
            WordMarkdownImportOptions options,
            WordList? currentList = null,
            int listLevel = 0,
            int quoteDepth = 0,
            double pageContentWidthPixels = 0) {
            RenderSharedBlockOmd(
                block,
                new DocumentWordBlockRenderHost(document, options),
                options,
                document,
                currentList,
                listLevel,
                quoteDepth,
                pageContentWidthPixels,
                Omd.MarkdownColumnAlignment.None);
        }
    }
}
