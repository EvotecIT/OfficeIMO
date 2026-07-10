using OfficeIMO.Excel;
using OfficeIMO.Markdown;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Threading;

namespace OfficeIMO.Reader;

public static partial class DocumentReader {
    private sealed class RegistrarCandidate {
        public RegistrarCandidate(MethodInfo method, ReaderHandlerRegistrarDescriptor descriptor) {
            Method = method ?? throw new ArgumentNullException(nameof(method));
            Descriptor = descriptor ?? throw new ArgumentNullException(nameof(descriptor));
        }

        public MethodInfo Method { get; }
        public ReaderHandlerRegistrarDescriptor Descriptor { get; }
    }

    private sealed class FolderIngestState {
        public int FilesScanned { get; set; }
        public int FilesParsed { get; set; }
        public int FilesSkipped { get; set; }
        public long BytesRead { get; set; }
        public int ChunksProduced { get; set; }
    }

    private sealed class MarkdownChunkBlock {
        public MarkdownChunkBlock(int blockIndex, int startLine, int endLine, int sourceStartLine, int sourceEndLine, string? headingPath, string? headingSlug, string blockKind, string blockAnchor, string markdown, bool startsHeading, IReadOnlyList<ReaderTable> tables, IReadOnlyList<ReaderVisual> visuals) {
            BlockIndex = blockIndex;
            StartLine = startLine;
            EndLine = endLine;
            SourceStartLine = sourceStartLine;
            SourceEndLine = sourceEndLine;
            HeadingPath = headingPath;
            HeadingSlug = headingSlug;
            BlockKind = string.IsNullOrWhiteSpace(blockKind) ? "unknown" : blockKind;
            BlockAnchor = string.IsNullOrWhiteSpace(blockAnchor) ? "block-" + blockIndex.ToString(CultureInfo.InvariantCulture) : blockAnchor;
            Markdown = markdown ?? string.Empty;
            StartsHeading = startsHeading;
            Tables = tables ?? Array.Empty<ReaderTable>();
            Visuals = visuals ?? Array.Empty<ReaderVisual>();
        }

        public int BlockIndex { get; }
        public int StartLine { get; }
        public int EndLine { get; }
        public int SourceStartLine { get; }
        public int SourceEndLine { get; }
        public string? HeadingPath { get; }
        public string? HeadingSlug { get; }
        public string BlockKind { get; }
        public string BlockAnchor { get; }
        public string Markdown { get; }
        public bool StartsHeading { get; }
        public IReadOnlyList<ReaderTable> Tables { get; }
        public IReadOnlyList<ReaderVisual> Visuals { get; }
    }

    private sealed class MarkdownHeadingState {
        public MarkdownHeadingState(int level, string text, string slug) {
            Level = level;
            Text = text ?? string.Empty;
            Slug = slug ?? string.Empty;
        }

        public int Level { get; }
        public string Text { get; }
        public string Slug { get; }
    }

    private sealed class SourceInfo {
        public string Path { get; set; } = string.Empty;
        public string SourceId { get; set; } = string.Empty;
        public string? SourceHash { get; set; }
        public DateTime? LastWriteUtc { get; set; }
        public long? LengthBytes { get; set; }
    }
}
