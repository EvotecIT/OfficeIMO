using System;
using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_Source_Slice_Failure_Reason_Tests {
    [Fact]
    public void SourceMapping_Reports_Exact_Original_And_Normalized_Slices() {
        var result = OfficeIMO.Markdown.MarkdownReader.ParseWithSyntaxTree("# Title\n\nText\n", new MarkdownReaderOptions {
            PreserveTrivia = true
        });
        var heading = result.SyntaxTree.Children[0];

        var created = result.TryCreateSourceMapping(heading, out var mapping);

        Assert.True(created);
        Assert.Equal("# Title", mapping.NormalizedSourceSlice.Text);
        Assert.True(mapping.HasOriginalSource);
        Assert.Equal("# Title", mapping.OriginalSourceSlice.Text);
        Assert.Equal(MarkdownOriginalSourceMappingKind.Exact, mapping.OriginalSourceMappingKind);
        Assert.Equal(MarkdownOriginalSourceSliceFailureReason.None, mapping.OriginalSourceFailureReason);
    }

    [Fact]
    public void SourceMapping_Reports_LineEnding_Equivalent_Original_Slices_With_Original_Offsets() {
        var result = OfficeIMO.Markdown.MarkdownReader.ParseWithSyntaxTree("# Title\r\n\r\nText\r\n", new MarkdownReaderOptions {
            PreserveTrivia = true
        });
        var paragraph = result.SyntaxTree.Children[1];

        var created = result.TryCreateSourceMapping(paragraph, out var mapping);

        Assert.True(created);
        Assert.Equal("Text", mapping.NormalizedSourceSlice.Text);
        Assert.Equal(9, mapping.NormalizedSourceSlice.StartOffset);
        Assert.True(mapping.HasOriginalSource);
        Assert.Equal("Text", mapping.OriginalSourceSlice.Text);
        Assert.Equal(11, mapping.OriginalSourceSlice.StartOffset);
        Assert.Equal(MarkdownOriginalSourceMappingKind.LineEndingEquivalent, mapping.OriginalSourceMappingKind);
        Assert.Equal(MarkdownOriginalSourceSliceFailureReason.None, mapping.OriginalSourceFailureReason);
    }

    [Fact]
    public void SourceMapping_Keeps_Normalized_Slice_When_Original_Markdown_Was_Not_Preserved() {
        var result = OfficeIMO.Markdown.MarkdownReader.ParseWithSyntaxTree("# Title\n");
        var heading = result.SyntaxTree.Children[0];

        var created = result.TryCreateSourceMapping(heading, out var mapping);

        Assert.True(created);
        Assert.Equal("# Title", mapping.NormalizedSourceSlice.Text);
        Assert.False(mapping.HasOriginalSource);
        Assert.Equal(MarkdownOriginalSourceMappingKind.Unavailable, mapping.OriginalSourceMappingKind);
        Assert.Equal(MarkdownOriginalSourceSliceFailureReason.OriginalMarkdownNotPreserved, mapping.OriginalSourceFailureReason);
    }

    [Fact]
    public void SourceMapping_Keeps_Normalized_Slice_When_Original_Text_Is_Not_Equivalent() {
        var options = new MarkdownReaderOptions {
            PreserveTrivia = true,
            InputNormalization = new MarkdownInputNormalizationOptions {
                NormalizeZeroWidthSpacingArtifacts = true
            }
        };
        var result = OfficeIMO.Markdown.MarkdownReader.ParseWithSyntaxTree("# Ti\u200Btle\n", options);
        var heading = result.SyntaxTree.Children[0];

        var created = result.TryCreateSourceMapping(heading, out var mapping);

        Assert.True(created);
        Assert.Equal("# Title", mapping.NormalizedSourceSlice.Text);
        Assert.False(mapping.HasOriginalSource);
        Assert.Equal(MarkdownOriginalSourceMappingKind.Unavailable, mapping.OriginalSourceMappingKind);
        Assert.Equal(MarkdownOriginalSourceSliceFailureReason.OriginalTextNotEquivalent, mapping.OriginalSourceFailureReason);
    }

    [Fact]
    public void NativeSourceMapping_Reports_Generated_Node_Reason_While_Keeping_Normalized_Slice() {
        var native = MarkdownNativeDocument.Parse("text\n", new MarkdownReaderOptions {
            PreserveTrivia = true
        });
        var syntaxNode = new MarkdownSyntaxNode(
            MarkdownSyntaxKind.InlineText,
            new MarkdownSourceSpan(1, 1, 1, 4),
            literal: "text",
            isGenerated: true);
        var inline = new MarkdownNativeInline(
            MarkdownNativeInlineKind.Text,
            syntaxNode,
            Array.Empty<MarkdownNativeInline>(),
            Array.Empty<MarkdownNativeInlineMetadata>());

        var created = native.TryCreateSourceMapping(inline, out var mapping);

        Assert.True(created);
        Assert.Equal("text", mapping.NormalizedSourceSlice.Text);
        Assert.False(mapping.HasOriginalSource);
        Assert.Equal(MarkdownOriginalSourceMappingKind.Unavailable, mapping.OriginalSourceMappingKind);
        Assert.Equal(MarkdownOriginalSourceSliceFailureReason.GeneratedSyntaxNode, mapping.OriginalSourceFailureReason);
    }

    [Fact]
    public void OriginalSourceSlice_Returns_NotPreserved_Reason_When_Trivia_Is_Disabled() {
        var result = OfficeIMO.Markdown.MarkdownReader.ParseWithSyntaxTree("# Title\n");
        var heading = result.SyntaxTree.Children[0];

        var created = result.TryCreateOriginalSourceSlice(heading, out _, out var failureReason);

        Assert.False(created);
        Assert.Equal(MarkdownOriginalSourceSliceFailureReason.OriginalMarkdownNotPreserved, failureReason);
    }

    [Fact]
    public void OriginalSourceSlice_Returns_TextNotEquivalent_Reason_When_Input_Normalization_Changed_Text() {
        var options = new MarkdownReaderOptions {
            PreserveTrivia = true,
            InputNormalization = new MarkdownInputNormalizationOptions {
                NormalizeZeroWidthSpacingArtifacts = true
            }
        };
        var result = OfficeIMO.Markdown.MarkdownReader.ParseWithSyntaxTree("# Ti\u200Btle\n", options);
        var heading = result.SyntaxTree.Children[0];

        var created = result.TryCreateOriginalSourceSlice(heading, out _, out var failureReason);

        Assert.False(created);
        Assert.Equal(MarkdownOriginalSourceSliceFailureReason.OriginalTextNotEquivalent, failureReason);
    }

    [Fact]
    public void OriginalSourceSlice_Returns_AssociatedObjectNotFound_Reason_For_Untracked_Object() {
        var result = OfficeIMO.Markdown.MarkdownReader.ParseWithSyntaxTree("# Title\n", new MarkdownReaderOptions {
            PreserveTrivia = true
        });

        var created = result.TryCreateOriginalSourceSlice(new object(), out _, out var failureReason);

        Assert.False(created);
        Assert.Equal(MarkdownOriginalSourceSliceFailureReason.AssociatedObjectNotFound, failureReason);
    }

    [Fact]
    public void NativeDocument_OriginalSourceSlice_Returns_Field_Mapping_Failure_Reason() {
        var native = MarkdownNativeDocument.Parse("# Title\n");
        var heading = Assert.IsType<MarkdownNativeHeadingBlock>(Assert.Single(native.Blocks));
        var text = heading.FindSourceField("text");
        Assert.NotNull(text);

        var created = native.TryCreateOriginalSourceSlice(text!, out _, out var failureReason);

        Assert.False(created);
        Assert.Equal(MarkdownOriginalSourceSliceFailureReason.OriginalMarkdownNotPreserved, failureReason);
    }

    [Fact]
    public void NativeDocument_OriginalSourceSlice_Returns_Generated_Node_Reason_For_Generated_Inline() {
        var native = MarkdownNativeDocument.Parse("text\n", new MarkdownReaderOptions {
            PreserveTrivia = true
        });
        var syntaxNode = new MarkdownSyntaxNode(
            MarkdownSyntaxKind.InlineText,
            new MarkdownSourceSpan(1, 1, 1, 4),
            literal: "text",
            isGenerated: true);
        var inline = new MarkdownNativeInline(
            MarkdownNativeInlineKind.Text,
            syntaxNode,
            Array.Empty<MarkdownNativeInline>(),
            Array.Empty<MarkdownNativeInlineMetadata>());

        var created = native.TryCreateOriginalSourceSlice(inline, out _, out var failureReason);

        Assert.False(created);
        Assert.Equal(MarkdownOriginalSourceSliceFailureReason.GeneratedSyntaxNode, failureReason);
    }

    [Fact]
    public void NativeDocument_OriginalSourceSlice_Returns_Generated_Node_Reason_For_Generated_Inline_Metadata() {
        var native = MarkdownNativeDocument.Parse("[x](https://example.com)\n", new MarkdownReaderOptions {
            PreserveTrivia = true
        });
        var syntaxNode = new MarkdownSyntaxNode(
            MarkdownSyntaxKind.InlineLinkTarget,
            new MarkdownSourceSpan(1, 5, 1, 23),
            literal: "https://example.com",
            isGenerated: true);
        var metadata = new MarkdownNativeInlineMetadata("target", "https://example.com", syntaxNode);

        var created = native.TryCreateOriginalSourceSlice(metadata, out _, out var failureReason);

        Assert.False(created);
        Assert.Equal(MarkdownOriginalSourceSliceFailureReason.GeneratedSyntaxNode, failureReason);
    }

    [Fact]
    public void Roundtrip_SourceEdit_Fallback_Diagnostic_Includes_Original_Mapping_Reason() {
        var options = new MarkdownReaderOptions {
            PreserveTrivia = true,
            InputNormalization = new MarkdownInputNormalizationOptions {
                NormalizeZeroWidthSpacingArtifacts = true
            }
        };
        var native = MarkdownNativeDocument.Parse("# Ol\u200Bd\r\n", options);
        var heading = Assert.IsType<MarkdownNativeHeadingBlock>(Assert.Single(native.Blocks));
        var edit = native.CreateReplaceEdit(heading.TextSourceSpan!.Value, "New");

        Assert.Equal(MarkdownOriginalSourceSliceFailureReason.OriginalTextNotEquivalent, edit.OriginalSourceFailureReason);

        var roundtrip = native.WriteWithSourceEdit(edit);

        var diagnostic = Assert.Single(roundtrip.Diagnostics);
        Assert.Equal("roundtrip.original-source-slice-unavailable", diagnostic.Id);
        Assert.Equal(MarkdownOriginalSourceSliceFailureReason.OriginalTextNotEquivalent, diagnostic.OriginalSourceFailureReason);
        Assert.Contains("original reader input is not equivalent to normalized markdown", diagnostic.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void NativeSourceEdit_Carries_NotPreserved_Reason_Without_Duplicating_Roundtrip_Diagnostics() {
        var native = MarkdownNativeDocument.Parse("# Old\n");
        var heading = Assert.IsType<MarkdownNativeHeadingBlock>(Assert.Single(native.Blocks));
        var edit = native.CreateReplaceEdit(heading.TextSourceSpan!.Value, "New");

        Assert.Equal(MarkdownOriginalSourceSliceFailureReason.OriginalMarkdownNotPreserved, edit.OriginalSourceFailureReason);

        var roundtrip = native.WriteWithSourceEdit(edit);

        var diagnostic = Assert.Single(roundtrip.Diagnostics);
        Assert.Equal("roundtrip.preserve-trivia-required", diagnostic.Id);
        Assert.Equal(MarkdownOriginalSourceSliceFailureReason.OriginalMarkdownNotPreserved, diagnostic.OriginalSourceFailureReason);
        Assert.Contains("PreserveTrivia enabled", diagnostic.Message, StringComparison.Ordinal);
        Assert.Equal("# New\n", roundtrip.Markdown);
    }
}
