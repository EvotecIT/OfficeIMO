using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Diagnostics;
using OfficeIMO.Rtf.Syntax;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfLosslessRoundTripTests {
    [Fact]
    public void SyntaxTree_ToRtf_Preserves_Unknown_Destinations_Control_Spacing_And_Binary() {
        const string rtf = @"{\rtf1\ansi{\*\unknown\foo-12 bar \'80}{\object\objdata 0102}\pard Before \bin4 a{b} after\par}";

        RtfSyntaxTree tree = RtfSyntaxTree.Parse(rtf);
        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.DoesNotContain(tree.Diagnostics, diagnostic => diagnostic.Severity == RtfDiagnosticSeverity.Error);
        Assert.Equal(rtf, tree.ToRtf());
        Assert.Equal(rtf, result.ToRtfLossless());
    }

    [Fact]
    public void Load_And_SaveLossless_Preserve_Raw_Binary_File_Bytes() {
        string inputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".rtf");
        string outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".rtf");

        byte[] bytes = new byte[] {
            123, 92, 114, 116, 102, 49, 92, 97, 110, 115, 105, 92, 98, 105, 110, 49, 32, 0x80, 125
        };

        try {
            File.WriteAllBytes(inputPath, bytes);

            RtfReadResult result = RtfDocument.Load(inputPath);
            result.SaveLossless(outputPath);

            Assert.Equal(bytes, File.ReadAllBytes(outputPath));
            Assert.Equal(@"{\rtf1\ansi\bin1 " + (char)0x80 + "}", result.ToRtfLossless());
        } finally {
            if (File.Exists(inputPath)) File.Delete(inputPath);
            if (File.Exists(outputPath)) File.Delete(outputPath);
        }
    }

    [Fact]
    public void Load_And_ToBytesLossless_Preserve_Raw_Binary_Bytes() {
        byte[] bytes = new byte[] {
            123, 92, 114, 116, 102, 49, 92, 97, 110, 115, 105, 92, 98, 105, 110, 49, 32, 0x80, 125
        };

        RtfReadResult result = RtfDocument.Load(bytes);

        Assert.Equal(bytes, result.ToBytesLossless());
        Assert.Equal(@"{\rtf1\ansi\bin1 " + (char)0x80 + "}", result.ToRtfLossless());
    }
}
