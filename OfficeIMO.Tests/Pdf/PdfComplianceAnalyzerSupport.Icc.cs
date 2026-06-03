using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfComplianceAnalyzerTests {
    private static byte[] CreateMinimalIccProfile(string colorSpace = "RGB ") {
        byte[] profile = new byte[132];
        profile[3] = 132;
        Encoding.ASCII.GetBytes(colorSpace, 0, 4, profile, 16);
        profile[36] = (byte)'a';
        profile[37] = (byte)'c';
        profile[38] = (byte)'s';
        profile[39] = (byte)'p';
        return profile;
    }


}
