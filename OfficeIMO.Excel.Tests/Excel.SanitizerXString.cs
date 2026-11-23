using OfficeIMO.Excel.Utilities;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData("plain\nline", "plain\nline")]
        [InlineData("line\nthen _x0041_ token", "line\nthen _x005F_x0041_ token")]
        [InlineData("tab\tthen\u0001control", "tab\tthen_x0001_control")]
        [InlineData("pair 😀 kept", "pair 😀 kept")]
        [InlineData("non￾character", "non_xFFFE_character")]
        [InlineData("under_score_x12_", "under_score_x12_")]
        [InlineData("😀😀😀 then _x0041_", "😀😀😀 then _x005F_x0041_")]
        [InlineData("😀😀 then\u0002 control", "😀😀 then_x0002_ control")]
        public void ExcelSanitizer_EncodeXStringFindsEscapesAfterLineBreaksAndLongText(string value, string expected) {
            string padding = new string('n', 300);
            Assert.Equal(expected, ExcelSanitizer.EncodeXString(value));
            Assert.Equal(padding + "\n" + padding + expected, ExcelSanitizer.EncodeXString(padding + "\n" + padding + value));
        }

        [Fact]
        public void ExcelSanitizer_EncodeXStringEscapesLoneSurrogatesAfterLongText() {
            // Attribute arguments cannot carry lone surrogates, so these values are built here.
            string prefix = new string('n', 300) + "\n";
            Assert.Equal(prefix + "lone _xD83D_ high", ExcelSanitizer.EncodeXString(prefix + "lone " + (char)0xD83D + " high"));
            Assert.Equal(prefix + "lone _xDE00_ low", ExcelSanitizer.EncodeXString(prefix + "lone " + (char)0xDE00 + " low"));
            Assert.Equal(prefix + "😀😀 then _xD83D_", ExcelSanitizer.EncodeXString(prefix + "😀😀 then " + (char)0xD83D));
        }

        [Fact]
        public void ExcelSanitizer_EncodeXStringReturnsTheSameStringWhenNothingNeedsEscaping() {
            string value = new string('a', 2048) + " & <tag> Łódź\n" + new string('x', 2048);
            Assert.Same(value, ExcelSanitizer.EncodeXString(value));
            Assert.Same(value, ExcelSanitizer.SanitizeString(value, int.MaxValue));
        }
    }
}
