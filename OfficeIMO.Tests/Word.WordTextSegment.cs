using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_WordTextSegmentConstructorsAndProperties() {
            var defaultSegment = new WordTextSegment();
            Assert.Equal(0, defaultSegment.BeginIndex);
            Assert.Equal(0, defaultSegment.EndIndex);
            Assert.Equal(0, defaultSegment.BeginChar);
            Assert.Equal(0, defaultSegment.EndChar);

            var segment = new WordTextSegment(1, 2, 3, 4, 5, 6);
            Assert.Equal(1, segment.BeginIndex);
            Assert.Equal(2, segment.EndIndex);
            Assert.Equal(5, segment.BeginChar);
            Assert.Equal(6, segment.EndChar);

            var begin = new WordPositionInParagraph(7, 8, 9);
            var end = new WordPositionInParagraph(10, 11, 12);
            var segment2 = new WordTextSegment(begin, end);
            Assert.Equal(7, segment2.BeginIndex);
            Assert.Equal(10, segment2.EndIndex);
            Assert.Equal(9, segment2.BeginChar);
            Assert.Equal(12, segment2.EndChar);

            segment2.BeginIndex = 13;
            segment2.EndIndex = 14;
            segment2.BeginChar = 15;
            segment2.EndChar = 16;

            Assert.Equal(13, segment2.BeginIndex);
            Assert.Equal(14, segment2.EndIndex);
            Assert.Equal(15, segment2.BeginChar);
            Assert.Equal(16, segment2.EndChar);
        }
    }
}
