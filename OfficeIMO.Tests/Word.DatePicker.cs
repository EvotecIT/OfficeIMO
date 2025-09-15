using System;
using System.IO;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_AddingDatePicker() {
            string filePath = Path.Combine(_directoryWithFiles, "DocumentWithDatePicker.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var dp = document.AddParagraph("Date:").AddDatePicker(DateTime.Today, "DP", "DPTag");

                Assert.Single(document.DatePickers);
                Assert.Equal(DateTime.Today.Date, dp.Date?.Date);
                Assert.Equal("DPTag", dp.Tag);
                Assert.Equal("DP", dp.Alias);

                var picker = document.GetDatePickerByTag("DPTag");
                Assert.NotNull(picker);
                Assert.Equal("DP", document.GetDatePickerByAlias("DP")?.Alias);
                picker!.Date = DateTime.Today.AddDays(1);
                Assert.Equal(DateTime.Today.AddDays(1).Date, document.DatePickers[0].Date?.Date);
                document.DatePickers[0].Remove();
                Assert.Empty(document.DatePickers);
            }
        }
    }
}
