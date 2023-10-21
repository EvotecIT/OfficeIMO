using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;
using Color = SixLabors.ImageSharp.Color;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_CreatingDocumentWithSettings() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatingDocumentWithSettings.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {

                Assert.True(document.CompatibilitySettings.CompatibilityMode == CompatibilityMode.Word2013);

                document.Settings.ProtectionPassword = "Test";

                Assert.True(document.Settings.ProtectionType == DocumentProtectionValues.ReadOnly);

                Assert.True(document.Settings.Language == "en-US");

                document.Settings.Language = "pl-PL";

                Assert.True(document.Settings.Language == "pl-PL");

                document.Settings.SetBackgroundColor(Color.BlueViolet);

                Assert.True(document.Settings.BackgroundColor == "8A2BE2");

                document.Settings.ZoomPercentage = 150;

                Assert.True(document.Settings.ZoomPercentage == 150);

                Assert.True(document.Settings.UpdateFieldsOnOpen == false);

                document.Settings.UpdateFieldsOnOpen = true;

                Assert.True(document.Settings.UpdateFieldsOnOpen == true);

                Assert.True(document.Settings.FontSize == 11); // default value

                document.Settings.FontSize = 30;

                Assert.True(document.Settings.FontSize == 30);

                Assert.True(document.Settings.FontSizeComplexScript == 11);

                document.Settings.FontSizeComplexScript = 20;

                Assert.True(document.Settings.FontSizeComplexScript == 20);

                // those are default values
                Assert.True(document.Settings.FontFamily == null);
                Assert.True(document.Settings.FontFamilyHighAnsi == null);

                document.Settings.FontFamily = "Courier New";

                Assert.True(document.Settings.FontFamily == "Courier New");
                Assert.True(document.Settings.FontFamilyEastAsia == "Courier New");
                Assert.True(document.Settings.FontFamilyComplexScript == "Courier New");
                Assert.True(document.Settings.FontFamilyHighAnsi == "Courier New");

                document.Settings.FontFamilyHighAnsi = "Arial";
                Assert.True(document.Settings.FontFamilyHighAnsi == "Arial");
                Assert.True(document.Settings.FontFamily == "Courier New");
                Assert.True(document.Settings.FontFamilyEastAsia == "Courier New");
                Assert.True(document.Settings.FontFamilyComplexScript == "Courier New");

                document.Settings.FontFamily = "Times New Roman";
                Assert.True(document.Settings.FontFamily == "Times New Roman");
                Assert.True(document.Settings.FontFamilyHighAnsi == "Times New Roman");
                Assert.True(document.Settings.FontFamilyEastAsia == "Times New Roman");
                Assert.True(document.Settings.FontFamilyComplexScript == "Times New Roman");

                document.Settings.FontFamilyHighAnsi = "Arial";
                Assert.True(document.Settings.FontFamilyHighAnsi == "Arial");
                Assert.True(document.Settings.FontFamily == "Times New Roman");
                Assert.True(document.Settings.FontFamilyEastAsia == "Times New Roman");
                Assert.True(document.Settings.FontFamilyComplexScript == "Times New Roman");

                document.Settings.FontFamilyEastAsia = null;
                Assert.True(document.Settings.FontFamilyEastAsia == null);

                document.Settings.FontFamilyComplexScript = null;
                Assert.True(document.Settings.FontFamilyComplexScript == null);

                document.CompatibilitySettings.CompatibilityMode = CompatibilityMode.Word2003;
                Assert.True(document.CompatibilitySettings.CompatibilityMode == CompatibilityMode.Word2003);

                Assert.True(document.Settings.ReadOnlyRecommended == false);
                document.Settings.ReadOnlyRecommended = true;
                Assert.True(document.Settings.ReadOnlyRecommended == true);

                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatingDocumentWithSettings.docx"))) {
                Assert.True(document.CompatibilitySettings.CompatibilityMode == CompatibilityMode.Word2003);
                document.CompatibilitySettings.CompatibilityMode = CompatibilityMode.Word2007;
                Assert.True(document.CompatibilitySettings.CompatibilityMode == CompatibilityMode.Word2007);

                Assert.True(document.Settings.FontFamilyHighAnsi == "Arial");
                Assert.True(document.Settings.Language == "pl-PL");

                document.Settings.Language = "en-US";

                Assert.True(document.Settings.Language == "en-US");

                Assert.True(document.Settings.ProtectionType == DocumentProtectionValues.ReadOnly);

                document.Settings.RemoveProtection();

                Assert.True(document.Settings.BackgroundColor == "8A2BE2");

                document.Settings.SetBackgroundColor("FFA07A");

                Assert.True(document.Settings.ZoomPercentage == 150);

                document.Settings.ZoomPercentage = 100;

                Assert.True(document.Settings.ZoomPercentage == 100);

                Assert.True(document.Settings.UpdateFieldsOnOpen == true);

                document.Settings.UpdateFieldsOnOpen = false;

                Assert.True(document.Settings.FontSizeComplexScript == 20);
                Assert.True(document.Settings.FontSize == 30);
                Assert.True(document.Settings.FontFamily == "Times New Roman");

                document.Settings.FontFamilyHighAnsi = "Abadi";

                Assert.True(document.Settings.ReadOnlyRecommended == true);
                document.Settings.ReadOnlyRecommended = false;
                Assert.True(document.Settings.ReadOnlyRecommended == false);

                document.Save();
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatingDocumentWithSettings.docx"))) {
                Assert.True(document.CompatibilitySettings.CompatibilityMode == CompatibilityMode.Word2007);
                document.CompatibilitySettings.CompatibilityMode = CompatibilityMode.Word2010;
                Assert.True(document.CompatibilitySettings.CompatibilityMode == CompatibilityMode.Word2010);

                Assert.True(document.Settings.FontFamilyHighAnsi == "Abadi");
                Assert.True(document.Settings.FontFamily == "Times New Roman");

                Assert.True(document.Settings.ProtectionType == null);
                Assert.True(document.Settings.BackgroundColor == "FFA07A");
                Assert.True(document.Settings.ZoomPercentage == 100);
                Assert.True(document.Settings.UpdateFieldsOnOpen == false);

                Assert.True(document.Settings.ReadOnlyRecommended == false);

                Assert.True(document.Settings.FinalDocument == false);
                document.Settings.FinalDocument = true;
                Assert.True(document.Settings.FinalDocument == true);
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatingDocumentWithSettings.docx"))) {

                Assert.True(document.CompatibilitySettings.CompatibilityMode == CompatibilityMode.Word2010);
                document.CompatibilitySettings.CompatibilityMode = CompatibilityMode.None;
                Assert.True(document.CompatibilitySettings.CompatibilityMode == CompatibilityMode.None);

                document.Settings.ZoomPreset = PresetZoomValues.BestFit;

                Assert.True(document.Settings.ZoomPreset == PresetZoomValues.BestFit);

                document.Settings.ZoomPreset = PresetZoomValues.FullPage;

                Assert.True(document.Settings.ZoomPreset == PresetZoomValues.FullPage);

                document.Settings.ZoomPreset = PresetZoomValues.None;

                Assert.True(document.Settings.ZoomPreset == PresetZoomValues.None);

                document.Settings.ZoomPreset = PresetZoomValues.TextFit;

                Assert.True(document.Settings.ZoomPreset == PresetZoomValues.TextFit);

                Assert.True(document.Settings.FinalDocument == true);
                document.Settings.FinalDocument = false;
                Assert.True(document.Settings.FinalDocument == false);

                document.Save();
            }
        }
    }
}
