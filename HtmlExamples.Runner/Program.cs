using System;
using System.IO;

namespace HtmlExamples.Runner
{
    internal static class Program
    {
        private static void Setup(string path)
        {
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
        }

        private static void Main(string[] args)
        {
            string baseFolder = Path.TrimEndingDirectorySeparator(AppContext.BaseDirectory);
            Directory.SetCurrentDirectory(baseFolder);
            string folderPath = Path.Combine(baseFolder, "Documents");
            Setup(folderPath);

            // Run only HTML examples to isolate converter issues
            OfficeIMO.Examples.Html.Html.Example_Html01_LoadAndRoundTripBasics(folderPath, false);
            OfficeIMO.Examples.Html.Html.Example_Html02_SaveAsHtmlFromWord(folderPath, false);
            OfficeIMO.Examples.Html.Html.Example_Html03_TextFormatting(folderPath, false);
            OfficeIMO.Examples.Html.Html.Example_Html04_ListsAndNumbering(folderPath, false);
            OfficeIMO.Examples.Html.Html.Example_Html05_TablesComplex(folderPath, false);
            OfficeIMO.Examples.Html.Html.Example_Html06_ImagesAllModes(folderPath, false);
            OfficeIMO.Examples.Html.Html.Example_Html07_LinksAndAnchors(folderPath, false);
            OfficeIMO.Examples.Html.Html.Example_Html08_SemanticsAndCitations(folderPath, false);
            OfficeIMO.Examples.Html.Html.Example_Html09_CodePreWhitespace(folderPath, false);
            OfficeIMO.Examples.Html.Html.Example_Html10_OptionsAndAsync(folderPath, false).GetAwaiter().GetResult();
            OfficeIMO.Examples.Html.Html.Example_Html00_AllInOne(folderPath, false);
        }
    }
}

