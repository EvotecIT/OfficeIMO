using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word;

/// <summary>
/// Examples demonstrating document cleanup operations.
/// </summary>
internal static partial class CleanupDocuments {
    /// <summary>
    /// Loads a template and performs cleanup of redundant runs.
    /// </summary>
    /// <param name="openWord">Opens Word when <c>true</c>.</param>
    public static void CleanupDocuments_Sample01(bool openWord) {
        Console.WriteLine("[*] Load external Word Document - Sample 1");
        string documentPaths = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Templates");
        string fullPath = System.IO.Path.Combine(documentPaths, "sample1.docx");
        using (WordDocument document = WordDocument.Load(fullPath, false)) {
            Console.WriteLine(fullPath);
            Console.WriteLine("Sections count: " + document.Sections.Count);
            Console.WriteLine("Tables count: " + document.Tables.Count);
            Console.WriteLine("Paragraphs count: " + document.Paragraphs.Count);

            var cleanupCount = document.CleanupDocument();

            Console.WriteLine("Removed " + cleanupCount + " runs because of identical formatting.");

            Console.WriteLine("Paragraphs count: " + document.Paragraphs.Count);

            document.Save(openWord);
        }
    }
}
