using System;
using System.Text;
using OfficeIMO.Word;

internal static partial class Protect {
    internal static void Debug_TestPasswordHashes() {
        Console.WriteLine("[*] Testing password hash calculations for 'Test123':");
        
        // Test 1: Create documents with known passwords and inspect what's generated
        string folderPath = System.IO.Path.GetTempPath();
        
        // Test enforced protection
        string filePath1 = System.IO.Path.Combine(folderPath, "Debug_Enforced.docx");
        using (WordDocument document = WordDocument.Create(filePath1)) {
            document.AddParagraph("Debug enforced protection");
            document.Settings.ProtectionPassword = "Test123";
            document.Settings.ProtectionType = DocumentFormat.OpenXml.Wordprocessing.DocumentProtectionValues.ReadOnly;
            document.Save(false);
        }
        
        // Test read-only recommendation  
        string filePath2 = System.IO.Path.Combine(folderPath, "Debug_ReadOnly.docx");
        using (WordDocument document = WordDocument.Create(filePath2)) {
            document.AddParagraph("Debug read-only recommendation");
            document.Settings.ReadOnlyPassword = "Test123";
            document.Settings.ReadOnlyRecommended = true;
            document.Save(false);
        }
        
        Console.WriteLine("Created debug documents:");
        Console.WriteLine($"- Enforced: {filePath1}");
        Console.WriteLine($"- ReadOnly: {filePath2}");
        Console.WriteLine("Check these in Word with password 'Test123'");
    }
}
