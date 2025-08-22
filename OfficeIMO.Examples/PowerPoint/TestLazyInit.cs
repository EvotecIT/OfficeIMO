using System;
using System.IO;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.Examples.PowerPoint {
    public static class TestLazyInit {
        public static void Example_TestLazyInit(string folderPath, bool openPowerPoint) {
            Console.WriteLine("[*] PowerPoint - Testing Lazy Initialization");
            string filePath = Path.Combine(folderPath, "Test Lazy Init.pptx");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                Console.WriteLine($"  Initial slides: {presentation.Slides.Count} (expected: 0)");
                
                presentation.AddSlide().AddTextBox("First slide");
                Console.WriteLine($"  After 1st AddSlide: {presentation.Slides.Count} (expected: 1)");
                
                presentation.AddSlide().AddTextBox("Second slide");
                Console.WriteLine($"  After 2nd AddSlide: {presentation.Slides.Count} (expected: 2)");
                
                presentation.Save();
            }

            // Test opening existing file
            using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                Console.WriteLine($"  Reopened slides: {presentation.Slides.Count} (expected: 2)");
                
                presentation.AddSlide().AddTextBox("Third slide");
                Console.WriteLine($"  After 3rd AddSlide: {presentation.Slides.Count} (expected: 3)");
                
                presentation.Save();
            }
        }
    }
}