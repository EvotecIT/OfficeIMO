using System;
using System.IO;
using OfficeIMO.PowerPoint;

class TestPPTXBehavior {
    static void Main() {
        string file1 = "test_new.pptx";
        
        // Test new presentation behavior
        using (var pres = PowerPointPresentation.Create(file1)) {
            Console.WriteLine($"New presentation - Initial slides: {pres.Slides.Count}");
            
            var slide1 = pres.AddSlide();
            Console.WriteLine($"After first AddSlide() - Slides: {pres.Slides.Count}");
            
            var slide2 = pres.AddSlide();
            Console.WriteLine($"After second AddSlide() - Slides: {pres.Slides.Count}");
            
            pres.Save();
        }
        
        // Verify after reload
        using (var pres = PowerPointPresentation.Open(file1)) {
            Console.WriteLine($"After reload - Slides: {pres.Slides.Count}");
        }
        
        File.Delete(file1);
        Console.WriteLine("\nExpected behavior: 1, 1, 2, 2 (first AddSlide reuses initial)");
    }
}