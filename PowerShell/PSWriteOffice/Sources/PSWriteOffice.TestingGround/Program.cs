using OfficeIMO.PowerPoint;

namespace PSWriteOffice.TestingGround;

internal class Program
{
    static void Main(string[] args)
    {
        Console.WriteLine("Hello, World!");
        ExampleCreatePresentation();
    }

    public static void ExampleCreatePresentation()
    {
        var filePath = "my_pres.pptx";
        using var presentation = PowerPointPresentation.Create(filePath);
        var slide = presentation.AddSlide();
        slide.AddTitle("Hello World");
        slide.AddTextBoxPoints("Generated with PSWriteOffice", 72, 144, 400, 60);
        presentation.Save();
    }
}
