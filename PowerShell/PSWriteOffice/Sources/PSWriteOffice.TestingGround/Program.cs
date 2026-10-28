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
        using var presentation = PowerPointDocument.Create(filePath);
        var slide = presentation.AddSlide();
        slide.AddTitle("Hello World");
        slide.AddTextBox("Generated with PSWriteOffice", OfficeIMO.OfficeLength.Points(72), OfficeIMO.OfficeLength.Points(144), OfficeIMO.OfficeLength.Points(400), OfficeIMO.OfficeLength.Points(60));
        presentation.Save();
    }
}
