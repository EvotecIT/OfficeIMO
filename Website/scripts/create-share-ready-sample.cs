#:project ../../OfficeIMO.Word/OfficeIMO.Word.csproj
// Creates the sample the "Remove hidden data" browser tool loads, in Apps/OfficeIMO.Web.Converter/wwwroot/samples:
//   share-ready-sample.docx  a short proposal carrying the hidden data people forget about: author and company properties,
//                            a custom property, a reviewer comment, a tracked change, hidden text, a document variable,
//                            and a picture with a GPS location and an AI source label (the label comes from provenance-demo.png).
// Every name and value is invented for the demo. Run from the Website folder:  dotnet run scripts/create-share-ready-sample.cs
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Word;

string folder = Path.GetFullPath(Path.Combine("Apps", "OfficeIMO.Web.Converter", "wwwroot", "samples"));
string picture = Path.Combine(Path.GetTempPath(), "share-ready-sample-picture.png");
byte[] source = File.ReadAllBytes(Path.Combine(folder, "provenance-demo.png"));
OfficeImageMetadata metadata = OfficeImageMetadata.Read(source);
metadata.SetExifValue(OfficeExifTag.GPSLatitudeRef, "N");
metadata.SetExifValue(OfficeExifTag.Artist, "Dana Whitfield");
File.WriteAllBytes(picture, OfficeImageMetadata.Apply(source, metadata));

string output = Path.Combine(folder, "share-ready-sample.docx");
using (WordDocument document = WordDocument.Create(output)) {
    document.Properties.Title = "Office move proposal";
    document.Properties.Creator = "Dana Whitfield";
    document.Properties.LastModifiedBy = "Sam Okafor";
    document.ApplicationProperties.Company = "Northwind Facilities (sample)";
    document.ApplicationProperties.Manager = "Priya Raman";
    document.CustomProperties.Add("ClientReference", new WordCustomProperty("NW-2026-118"));
    document.SetDocumentVariable("InternalBudgetCode", "FAC-OPS-7");

    document.AddParagraph("Office move proposal").Style = WordParagraphStyles.Heading1;
    WordParagraph intro = document.AddParagraph("We propose moving the support team to the second floor in March. ");
    intro.AddText("Internal note: the landlord offered a 12% discount if we sign before Friday.").Hidden = true;
    document.AddParagraph("The move takes two weekends and keeps every desk online during working hours.")
        .AddComment("Sam Okafor", "SO", "Can we promise this? IT hasn't confirmed the network work yet.");
    document.AddParagraph("Estimated cost: ").AddDeletedText("48,000", "Priya Raman").AddInsertedText("52,000", "Priya Raman").AddText(" EUR.");
    document.AddParagraph("Floor plan concept:");
    document.AddParagraph().AddImage(picture, OfficeLength.Points(240), OfficeLength.Points(160));
    document.AddParagraph("Sample document created for the OfficeIMO browser tools. All names are invented.");
    document.Save();
}
File.Delete(picture);
Console.WriteLine("Wrote " + output);
