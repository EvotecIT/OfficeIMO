#region QuickStart
using System.IO;
using OfficeIMO.Rtf;

RtfDocument document = RtfDocument.Create();
document.AddParagraph().AddText("Quarterly report").SetBold();
document.AddParagraph("Prepared by OfficeIMO");
File.WriteAllText("Report.rtf", document.ToRtf());

RtfDocument parsed = RtfDocument.Parse(File.ReadAllText("Report.rtf"));
#endregion
