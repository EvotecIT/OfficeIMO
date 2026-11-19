using OfficeIMO.PowerPoint;

namespace OfficeIMO.Examples.PowerPoint {
    /// <summary>
    /// Reads a deck slide by slide with PowerPointSlideReader, as the README shows, without loading the presentation model.
    /// </summary>
    public static class StreamingSlideReaderPowerPoint {
        public static void Example_StreamingSlideReader(string folderPath, bool openPowerPoint) {
            Console.WriteLine("[*] PowerPoint - Read slides one at a time");
            string filePath = Path.Combine(folderPath, "StreamingSlideReader.pptx");
            using (PowerPointDocument presentation = PowerPointDocument.Create(filePath)) {
                for (int number = 1; number <= 3; number++) {
                    PowerPointSlide slide = presentation.AddSlide();
                    slide.AddTitle("Review " + number);
                    PowerPointTable table = slide.AddTable(2, 2);
                    table.GetRow(0).GetCell(0).Text = "Region";
                    table.GetRow(0).GetCell(1).Text = "Revenue";
                    table.GetRow(1).GetCell(0).Text = "North";
                    table.GetRow(1).GetCell(1).Text = (number * 100).ToString(System.Globalization.CultureInfo.InvariantCulture);
                    slide.Notes.Text = "Talk about region " + number;
                }
                presentation.Save();
            }

            using var reader = PowerPointSlideReader.Open(filePath);
            foreach (PowerPointSlideContent slide in reader.ReadSlides()) {
                Console.WriteLine($"{slide.Number}: {slide.Title}");
                foreach (PowerPointTableContent table in slide.Tables) {
                    Console.WriteLine($"  Table with {table.Rows.Count} rows");
                }
                foreach (PowerPointImageReference image in slide.Images) {
                    Console.WriteLine($"  Image {image.PartUri} ({image.ContentType})");
                }
                if (slide.Notes != null) Console.WriteLine($"  Notes: {slide.Notes}");
            }

            foreach (PowerPointExtractChunk chunk in reader.ExtractMarkdownChunks(sourcePath: filePath)) {
                Console.WriteLine(chunk.Markdown);
            }
        }
    }
}
