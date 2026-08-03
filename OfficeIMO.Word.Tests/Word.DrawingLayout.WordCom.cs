using OfficeIMO.Word;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [WordDesktopLayoutFact]
        [Trait("Category", "MicrosoftOfficeInteroperability")]
        public void DrawingLayout_AnchoredGroupMatchesDesktopWordGeometryWhenRequested() {
            Assert.True(IsWindowsPlatform(), "Desktop Word layout validation requires Windows.");
            Assert.True(IsWordComAvailable(), "Desktop Word layout validation requires Microsoft Word COM automation.");

            string directory = Path.Combine(_directoryWithFiles, "WordDesktopLayout", GetCurrentTargetFrameworkLabel());
            Directory.CreateDirectory(directory);
            string path = Path.Combine(directory, "anchored-shape-group.docx");
            if (File.Exists(path)) File.Delete(path);

            using (WordDocument document = WordDocument.Create(path)) {
                WordShapeGroup group = document.AddParagraph().AddShapeGroup(new[] {
                    new WordShapeGroupItem(ShapeType.Chevron, 0, 0, 80, 40),
                    new WordShapeGroupItem(ShapeType.Chevron, 72, 0, 80, 40),
                    new WordShapeGroupItem(ShapeType.Chevron, 144, 0, 80, 40)
                }, 24, 48);
                Assert.True(group.TryGetLayoutSnapshot(out WordDrawingLayoutSnapshot packageLayout));
                Assert.Equal(224D, packageLayout.WidthPoints, 6);
                Assert.Equal(40D, packageLayout.HeightPoints, 6);
                Assert.Equal(24D, packageLayout.HorizontalOffsetPoints!.Value, 6);
                Assert.Equal(48D, packageLayout.VerticalOffsetPoints!.Value, 6);
                document.Save();
            }

            (double Left, double Top, double Width, double Height) rendered = ReadFirstShapeGeometryViaWordCom(path);
            Assert.InRange(rendered.Left, 23D, 25D);
            Assert.InRange(rendered.Top, 47D, 49D);
            Assert.InRange(rendered.Width, 223D, 225D);
            Assert.InRange(rendered.Height, 39D, 41D);
        }

        private static (double Left, double Top, double Width, double Height) ReadFirstShapeGeometryViaWordCom(string path) {
            var failures = new List<string>();
            (double Left, double Top, double Width, double Height)? result = null;
            var thread = new Thread(() => {
                object? word = null;
                object? documents = null;
                object? document = null;
                object? shapes = null;
                object? shape = null;
                try {
                    word = CreateWordComApplication();
                    documents = GetComProperty(word, "Documents");
                    document = InvokeCom(documents!, "Open", path, false, true, false);
                    InvokeCom(document!, "Repaginate");
                    shapes = GetComProperty(document!, "Shapes");
                    int count = Convert.ToInt32(GetComProperty(shapes!, "Count"), System.Globalization.CultureInfo.InvariantCulture);
                    if (count < 1) {
                        throw new InvalidOperationException("Desktop Word did not expose the anchored group through Document.Shapes.");
                    }
                    shape = InvokeCom(shapes!, "Item", 1);
                    result = (
                        Convert.ToDouble(GetComProperty(shape!, "Left"), System.Globalization.CultureInfo.InvariantCulture),
                        Convert.ToDouble(GetComProperty(shape!, "Top"), System.Globalization.CultureInfo.InvariantCulture),
                        Convert.ToDouble(GetComProperty(shape!, "Width"), System.Globalization.CultureInfo.InvariantCulture),
                        Convert.ToDouble(GetComProperty(shape!, "Height"), System.Globalization.CultureInfo.InvariantCulture));
                } catch (Exception exception) when (exception is COMException or InvalidOperationException or MissingMethodException or TargetInvocationException) {
                    failures.Add(DescribeWordComFailure(exception));
                } finally {
                    try {
                        CloseWordComDocument(document);
                    } catch (Exception exception) {
                        failures.Add(DescribeWordComFailure(exception));
                    }
                    QuitWordComApplication(word);
                    ReleaseComObject(shape);
                    ReleaseComObject(shapes);
                    ReleaseComObject(document);
                    ReleaseComObject(documents);
                    ReleaseComObject(word);
                }
            });
            thread.IsBackground = true;
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            if (!thread.Join(WordComOpenTimeout)) {
                failures.Add($"Desktop Word layout validation timed out after {WordComOpenTimeout.TotalSeconds:0} seconds.");
            }

            Assert.True(failures.Count == 0, string.Join(Environment.NewLine, failures));
            Assert.True(result.HasValue, "Desktop Word did not return anchored shape geometry.");
            return result!.Value;
        }
    }
}
