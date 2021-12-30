using System.IO;

namespace OfficeIMO.Tests
{
    public partial class Word
    {
        private readonly string _directoryDocuments;
        private readonly string _directoryWithFiles;

        private static void Setup(string path)
        {
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            } else {
                Directory.Delete(path, true);
                Directory.CreateDirectory(path);
            }
        }
        public Word() {
            _directoryDocuments = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "DocXTests", "documents");
            Setup(_directoryDocuments); // prepare temp documents directory 
            _directoryWithFiles = TestHelper.DirectoryWithFiles;
        }
    }
}
