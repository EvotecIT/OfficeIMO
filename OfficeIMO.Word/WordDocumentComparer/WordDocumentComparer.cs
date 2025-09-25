namespace OfficeIMO.Word {
    /// <summary>
    /// Provides simple comparison between two Word documents.
    /// </summary>
    public static partial class WordDocumentComparer {
        /// <summary>
        /// Compares two documents and produces a new document with revision marks.
        /// </summary>
        /// <param name="sourcePath">Path to the original document.</param>
        /// <param name="targetPath">Path to the modified document.</param>
        /// <returns>Document containing revision marks highlighting differences.</returns>
        public static WordDocument Compare(string sourcePath, string targetPath) {
            if (string.IsNullOrEmpty(sourcePath)) throw new ArgumentNullException(nameof(sourcePath));
            if (string.IsNullOrEmpty(targetPath)) throw new ArgumentNullException(nameof(targetPath));

            string resultPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".docx");
            File.Copy(sourcePath, resultPath, true);

            using (WordDocument source = WordDocument.Load(sourcePath))
            using (WordDocument target = WordDocument.Load(targetPath))
            using (WordDocument result = WordDocument.Load(resultPath)) {
                CompareParagraphs(source, target, result);
                CompareTables(source, target, result);
                CompareImages(source, target, result);

                result.Save(false);
            }

            return WordDocument.Load(resultPath);
        }
    }
}
