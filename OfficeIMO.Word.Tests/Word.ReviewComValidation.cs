using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.CSharp.RuntimeBinder;
#if NET5_0_OR_GREATER
using System.Runtime.Versioning;
#endif
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        private const string WordComValidationEnv = "OFFICEIMO_RUN_WORD_COM_VALIDATION";
        private const int WdFormatXmlDocument = 16;
        private static readonly TimeSpan WordComValidationTimeout = TimeSpan.FromMinutes(2);

        [Fact]
        public void Test_InspectReview_WordComGeneratedReviewDocumentWhenRequested() {
            if (!IsWordComValidationRequested()) {
                return;
            }

            Assert.True(IsWindowsPlatform(), "Word COM validation requires Windows.");
            Assert.True(IsWordComAvailable(), "Word COM validation requires Microsoft Word COM automation.");

            string directory = Path.Combine(_directoryWithFiles, "WordComReview", GetCurrentTargetFrameworkLabel());
            Directory.CreateDirectory(directory);
            string filePath = Path.Combine(directory, "word-com-review-corpus.docx");
            File.Delete(filePath);

            CreateReviewedDocumentViaWordCom(filePath);

            using WordDocument document = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
            WordReviewInfo review = document.InspectReview();

            Assert.True(review.HasReviewMetadata);
            Assert.Contains(review.Comments, comment =>
                comment.Text.Contains("Word COM body comment", StringComparison.Ordinal)
                && comment.TargetText.Contains("Word COM comment target", StringComparison.Ordinal));
            Assert.Contains(review.Comments, comment =>
                comment.Text.Contains("Word COM table comment", StringComparison.Ordinal)
                && comment.TargetText.Contains("Word COM table target", StringComparison.Ordinal));
            Assert.Contains(review.Revisions, revision =>
                revision.RevisionType == WordReviewRevisionType.Insertion
                && revision.AffectedText.Contains("Word COM inserted body revision", StringComparison.Ordinal));
            Assert.Contains(review.Revisions, revision =>
                revision.RevisionType == WordReviewRevisionType.Deletion
                && revision.AffectedText.Contains("Word COM deletion target", StringComparison.Ordinal));

            WordReviewReport report = document.InspectReviewReport();
            string json = report.ToJson();
            string markdown = report.ToMarkdown();
            Assert.Contains("Word COM body comment", json, StringComparison.Ordinal);
            Assert.Contains("Word COM table comment", json, StringComparison.Ordinal);
            Assert.Contains("Word COM inserted body revision", markdown, StringComparison.Ordinal);
            Assert.Contains("Word COM deletion target", markdown, StringComparison.Ordinal);
        }

        private static bool IsWordComValidationRequested() {
            string? value = Environment.GetEnvironmentVariable(WordComValidationEnv);
            return string.Equals(value, "1", StringComparison.Ordinal)
                || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);
        }

#if NET5_0_OR_GREATER
        [SupportedOSPlatform("windows")]
#endif
        private static void CreateReviewedDocumentViaWordCom(string path) {
            var failures = new List<string>();
            var thread = new Thread(() => CreateReviewedDocumentViaWordComOnStaThread(path, failures));
            thread.IsBackground = true;
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            if (!thread.Join(WordComValidationTimeout)) {
                failures.Add($"Word COM review validation timed out after {WordComValidationTimeout.TotalSeconds:0} seconds.");
            }

            Assert.True(failures.Count == 0, string.Join(Environment.NewLine, failures));
            Assert.True(File.Exists(path), "Word COM did not create the expected review corpus document.");
        }

#if NET5_0_OR_GREATER
        [SupportedOSPlatform("windows")]
#endif
        private static void CreateReviewedDocumentViaWordComOnStaThread(string path, List<string> failures) {
            dynamic? word = null;
            dynamic? documents = null;
            dynamic? document = null;

            try {
                Type wordType = Type.GetTypeFromProgID("Word.Application")
                    ?? throw new InvalidOperationException("Word COM automation is not available.");
                word = Activator.CreateInstance(wordType)
                    ?? throw new InvalidOperationException("Failed to create Word COM automation instance.");
                word.DisplayAlerts = 0;
                word.Visible = false;

                documents = word.Documents;
                document = documents.Add();

                dynamic body = document.Content;
                body.Text = "Word COM comment target\rWord COM deletion target\r";

                dynamic commentRange = document.Range(0, "Word COM comment target".Length);
                document.Comments.Add(commentRange, "Word COM body comment");

                dynamic tableAnchor = document.Range(document.Content.End - 1, document.Content.End - 1);
                dynamic table = document.Tables.Add(tableAnchor, 1, 1);
                dynamic cellRange = table.Cell(1, 1).Range;
                cellRange.Text = "Word COM table target";
                document.Comments.Add(cellRange, "Word COM table comment");

                document.TrackRevisions = true;
                dynamic insertionRange = document.Range(document.Content.End - 1, document.Content.End - 1);
                insertionRange.InsertAfter("\rWord COM inserted body revision");

                dynamic deletionRange = document.Content;
                dynamic find = deletionRange.Find;
                find.ClearFormatting();
                bool found = find.Execute("Word COM deletion target");
                if (!found) {
                    failures.Add("Word COM could not find the deletion target text.");
                } else {
                    deletionRange.Delete();
                }

                document.TrackRevisions = false;
                document.SaveAs2(path, WdFormatXmlDocument);
            } catch (Exception ex) when (ex is COMException or InvalidOperationException or MissingMethodException or TargetInvocationException or RuntimeBinderException) {
                failures.Add(DescribeWordComFailure(ex));
            } finally {
                try {
                    document?.Close(false);
                } catch (Exception ex) when (ex is COMException or MissingMethodException or TargetInvocationException or RuntimeBinderException) {
                    failures.Add("Word document close: " + DescribeWordComFailure(ex));
                }

                try {
                    word?.Quit();
                } catch (Exception ex) when (ex is COMException or MissingMethodException or TargetInvocationException or RuntimeBinderException) {
                    failures.Add("Word quit: " + DescribeWordComFailure(ex));
                }

                ReleaseComObject(document);
                ReleaseComObject(documents);
                ReleaseComObject(word);
            }
        }

    }
}
