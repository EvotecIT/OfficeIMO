namespace AngleSharp.Dom
{
    using AngleSharp.Browser;
    using AngleSharp.Js;
    using System;
    using System.Threading;
    using System.Threading.Tasks;

    /// <summary>
    /// A set of extensions for interacting with the document.
    /// </summary>
    public static class JsDocumentExtensions
    {
        /// <summary>
        /// Enqueues the given action as a normal task to the document,
        /// thus ensuring it runs after any script-related actions.
        /// </summary>
        /// <param name="document">The document to extend.</param>
        /// <param name="action">The action on the document to perform.</param>
        /// <returns>A task that is finished when the enqueued task completed.</returns>
        public static Task<IDocument> Then(this IDocument document, Action<IDocument> action)
        {
            var context = document.Context;
            var evts = context.GetService<IEventLoop>();

            if (evts != null)
            {
                var tcs = new TaskCompletionSource<Boolean>();
                evts.Enqueue(cancel =>
                {
                    try
                    {
                        action?.Invoke(document);
                        tcs.SetResult(true);
                    }
                    catch (Exception ex)
                    {
                        tcs.SetException(ex);
                    }
                }, TaskPriority.None);
                return tcs.Task.ContinueWith(_ => document);
            }

            action?.Invoke(document);
            return Task.FromResult(document);
        }

        /// <summary>
        /// Enqueues the given JavaScript code as a normal task to the document,
        /// thus ensuring it runs after any script-related actions.
        /// </summary>
        /// <param name="document">The document to extend.</param>
        /// <param name="jsSource">The JavaScript action to perform.</param>
        /// <returns>A task that is finished when the enqueued task completed.</returns>
        public static Task<IDocument> Then(this IDocument document, String jsSource) =>
            document.Then(current => current.ExecuteScript(jsSource));

        /// <summary>
        /// Enqueues the given action as a normal task to the document,
        /// thus ensuring it runs after any script-related actions.
        /// </summary>
        /// <param name="documentTask">The soon available document.</param>
        /// <param name="action">The action on the document to perform.</param>
        /// <returns>A task that is finished when the enqueued task completed.</returns>
        public static async Task<IDocument> Then(this Task<IDocument> documentTask, Action<IDocument> action)
        {
            var document = await documentTask.ConfigureAwait(false);
            return await document.Then(action).ConfigureAwait(false);
        }

        /// <summary>
        /// Enqueues the given JavaScript code as a normal task to the document,
        /// thus ensuring it runs after any script-related actions.
        /// </summary>
        /// <param name="documentTask">The soon available document.</param>
        /// <param name="jsSource">The JavaScript action to perform.</param>
        /// <returns>A task that is finished when the enqueued task completed.</returns>
        public static Task<IDocument> Then(this Task<IDocument> documentTask, String jsSource) =>
            documentTask.Then(document => document.ExecuteScript(jsSource));

        /// <summary>
        /// Waits until all currently queued tasks finished.
        /// </summary>
        /// <param name="document">The available document.</param>
        /// <returns>A task that is finished when the enqueued tasks completed.</returns>
        public static Task<IDocument> WhenStable(this IDocument document) =>
            document.Then(_ => { });

        /// <summary>
        /// Waits until all initially queued tasks finished.
        /// </summary>
        /// <param name="documentTask">The soon available document.</param>
        /// <returns>A task that is finished when the enqueued tasks completed.</returns>
        public static Task<IDocument> WhenStable(this Task<IDocument> documentTask) =>
            documentTask.Then(_ => { });

        /// <summary>
        /// Waits until the document from the task is fully available.
        /// This includes that the document is completed and that a stable
        /// point has been reached.
        /// </summary>
        /// <param name="documentTask">The document task to await.</param>
        /// <param name="cancellation">The timeout cancellation, if any.</param>
        /// <returns>A task that is finished when the document is available.</returns>
        public static async Task<IDocument> WaitUntilAvailable(this Task<IDocument> documentTask, CancellationToken cancellation = default)
        {
            var document = await documentTask.ConfigureAwait(false);
            await document.WaitUntilAvailable(cancellation).ConfigureAwait(false);
            return document;
        }

        /// <summary>
        /// Waits until the document is fully available. This includes
        /// that the document is completed and that a stable point has
        /// been reached.
        /// </summary>
        /// <param name="document">The document to await.</param>
        /// <param name="cancellation">The timeout cancellation, if any.</param>
        /// <returns>A task that is finished when the document is available.</returns>
        public static async Task<IDocument> WaitUntilAvailable(this IDocument document, CancellationToken cancellation = default)
        {
            if (document.ReadyState != DocumentReadyState.Complete)
            {
                var abort = new TaskCompletionSource<object>();
                cancellation.Register(() => abort.TrySetCanceled());
                await Task.WhenAny(document.AwaitEventAsync("load"), abort.Task).ConfigureAwait(false);
            }

            await document.WhenStable().ConfigureAwait(false);
            return document;
        }
    }
}
