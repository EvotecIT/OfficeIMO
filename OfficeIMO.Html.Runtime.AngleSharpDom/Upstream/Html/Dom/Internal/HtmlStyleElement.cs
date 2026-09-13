namespace AngleSharp.Html.Dom
{
    using AngleSharp.Css;
    using AngleSharp.Dom;
    using AngleSharp.Html.Construction;
    using AngleSharp.Io;
    using System;
    using System.Threading;
    using System.Threading.Tasks;

    /// <summary>
    /// Represents the HTML style element.
    /// </summary>
    sealed class HtmlStyleElement : HtmlElement, IHtmlStyleElement, IConstructableStyleSheetElement
    {
        #region Fields

        private IStyleSheet? _sheet;
        private Boolean _parserInserted;
        private Boolean? _scriptBlockingEligible;
        private Int32 _sheetGeneration;
        private TaskCompletionSource<Boolean> _sheetRetired = NewSheetRetired();
        private SheetLoad? _sheetLoad;

        #endregion

        #region ctor

        public HtmlStyleElement(Document owner, String? prefix = null)
            : base(owner, TagNames.Style, prefix, NodeFlags.Special | NodeFlags.LiteralText)
        {
        }

        #endregion

        #region Properties

        public Boolean IsScoped
        {
            get => this.GetBoolAttribute(AttributeNames.Scoped);
            set => this.SetBoolAttribute(AttributeNames.Scoped, value);
        }

        public IStyleSheet? Sheet => _sheet;

        public Boolean IsDisabled
        {
            get => this.GetBoolAttribute(AttributeNames.Disabled);
            set
            {
                this.SetBoolAttribute(AttributeNames.Disabled, value);

                if (_sheet != null)
                {
                    _sheet.IsDisabled = value;
                }
                Owner?.SignalScriptBlockingStylesChanged();
            }
        }

        public String? Media
        {
            get => this.GetOwnAttribute(AttributeNames.Media);
            set => this.SetOwnAttribute(AttributeNames.Media, value);
        }

        public String? Type
        {
            get => this.GetOwnAttribute(AttributeNames.Type);
            set => this.SetOwnAttribute(AttributeNames.Type, value);
        }

        #endregion

        #region Internal Methods

        internal override void SetupElement()
        {
            base.SetupElement();
            UpdateSheet();
        }

        internal void UpdateMedia(String value)
        {
            if (_sheet != null)
            {
                _sheet.Media.MediaText = value;
            }
            Owner?.SignalScriptBlockingStylesChanged();
        }

        #endregion

        #region Helpers

        protected override void NodeIsInserted(Node newNode)
        {
            base.NodeIsInserted(newNode);
            if (!IsReplacingAll)
            {
                UpdateSheet();
            }
        }

        protected override void NodeIsRemoved(Node removedNode, Node? oldPreviousSibling)
        {
            base.NodeIsRemoved(removedNode, oldPreviousSibling);
            if (!IsReplacingAll)
            {
                UpdateSheet();
            }
        }

        protected override void ReplacedAll() => UpdateSheet();

        private void UpdateSheet()
        {
            var document = Owner;

            if (document != null)
            {
                var context = Context;
                var type = Type ?? MimeTypeNames.Css;
                var engine = context.GetStyling(type);

                if (engine != null)
                {
                    if (_parserInserted && !_scriptBlockingEligible.HasValue)
                    {
                        _scriptBlockingEligible = !IsDisabled;
                    }

                    var cancellation = new CancellationTokenSource();
                    var previousLoad = _sheetLoad;
                    var retired = _sheetRetired;
                    _sheetRetired = NewSheetRetired();
                    retired.TrySetResult(true);
                    var generation = Interlocked.Increment(ref _sheetGeneration);
                    var task = CreateSheetAsync(engine, document, generation, cancellation);
                    document.DelayLoadUntilRetired(task, _sheetRetired.Task);
                    var currentLoad = new SheetLoad(cancellation);
                    _sheetLoad = currentLoad;
                    if (_scriptBlockingEligible == true)
                    {
                        document.AddScriptBlockingStyle(task,
                            () => !currentLoad.IsRetired && !IsDisabled && document.IsScriptBlockingMedia(Media));
                    }
                    previousLoad?.Retire();
                    document.SignalScriptBlockingStylesChanged();
                }
            }
        }

        void IConstructableStyleSheetElement.MarkParserInserted() => _parserInserted = true;

        private async Task CreateSheetAsync(IStylingService engine, IDocument document, Int32 generation, CancellationTokenSource cancellation)
        {
            using var response = VirtualResponse.Create(res => res.Content(TextContent).Address(default(Url)));
            var options = new StyleOptions(document)
            {
                Element = this,
                IsDisabled = IsDisabled,
                IsAlternate = false,
            };
            try
            {
                var sheet = await engine.ParseStylesheetAsync(response, options, cancellation.Token).ConfigureAwait(false);
                var syncRoot = document.Context.GetService<IDomSynchronization>()?.SyncRoot;
                if (syncRoot is null) ApplySheet(sheet);
                else lock (syncRoot) ApplySheet(sheet);
            }
            catch (OperationCanceledException) when (cancellation.IsCancellationRequested)
            {
            }
            finally
            {
                cancellation.Dispose();
            }

            void ApplySheet(IStyleSheet sheet)
            {
                if (generation != Volatile.Read(ref _sheetGeneration))
                {
                    return;
                }

                sheet.IsDisabled = IsDisabled;
                sheet.Media.MediaText = Media ?? String.Empty;
                _sheet = sheet;
            }
        }

        private static TaskCompletionSource<Boolean> NewSheetRetired() =>
            new(TaskCreationOptions.RunContinuationsAsynchronously);

        private sealed class SheetLoad
        {
            private readonly CancellationTokenSource _cancellation;
            private Int32 _retired;

            internal SheetLoad(CancellationTokenSource cancellation) => _cancellation = cancellation;

            internal Boolean IsRetired => Volatile.Read(ref _retired) != 0;

            internal void Retire()
            {
                Interlocked.Exchange(ref _retired, 1);
                try
                {
                    _cancellation.Cancel();
                }
                catch (ObjectDisposedException)
                {
                }
            }
        }

        #endregion
    }
}
