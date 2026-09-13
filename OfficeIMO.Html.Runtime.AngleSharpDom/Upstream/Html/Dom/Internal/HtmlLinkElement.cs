namespace AngleSharp.Html.Dom
{
    using AngleSharp.Dom;
    using AngleSharp.Html.Construction;
    using AngleSharp.Html.LinkRels;
    using AngleSharp.Io;
    using AngleSharp.Text;
    using System;
    using System.Threading;
    using System.Threading.Tasks;

    /// <summary>
    /// Represents the HTML link element.
    /// </summary>
    sealed class HtmlLinkElement : HtmlElement, IHtmlLinkElement, IConstructableStyleSheetElement
    {
        #region Fields

        private BaseLinkRelation? _relation;
        private TokenList? _relList;
        private SettableTokenList? _sizes;

        private String? _source;
        private String? _relationValue;
        private Boolean _relationLoaded;
        private Boolean _parserInserted;
        private Boolean? _scriptBlockingEligible;
        private TaskCompletionSource<Boolean> _loadRetired = NewLoadRetired();
        private ScriptBlockingLoad? _scriptBlockingLoad;

        #endregion

        #region ctor

        public HtmlLinkElement(Document owner, String? prefix = null)
            : base(owner, TagNames.Link, prefix, NodeFlags.Special | NodeFlags.SelfClosing)
        {
        }

        #endregion

        #region Design properties

        internal Boolean IsVisited
        {
            get;
            set;
        }

        internal Boolean IsActive
        {
            get;
            set;
        }

        #endregion

        #region Properties

        public IDownload? CurrentDownload => _relation?.Processor?.Download;

        public String? Href
        {
            get => this.GetUrlAttribute(AttributeNames.Href);
            set => this.SetOwnAttribute(AttributeNames.Href, value);
        }

        public String? TargetLanguage
        {
            get => this.GetOwnAttribute(AttributeNames.HrefLang);
            set => this.SetOwnAttribute(AttributeNames.HrefLang, value);
        }

        public String? Charset
        {
            get => this.GetOwnAttribute(AttributeNames.Charset);
            set => this.SetOwnAttribute(AttributeNames.Charset, value);
        }

        public String? Relation
        {
            get => this.GetOwnAttribute(AttributeNames.Rel);
            set => this.SetOwnAttribute(AttributeNames.Rel, value);
        }

        public String? ReverseRelation
        {
            get => this.GetOwnAttribute(AttributeNames.Rev);
            set => this.SetOwnAttribute(AttributeNames.Rev, value);
        }

        public String? NumberUsedOnce
        {
            get => this.GetOwnAttribute(AttributeNames.Nonce);
            set => this.SetOwnAttribute(AttributeNames.Nonce, value);
        }

        public ITokenList RelationList
        {
            get
            {
                if (_relList is null)
                {
                    _relList = new TokenList(this.GetOwnAttribute(AttributeNames.Rel));
                    _relList.Changed += value => UpdateAttribute(AttributeNames.Rel, value);
                }

                return _relList;
            }
        }

        public ISettableTokenList Sizes
        {
            get
            {
                if (_sizes is null)
                {
                    _sizes = new SettableTokenList(this.GetOwnAttribute(AttributeNames.Sizes));
                    _sizes.Changed += value => UpdateAttribute(AttributeNames.Sizes, value);
                }

                return _sizes;
            }
        }

        public String? Rev
        {
            get => this.GetOwnAttribute(AttributeNames.Rev);
            set => this.SetOwnAttribute(AttributeNames.Rev, value);
        }

        public Boolean IsDisabled
        {
            get => this.GetBoolAttribute(AttributeNames.Disabled);
            set => this.SetBoolAttribute(AttributeNames.Disabled, value);
        }

        public String? Target
        {
            get => this.GetOwnAttribute(AttributeNames.Target);
            set => this.SetOwnAttribute(AttributeNames.Target, value);
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

        public String? Integrity
        {
            get => this.GetOwnAttribute(AttributeNames.Integrity);
            set => this.SetOwnAttribute(AttributeNames.Integrity, value);
        }

        public IStyleSheet? Sheet
        {
            get
            {
                var sheetRelation = _relation as StyleSheetLinkRelation;
                return sheetRelation?.Sheet;
            }
        }

        public IDocument? Import
        {
            get
            {
                var importRelation = _relation as ImportLinkRelation;
                return importRelation?.Import;
            }
        }

        public String? CrossOrigin
        {
            get => this.GetOwnAttribute(AttributeNames.CrossOrigin);
            set => this.SetOwnAttribute(AttributeNames.CrossOrigin, value);
        }

        #endregion

        #region Internal Methods

        internal void UpdateRel(String value)
        {
            _relList?.Update(value);
            if (value.Isi(_relationValue))
            {
                return;
            }
            _relationValue = value;
            var previousRelation = _relation;
            var previousBlockingLoad = _scriptBlockingLoad;
            RetireLoad();
            _relation = CreateFirstLegalRelation();
            _relationLoaded = false;
            _scriptBlockingLoad = null;
            if (Owner != null && Object.ReferenceEquals(this.GetRoot(), Owner))
            {
                LoadRelation();
            }
            RetireScriptBlockingLoad(previousBlockingLoad);
            previousRelation?.Cancel();
        }

        internal void UpdateSizes(String value)
        {
            _sizes?.Update(value);
        }

        internal void UpdateMedia(String value)
        {
            var sheet = Sheet;

            if (sheet != null)
            {
                sheet.Media.MediaText = value;
            }
            Owner?.SignalScriptBlockingStylesChanged();
        }

        internal void UpdateDisabled(String value)
        {
            var sheet = Sheet;

            if (sheet != null)
            {
                sheet.IsDisabled = value != null;
            }
            Owner?.SignalScriptBlockingStylesChanged();

            if (value is null)
            {
                LoadRelation();
            }
        }

        internal void UpdateSource(String value)
        {
            if (_source is null && _relationLoaded)
            {
                _source = value;
                return;
            }

            if (!String.Equals(value, _source, StringComparison.Ordinal))
            {
                _source = value;
                var previousRelation = _relation;
                var previousBlockingLoad = _scriptBlockingLoad;
                RetireLoad();
                _relation = CreateFirstLegalRelation();
                _relationLoaded = false;
                _scriptBlockingLoad = null;
                if (Owner != null && Object.ReferenceEquals(this.GetRoot(), Owner))
                {
                    LoadRelation();
                }
                RetireScriptBlockingLoad(previousBlockingLoad);
                previousRelation?.Cancel();
            }
        }

        protected override void OnParentChanged()
        {
            base.OnParentChanged();

            LoadRelation();
        }

        internal void LoadRelation()
        {
            if (_relationLoaded || _relation == null || Href == null)
            {
                return;
            }

            var styleSheet = _relation is StyleSheetLinkRelation;
            if (_parserInserted && styleSheet && !_scriptBlockingEligible.HasValue)
            {
                _scriptBlockingEligible = !IsDisabled && !this.IsAlternate();
            }

            if (styleSheet && IsDisabled)
            {
                return;
            }

            _relationLoaded = true;

            var relation = _relation;
            var retired = _loadRetired.Task;
            var task = relation.LoadAsync();
            if (relation.DelaysDocumentLoad)
            {
                Owner?.DelayLoadUntilRetired(task, retired);
            }
            if (_scriptBlockingEligible == true && Owner != null)
            {
                var blockingLoad = new ScriptBlockingLoad();
                _scriptBlockingLoad = blockingLoad;
                Owner.AddScriptBlockingStyle(task, () => !blockingLoad.IsRetired
                    && !IsDisabled && !this.IsAlternate() && Owner.IsScriptBlockingMedia(Media));
            }
        }

        void IConstructableStyleSheetElement.MarkParserInserted() => _parserInserted = true;

        #endregion

        #region Helpers

        private BaseLinkRelation? CreateFirstLegalRelation()
        {
            var relations = RelationList;
            var factory = Context?.GetFactory<ILinkRelationFactory>();

            foreach (var relation in relations)
            {
                var rel = factory?.Create(this, relation);

                if (rel != null)
                {
                    return rel;
                }
            }

            return null;
        }

        private void RetireLoad()
        {
            var retired = _loadRetired;
            _loadRetired = NewLoadRetired();
            retired.TrySetResult(true);
        }

        private static TaskCompletionSource<Boolean> NewLoadRetired() =>
            new(TaskCreationOptions.RunContinuationsAsynchronously);

        private void RetireScriptBlockingLoad(ScriptBlockingLoad? load)
        {
            if (load != null)
            {
                load.Retire();
            }
            Owner?.SignalScriptBlockingStylesChanged();
        }

        private sealed class ScriptBlockingLoad
        {
            private Int32 _retired;

            internal Boolean IsRetired => Volatile.Read(ref _retired) != 0;

            internal void Retire() => Interlocked.Exchange(ref _retired, 1);
        }

        #endregion
    }
}
