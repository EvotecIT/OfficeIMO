namespace AngleSharp.Html.Dom;

using AngleSharp.Dom;
using System;

/// <summary>
/// Represents the HTML selectedcontent element.
/// </summary>
sealed class HtmlSelectedContentElement : HtmlElement
{
    #region ctor

    public HtmlSelectedContentElement(Document owner, String? prefix = null)
        : base(owner, TagNames.SelectedContent, prefix)
    {
    }

    #endregion
}
