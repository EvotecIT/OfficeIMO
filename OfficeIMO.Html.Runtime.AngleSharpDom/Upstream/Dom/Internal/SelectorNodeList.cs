namespace AngleSharp.Dom;

using AngleSharp.Attributes;
using AngleSharp.Common;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;

// Keep the retained IParentNode.QuerySelectorAll ABI expected by AngleSharp.Css,
// while exposing the browser's static NodeList identity to the script bridge.
[DomName("NodeList")]
sealed class SelectorNodeList : IHtmlCollection<IElement>, INodeList
{
    private readonly IElement[] _elements;

    internal SelectorNodeList(IEnumerable<IElement> elements) => _elements = elements.ToArray();

    public Int32 Length => _elements.Length;
    public IElement this[Int32 index] => _elements[index];
    public IElement? this[String id] => _elements.GetElementById(id);
    INode INodeList.this[Int32 index] => _elements[index];

    public IEnumerator<IElement> GetEnumerator() => ((IEnumerable<IElement>)_elements).GetEnumerator();
    IEnumerator<INode> IEnumerable<INode>.GetEnumerator() => _elements.Cast<INode>().GetEnumerator();
    IEnumerator IEnumerable.GetEnumerator() => _elements.GetEnumerator();

    public void ToHtml(TextWriter writer, IMarkupFormatter formatter) {
        foreach (IElement element in _elements) element.ToHtml(writer, formatter);
    }
}
