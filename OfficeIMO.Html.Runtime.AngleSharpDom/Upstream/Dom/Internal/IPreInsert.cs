namespace AngleSharp.Dom;

using System;

internal interface IPreInsert
{
    void PreInsert(Node parent, Node node, Int32 index);
}
