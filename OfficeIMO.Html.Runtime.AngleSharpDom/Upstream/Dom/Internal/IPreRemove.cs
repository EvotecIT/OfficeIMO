namespace AngleSharp.Dom;

using System;

internal interface IPreRemove
{
    void PreRemove(Node parent, Node node, Int32 index);
}
