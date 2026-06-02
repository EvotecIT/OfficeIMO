using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    public partial class VisioPage {

        private sealed class ShapeCollection : IList<VisioShape> {
            private readonly VisioPage _page;

            public ShapeCollection(VisioPage page) {
                _page = page;
            }

            public VisioShape this[int index] {
                get => _page._shapes[index];
                set {
                    if (ReferenceEquals(_page._shapes[index], value)) {
                        return;
                    }

                    _page.PrepareShapeForPage(value);
                    _page._shapes[index] = value;
                }
            }

            public int Count => _page._shapes.Count;

            public bool IsReadOnly => false;

            public void Add(VisioShape item) {
                _page.PrepareShapeForPage(item);
                _page._shapes.Add(item);
            }

            public void Clear() => _page._shapes.Clear();

            public bool Contains(VisioShape item) => _page._shapes.Contains(item);

            public void CopyTo(VisioShape[] array, int arrayIndex) => _page._shapes.CopyTo(array, arrayIndex);

            public IEnumerator<VisioShape> GetEnumerator() => _page._shapes.GetEnumerator();

            public int IndexOf(VisioShape item) => _page._shapes.IndexOf(item);

            public void Insert(int index, VisioShape item) {
                _page.PrepareShapeForPage(item);
                _page._shapes.Insert(index, item);
            }

            public bool Remove(VisioShape item) => _page._shapes.Remove(item);

            public void RemoveAt(int index) => _page._shapes.RemoveAt(index);

            System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
        }

        private sealed class ConnectorCollection : IList<VisioConnector> {
            private readonly VisioPage _page;

            public ConnectorCollection(VisioPage page) {
                _page = page;
            }

            public VisioConnector this[int index] {
                get => _page._connectors[index];
                set {
                    VisioConnector existing = _page._connectors[index];
                    if (ReferenceEquals(existing, value)) {
                        return;
                    }

                    _page.PrepareConnectorForPage(value, existing);
                    _page._connectors[index] = value;
                }
            }

            public int Count => _page._connectors.Count;

            public bool IsReadOnly => false;

            public void Add(VisioConnector item) {
                _page.PrepareConnectorForPage(item);
                _page._connectors.Add(item);
            }

            public void Clear() => _page._connectors.Clear();

            public bool Contains(VisioConnector item) => _page._connectors.Contains(item);

            public void CopyTo(VisioConnector[] array, int arrayIndex) => _page._connectors.CopyTo(array, arrayIndex);

            public IEnumerator<VisioConnector> GetEnumerator() => _page._connectors.GetEnumerator();

            public int IndexOf(VisioConnector item) => _page._connectors.IndexOf(item);

            public void Insert(int index, VisioConnector item) {
                _page.PrepareConnectorForPage(item);
                _page._connectors.Insert(index, item);
            }

            public bool Remove(VisioConnector item) => _page._connectors.Remove(item);

            public void RemoveAt(int index) => _page._connectors.RemoveAt(index);

            System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
        }
    }
}
