using System.Data;
using System.Globalization;
using System.ComponentModel;
using System.Text;
using System.Threading;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private sealed class DirectDataSetColumnModel {
            internal DirectDataSetColumnModel(string name, Type dataType) {
                Name = name;
                DataType = dataType;
            }

            internal string Name { get; }

            internal Type DataType { get; }
        }

        private sealed class DirectDataSetSaveCandidate : IDisposable {
            private readonly DataSet _dataSet;
            private readonly Action _invalidate;
            private readonly bool _subscribed;
            private readonly HashSet<DataTable> _subscribedTables = new();
            private bool _disposed;

            internal DirectDataSetSaveCandidate(DataSet dataSet, DirectDataSetWorkbookModel model, Action invalidate, bool isDeferred, bool subscribeToSourceChanges) {
                _dataSet = dataSet;
                Model = model;
                _invalidate = invalidate;
                IsDeferred = isDeferred;
                if (subscribeToSourceChanges) {
                    _subscribed = true;
                    Subscribe(dataSet);
                }
            }

            internal DirectDataSetWorkbookModel Model { get; }

            internal DataSet Owner => _dataSet;

            internal Action InvalidateCallback => _invalidate;

            internal bool IsDeferred { get; }

            internal bool SubscribesToSourceChanges => _subscribed;

            internal bool IsValid { get; private set; } = true;

            internal DirectDataSetSaveCandidate WithModel(DirectDataSetWorkbookModel model) {
                return new DirectDataSetSaveCandidate(_dataSet, model, _invalidate, IsDeferred, _subscribed);
            }

            private void Subscribe(DataSet dataSet) {
                dataSet.Tables.CollectionChanged += OnCollectionChanged;
                foreach (DataTable table in dataSet.Tables) {
                    Subscribe(table);
                }
            }

            private void Subscribe(DataTable table) {
                if (!_subscribedTables.Add(table)) {
                    return;
                }

                table.Columns.CollectionChanged += OnCollectionChanged;
                table.RowChanged += OnDataChanged;
                table.RowChanging += OnDataChanging;
                table.RowDeleted += OnDataChanged;
                table.RowDeleting += OnDataChanging;
                table.ColumnChanged += OnColumnChanged;
                table.ColumnChanging += OnColumnChanging;
                table.TableCleared += OnDataChanged;
                table.TableClearing += OnDataChanging;
            }

            private void Unsubscribe(DataSet dataSet) {
                dataSet.Tables.CollectionChanged -= OnCollectionChanged;
                foreach (DataTable table in _subscribedTables.ToArray()) {
                    Unsubscribe(table);
                }
            }

            private void Unsubscribe(DataTable table) {
                if (!_subscribedTables.Remove(table)) {
                    return;
                }

                table.Columns.CollectionChanged -= OnCollectionChanged;
                table.RowChanged -= OnDataChanged;
                table.RowChanging -= OnDataChanging;
                table.RowDeleted -= OnDataChanged;
                table.RowDeleting -= OnDataChanging;
                table.ColumnChanged -= OnColumnChanged;
                table.ColumnChanging -= OnColumnChanging;
                table.TableCleared -= OnDataChanged;
                table.TableClearing -= OnDataChanging;
            }

            private void OnCollectionChanged(object? sender, CollectionChangeEventArgs e) => Invalidate();

            private void OnDataChanged(object sender, DataRowChangeEventArgs e) => Invalidate();

            private void OnDataChanging(object sender, DataRowChangeEventArgs e) => Invalidate();

            private void OnDataChanged(object sender, DataTableClearEventArgs e) => Invalidate();

            private void OnDataChanging(object sender, DataTableClearEventArgs e) => Invalidate();

            private void OnColumnChanged(object sender, DataColumnChangeEventArgs e) => Invalidate();

            private void OnColumnChanging(object sender, DataColumnChangeEventArgs e) => Invalidate();

            private void Invalidate() {
                if (!IsValid) {
                    return;
                }

                IsValid = false;
                _invalidate();
            }

            public void Dispose() {
                if (_disposed) {
                    return;
                }

                _disposed = true;
                if (_subscribed) {
                    try {
                        Unsubscribe(_dataSet);
                    } catch {
                    }
                }
            }
        }
    }
}
