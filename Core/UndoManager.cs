using System;
using System.Collections.Generic;

namespace SpreadsheetApp.Core
{
    public class UndoManager
    {
        private record Edit(int Row, int Col, string? OldRaw, string? NewRaw);
        private abstract record UndoItem;
        private record Single(Edit E) : UndoItem;
        private record Bulk(List<Edit> Edits) : UndoItem;
        private record SheetAdd(int Index, string Name) : UndoItem;
        private record Composite(List<UndoItem> Items) : UndoItem;

        public readonly record struct BulkEdit(int Row, int Col, string? OldRaw, string? NewRaw);

        private readonly Stack<UndoItem> _undos = new();
        private readonly Stack<UndoItem> _redos = new();

        // Coalescing: merge rapid edits to the same cell into one undo item
        private const int MergeWindowMs = 1000; // time window to coalesce
        private long _lastSetTick = long.MinValue;
        private int _lastSetRow = -1;
        private int _lastSetCol = -1;

        public bool CanUndo => _undos.Count > 0;
        public bool CanRedo => _redos.Count > 0;

        public void Clear()
        {
            _undos.Clear();
            _redos.Clear();
            _lastSetTick = long.MinValue; _lastSetRow = _lastSetCol = -1;
        }

        public void RecordSet(int row, int col, string? oldRaw, string? newRaw)
        {
            if (oldRaw == newRaw) return;
            // Try to coalesce with the previous Single if same cell and within window
            bool merged = false;
            long now = Environment.TickCount64;
            if (_undos.Count > 0 && _redos.Count == 0 &&
                _lastSetRow == row && _lastSetCol == col && (now - _lastSetTick) <= MergeWindowMs)
            {
                var top = _undos.Peek();
                if (top is Single s)
                {
                    // Pop and push a new Single preserving the original OldRaw, updating NewRaw
                    _undos.Pop();
                    _undos.Push(new Single(new Edit(row, col, s.E.OldRaw, newRaw)));
                    merged = true;
                }
            }
            if (!merged)
            {
                _undos.Push(new Single(new Edit(row, col, oldRaw, newRaw)));
            }
            _redos.Clear();
            _lastSetTick = now; _lastSetRow = row; _lastSetCol = col;
        }

        public void RecordBulk(IEnumerable<(int row, int col, string? oldRaw, string? newRaw)> edits)
        {
            var list = new List<Edit>();
            foreach (var e in edits)
            {
                if (e.oldRaw == e.newRaw) continue;
                list.Add(new Edit(e.row, e.col, e.oldRaw, e.newRaw));
            }
            if (list.Count == 0) return;
            _undos.Push(new Bulk(list));
            _redos.Clear();
            _lastSetTick = long.MinValue; _lastSetRow = _lastSetCol = -1;
        }

        public void RecordSheetAdd(int index, string name)
        {
            _undos.Push(new SheetAdd(index, name));
            _redos.Clear();
            _lastSetTick = long.MinValue; _lastSetRow = _lastSetCol = -1;
        }

        public void RecordComposite(IEnumerable<BulkEdit> edits, (int index, string name)? sheetAdd)
        {
            var items = new List<UndoItem>();
            var list = new List<Edit>();
            foreach (var e in edits)
            {
                if (e.OldRaw == e.NewRaw) continue;
                list.Add(new Edit(e.Row, e.Col, e.OldRaw, e.NewRaw));
            }
            if (list.Count > 0) items.Add(new Bulk(list));
            if (sheetAdd.HasValue) items.Add(new SheetAdd(sheetAdd.Value.index, sheetAdd.Value.name));
            if (items.Count == 0) return;
            _undos.Push(new Composite(items));
            _redos.Clear();
            _lastSetTick = long.MinValue; _lastSetRow = _lastSetCol = -1;
        }

        public bool TryUndo(out int row, out int col, out string? raw)
        {
            row = col = 0; raw = null;
            if (!CanUndo) return false;
            var top = _undos.Peek();
            if (top is Single s)
            {
                _undos.Pop();
                _redos.Push(s);
                row = s.E.Row; col = s.E.Col; raw = s.E.OldRaw;
                _lastSetTick = long.MinValue; _lastSetRow = _lastSetCol = -1;
                return true;
            }
            return false;
        }

        public bool TryRedo(out int row, out int col, out string? raw)
        {
            row = col = 0; raw = null;
            if (!CanRedo) return false;
            var top = _redos.Peek();
            if (top is Single s)
            {
                _redos.Pop();
                _undos.Push(s);
                row = s.E.Row; col = s.E.Col; raw = s.E.NewRaw;
                _lastSetTick = long.MinValue; _lastSetRow = _lastSetCol = -1;
                return true;
            }
            return false;
        }

        public bool TryUndoComposite(out List<BulkEdit> edits, out (int index, string name)? sheetAdd)
        {
            edits = new List<BulkEdit>(); sheetAdd = null;
            if (!CanUndo) return false;
            var top = _undos.Peek();
            if (top is Composite comp)
            {
                _undos.Pop();
                _redos.Push(comp);
                foreach (var item in comp.Items)
                {
                    switch (item)
                    {
                        case Bulk b:
                            foreach (var e in b.Edits)
                                edits.Add(new BulkEdit(e.Row, e.Col, e.OldRaw, e.NewRaw));
                            break;
                        case SheetAdd sa:
                            sheetAdd = (sa.Index, sa.Name);
                            break;
                    }
                }
                _lastSetTick = long.MinValue; _lastSetRow = _lastSetCol = -1;
                return true;
            }
            return false;
        }

        public bool TryRedoComposite(out List<BulkEdit> edits, out (int index, string name)? sheetAdd)
        {
            edits = new List<BulkEdit>(); sheetAdd = null;
            if (!CanRedo) return false;
            var top = _redos.Peek();
            if (top is Composite comp)
            {
                _redos.Pop();
                _undos.Push(comp);
                foreach (var item in comp.Items)
                {
                    switch (item)
                    {
                        case Bulk b:
                            foreach (var e in b.Edits)
                                edits.Add(new BulkEdit(e.Row, e.Col, e.OldRaw, e.NewRaw));
                            break;
                        case SheetAdd sa:
                            sheetAdd = (sa.Index, sa.Name);
                            break;
                    }
                }
                _lastSetTick = long.MinValue; _lastSetRow = _lastSetCol = -1;
                return true;
            }
            return false;
        }

        public bool TryUndoSheetAdd(out int index, out string name)
        {
            index = -1; name = string.Empty;
            if (!CanUndo) return false;
            var top = _undos.Peek();
            if (top is SheetAdd sa)
            {
                _undos.Pop();
                _redos.Push(sa);
                index = sa.Index; name = sa.Name;
                _lastSetTick = long.MinValue; _lastSetRow = _lastSetCol = -1;
                return true;
            }
            return false;
        }

        public bool TryRedoSheetAdd(out int index, out string name)
        {
            index = -1; name = string.Empty;
            if (!CanRedo) return false;
            var top = _redos.Peek();
            if (top is SheetAdd sa)
            {
                _redos.Pop();
                _undos.Push(sa);
                index = sa.Index; name = sa.Name;
                _lastSetTick = long.MinValue; _lastSetRow = _lastSetCol = -1;
                return true;
            }
            return false;
        }

        public bool TryUndoBulk(out List<(int row, int col, string? raw)> edits)
        {
            edits = new List<(int, int, string?)>();
            if (!CanUndo) return false;
            var top = _undos.Peek();
            if (top is Bulk b)
            {
                _undos.Pop();
                _redos.Push(b);
                foreach (var e in b.Edits) edits.Add((e.Row, e.Col, e.OldRaw));
                _lastSetTick = long.MinValue; _lastSetRow = _lastSetCol = -1;
                return true;
            }
            return false;
        }

        public bool TryRedoBulk(out List<(int row, int col, string? raw)> edits)
        {
            edits = new List<(int, int, string?)>();
            if (!CanRedo) return false;
            var top = _redos.Peek();
            if (top is Bulk b)
            {
                _redos.Pop();
                _undos.Push(b);
                foreach (var e in b.Edits) edits.Add((e.Row, e.Col, e.NewRaw));
                _lastSetTick = long.MinValue; _lastSetRow = _lastSetCol = -1;
                return true;
            }
            return false;
        }
    }
}
