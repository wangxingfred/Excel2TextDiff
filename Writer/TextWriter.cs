using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace Excel2TextDiff
{
    public class TextWriter : IVisitor
    {
        /// <summary>
        /// 输出一行数据时，单元格之间的分隔符
        /// </summary>
        public string Separator { get; init; }

        /// <summary>
        /// 从第几行开始对齐
        /// </summary>
        public int AlignBeginRow { get; init; }

        private readonly StringBuilder _sb = new();

        /// <summary>
        /// 需要处理对齐时，记录每一列的最大宽度
        /// </summary>
        private readonly List<int> _maxWidths = new();

        private struct Cell
        {
            public string _value;
            public int _width;
        }

        /// <summary>
        /// 一个页签(sheet)的所有行数据
        /// </summary>
        private readonly List<List<Cell>> _rows = new();

        private static readonly LinkedList<List<Cell>> _pool = new();

        public void BeginSheet(string sheetName)
        {
            _sb.Append($"\n==========  [{sheetName}]   ==========");
            _rows.Clear();
            _maxWidths.Clear();
        }

        public void VisitRow(List<string> row)
        {
            // 只保留到最后一个非空白单元格
            var lastNotEmptyIndex = row.FindLastIndex(s => !string.IsNullOrEmpty(s));

            // 忽略空白行，没必要diff这个
            if (lastNotEmptyIndex < 0) return;

            var rowIndex = _rows.Count + 1;
            var shouldAlign = ShouldAlign(rowIndex);

            var cells = GetFromPool();
            for (var i = 0; i < row.Count; i++)
            {
                if (i > lastNotEmptyIndex) break;

                var str = row[i];
                var totalCnt = str.Length;
                var chineseCnt = GetChineseCnt(str);
                var englishCnt = totalCnt - chineseCnt;
                var width = englishCnt + 2 * chineseCnt;

                if (shouldAlign)
                {
                    if (i < _maxWidths.Count)
                    {
                        _maxWidths[i] = Math.Max(width, _maxWidths[i]);
                    }
                    else
                    {
                        _maxWidths.Add(width);
                    }
                }

                cells.Add(new Cell {_value = str, _width = width});
            }

            _rows.Add(cells);
        }

        public void EndSheet()
        {
            var rowIndex = 0;
            foreach (var cells in _rows)
            {
                ++rowIndex;
                var shouldAlign = ShouldAlign(rowIndex);

                _sb.Append('\n');

                for (var j = 0; j < cells.Count; j++)
                {
                    var cell = cells[j];
                    var value = cell._value;

                    if (j > 0) _sb.Append(Separator);
                    _sb.Append(value);

                    if (shouldAlign && j < _maxWidths.Count)
                    {
                        var maxWidth = _maxWidths[j];
                        var paddingCnt = maxWidth - cell._width;

                        if (paddingCnt > 0)
                        {
                            _sb.Append(' ', paddingCnt);
                        }
                    }
                }

                ReturnToPool(cells);
            }

            _sb.Append('\n');
        }

        public void Save(string outputTextFile)
        {
            System.IO.File.WriteAllText(outputTextFile, _sb.ToString(), Encoding.UTF8);
        }

        private bool ShouldAlign(int rowIndex)
        {
            return AlignBeginRow != 0 && rowIndex >= AlignBeginRow;
        }


        private static int GetChineseCnt(string str)
        {
            using var cEnumerator = str.GetEnumerator();
            var regex = new Regex("^[\u4E00-\u9FA5]{0,}$");
            var cnt = 0;
            while (cEnumerator.MoveNext())
            {
                if (regex.IsMatch(cEnumerator.Current.ToString(), 0))
                    cnt++;
            }

            return cnt;
        }

        private static List<Cell> GetFromPool()
        {
            var node = _pool.First;
            if (node == null) return new List<Cell>();

            _pool.RemoveFirst();
            return node.Value;
        }

        private static void ReturnToPool(List<Cell> list)
        {
            list.Clear();
            _pool.AddLast(list);
        }
    }
}