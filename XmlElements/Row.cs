//**************************************************************************************
//Create By Fred on 2020/12/15.
//
//@Description 一行数据
//**************************************************************************************

using System.Collections.Generic;
using System.Xml;

namespace XmlElements
{
    public class Row : Element
    {
        private uint _index;

        public uint Index
        {
            get => _index;
            internal set
            {
                if (_index == value) return;
                _index = value;
                XmlElement.RemoveAttribute("ss:Index");
            }
        }

        public List<Cell> Cells { get; private set; }

        public List<Cell> FilledCells { get; private set; }

        public int AllCellCount => FilledCells.Count;

        public Cell GetCell(int col)
        {
            return col < FilledCells.Count ? FilledCells[col] : Cell.EmptyCell;
        }

        internal List<Cell> ResizeCells(int cellsCount, System.Action<Cell, int> setCell = null)
        {
            var cells = Cells;
            if (cells.Capacity < cellsCount) cells.Capacity = cellsCount;

            for (var i = 0; i < cellsCount; i++)
            {
                Cell cell;
                if (i < cells.Count)
                {
                    cell = cells[i];
                }
                else
                {
                    cell = Create<Cell>(this);
                    cells.Add(cell);
                }

                cell.Index = (uint)i;
                setCell?.Invoke(cell, i);
            }

            // 移除多余的cell
            var lastIndex = cells.Count - 1;
            while (lastIndex >= cellsCount)
            {
                var cell = cells[lastIndex];
                cells.RemoveAt(lastIndex);
                Destroy(cell);
                --lastIndex;
            }

            FilledCells = cells;
            return cells;
        }

        protected override void OnCreate()
        {
            Cells = new List<Cell>();
            FilledCells = Cells;
        }

        internal override bool Parse(XmlElement xmlElement)
        {
            if (!base.Parse(xmlElement)) return false;

            Index = QueryAttrUnsigned("ss:Index");

            InitCells();

            return true;
        }

        private void InitCells()
        {
            Cells = ParseChildren<Cell>(XmlElement);

            if (Cells.Count == 0)
            {
                FilledCells = Cells;
                IsEmpty = true;
                return;
            }

            var capacity = Cells[^1].Index;
            FilledCells = new List<Cell>((int)capacity);

            var emptyRow = true;
            uint i = 1;

            // 在有数据的Cells中间填充一些空白cell，形成一个填充满的cell数组：FilledCells
            foreach (var cell in Cells)
            {
                if (!cell.IsEmpty) emptyRow = false;

                if (cell.Index == i)
                {
                    FilledCells.Add(cell);
                    ++i;
                    continue;
                }

                if (cell.Index == 0)
                {
                    // xml表格中没有记录顺序递增的索引，读出来就为0，这里把正确的值设置进去
                    cell.Index = i;
                }
                else
                {
                    // 如果有空的格子，则会产生跳跃的索引，这里要把中间跳过的空格子补上
                    for (; i < cell.Index; ++i) {
                        FilledCells.Add(Cell.EmptyCell);
                    }
                }

                FilledCells.Add(cell);
                ++i;
            }

            IsEmpty = emptyRow;
        }

    }
}