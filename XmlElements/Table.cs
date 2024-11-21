//**************************************************************************************
//Create By Fred on 2020/12/15.
//
//@Description 一个表格页中所有格子的数据
//**************************************************************************************

using System;
using System.Collections.Generic;
using System.Xml;

namespace XmlElements
{
    public class Table : Element
    {
        public List<Column> Columns { get; private set; }
        public List<Row> Rows { get; private set; }

        public int LastRow => Rows.Count - 1;
        public int LastCol => _maxRowCellCnt - 1;

        // FIXME NOW
        private int _maxRowCellCnt;

        internal override bool Parse(XmlElement xmlElement)
        {
            if (!base.Parse(xmlElement)) return false;

            uint columnIndex = 1;
            Columns = ParseChildren<Column>(XmlElement, (column) =>
            {
                if (column.Index == 0 || column.Index < columnIndex)
                {
                    column.Index = columnIndex;
                    ++columnIndex;
                }
                else
                {
                    columnIndex = column.Index + 1;
                }
            });

            uint rowIndex = 1;
            Rows = ParseChildren<Row>(XmlElement, (row) =>
            {
                if (row.Index == 0 || row.Index < rowIndex)
                {
                    row.Index = rowIndex;
                    ++rowIndex;
                }
                else
                {
                    rowIndex = row.Index + 1;
                }

                _maxRowCellCnt = Math.Max(_maxRowCellCnt, row.AllCellCount);
            });

            return true;
        }

        internal void UpdateAttributes()
        {
            // var rowCount = Rows.Count == 0 ? 0 : Rows[^1].Index;
            // XmlElement.SetAttribute("ss:ExpandedRowCount", rowCount.ToString());
            // 在之前的配置表中如果存在属性 "ExpandedRowCount"，
            // 这里再次设置"ExpandedRowCount"会导致xml文件中存在2条"ExpandedRowCount"属性
            // 从而导致xml文件无法打开的问题，这个字段目前会被配置表裁剪工具裁剪，因此这里不再设置该属性
            // var rowCount = Rows.Count == 0 ? 0 : Rows[^1].Index;
            // XmlElement.SetAttribute("ss:ExpandedRowCount", rowCount.ToString());

            IsDirty = true;

            // var colCount = Columns.Count == 0 ? 0 : Columns[Columns.Count - 1].Index;
            // XmlElement.SetAttribute("ss:ExpandedColumnCount", colCount.ToString());
            // _maxRowCellCnt = (int)colCount;
        }

        /// <summary>
        /// 保存前的处理工作
        /// </summary>
        internal void PrepareSave()
        {
            foreach (var row in Rows)
            {
                foreach (var cell in row.Cells)
                {
                    cell.Data.PrepareSave();
                }
            }

            IsDirty = false;
        }

        public CellType GetCellType(int row, int col)
        {
            return GetCell(row, col).Type;
        }

        public string ReadStr(int row, int col)
        {
            return GetCell(row, col).ReadStr();
        }

        protected Cell GetCell(int row, int col)
        {
            return Rows[row].GetCell(col);
        }
    }
}