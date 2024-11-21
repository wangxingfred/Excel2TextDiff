//**************************************************************************************
//Create By Fred on 2020/12/15.
//
//@Description 一个表格页，包含页属性和表格数据
//**************************************************************************************

using System.Xml;

namespace XmlElements
{
    public class Worksheet : Element
    {
        public string Name { get; private set; }
        public WorksheetOptions Options { get; private set; }
        public Table Table { get; private set; }

        public int LastRow => Table.LastRow;
        public int LastCol => Table.LastCol;

        internal override bool Parse(XmlElement xmlElement)
        {
            if (!base.Parse(xmlElement)) return false;

            Name = QueryAttrString("ss:Name");

            Options = ParseSingleChild<WorksheetOptions>(xmlElement);
            Table = ParseSingleChild<Table>(xmlElement);

            return true;
        }

        public Row GetRow(int row)
        {
            return row <= Table.LastRow ? Table.Rows[row] : null;
        }

        public bool CellExist(int row, int col)
        {
            return row <= Table.LastRow && col <= Table.LastCol;
        }

        public CellType GetCellType(int row, int col)
        {
            return Table.GetCellType(row, col);
        }

        public string ReadStr(int row, int col)
        {
            return Table.ReadStr(row, col);
        }
    }
}