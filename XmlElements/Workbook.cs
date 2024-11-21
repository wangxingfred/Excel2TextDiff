//**************************************************************************************
//Create By Fred on 2020/12/15.
//
//@Description 一个表格文件
//**************************************************************************************

using System.Xml;
using System.Collections.Generic;

namespace XmlElements
{
    public class Workbook : Element
    {
        public List<Worksheet> Worksheets { get; private set; }

        internal override bool Parse(XmlElement xmlElement)
        {
            if (!base.Parse(xmlElement)) return false;

            Worksheets = ParseChildren<Worksheet>(xmlElement);

            return true;
        }

        public Worksheet GetWorksheet(string name)
        {
            foreach (var worksheet in Worksheets)
            {
                if (worksheet.Name == name) return worksheet;
            }

            return null;
        }
    }
}