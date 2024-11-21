//**************************************************************************************
//Create By Fred on 2020/12/15.
//
//@Description 一列的属性等信息
//**************************************************************************************

using System.Xml;

namespace XmlElements
{
    public class Column : Element
    {
        public uint Index { get; internal set; }

        internal override bool Parse(XmlElement xmlElement)
        {
            if (!base.Parse(xmlElement)) return false;

            Index = QueryAttrUnsigned("ss:Index");

            return true;
        }
    }
}