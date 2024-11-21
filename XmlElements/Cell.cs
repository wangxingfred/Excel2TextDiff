//**************************************************************************************
//Create By Fred on 2020/12/15.
//
//@Description 一个单元格的属性（样式等）和数据
//**************************************************************************************

using System.Globalization;
using System.Xml;

namespace XmlElements
{
    public enum CellType
    {
        Blank,
        Number,
        String,
    }

    public class Cell : Element
    {
        public static readonly Cell EmptyCell = new();

        public Data Data { get; private set; }

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

        public CellType Type
        {
            get => _type;
            internal set
            {
                _type = value;
                Data.Type = value == CellType.Number ? DataType.Number : DataType.String;
            }
        }

        private uint _index;
        private CellType _type;

        protected override void OnCreate()
        {
            base.OnCreate();

            Data = Create<Data>(this);
        }

        internal override bool Parse(XmlElement xmlElement)
        {
            if (!base.Parse(xmlElement)) return false;

            Data = ParseSingleChild<Data>(xmlElement) ?? CreateWithoutLink<Data>(this);

            IsEmpty = Data.IsEmpty;

            if (IsEmpty)
            {
                _type = CellType.Blank;
            }
            else if (Data.Type == DataType.Number)
            {
                _type = CellType.Number;
            }
            else
            {
                _type = CellType.String;
            }

            _index = QueryAttrUnsigned("ss:Index");

            return true;
        }

        public string ReadStr()
        {
            return IsEmpty ? string.Empty : Data.GetText();
        }

        public int ReadInt32()
        {
            return _type == CellType.Number ? int.Parse(Data.GetText()) : 0;
        }

        public long ReadInt64()
        {
            return _type == CellType.Number ? long.Parse(Data.GetText()) : 0;
        }

        public double ReadDouble()
        {
            return _type == CellType.Number ? double.Parse(Data.GetText()) : 0;
        }

        /// <summary>
        /// 设置单元格类型为Number，并设置为double数值
        /// </summary>
        /// <param name="d"></param>
        public void WriteDouble(double d)
        {
            Data.Type = DataType.Number;
            Data.SetText(d.ToString(CultureInfo.InvariantCulture));
            IsEmpty = Data.IsEmpty;
        }

        /// <summary>
        /// 保持单元格类型不变，设置文本内容
        /// </summary>
        /// <param name="str"></param>
        public void WriteStr(string str)
        {
            Data.SetText(str);
            IsEmpty = Data.IsEmpty;
        }

        /// <summary>
        /// 设置单元格类型为Number，并设置为int值
        /// </summary>
        /// <param name="value"></param>
        public void WriteInt32(int value)
        {
            Data.Type = DataType.Number;
            Data.SetText(value.ToString());
            IsEmpty = Data.IsEmpty;
        }

        /// <summary>
        /// 设置单元格类型为Number，并设置为long值
        /// </summary>
        /// <param name="value"></param>
        public void WriteInt64(long value)
        {
            Data.Type = DataType.Number;
            Data.SetText(value.ToString());
            IsEmpty = Data.IsEmpty;
        }

        /// <summary>
        /// 设置单元格类型为Number，并设置为数字（以字符串形式提供，比如"22"、"55.8"等）
        /// </summary>
        /// <param name="value"></param>
        public void WriteNumber(string value)
        {
            Data.Type = DataType.Number;
            Data.SetText(value);
            IsEmpty = Data.IsEmpty;
        }
    }
}