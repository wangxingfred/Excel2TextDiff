//**************************************************************************************
//Create By Fred on 2020/12/15.
//
//@Description 一个单元格的数据
//**************************************************************************************

using System;
using System.Xml;
using System.Globalization;

namespace XmlElements
{
    public enum DataType
    {
        Invalid,
        Number,
        String,
    }

    public class Data : Element
    {
        public DataType Type
        {
            get => _type;
            set
            {
                _type = value;
                if (_typeAttr != null)
                {
                    _typeAttr.Value = value == DataType.Number ? "Number" : "String";    
                }
            }
        }

        private DataType _type;

        private XmlAttribute _typeAttr;

        internal override bool Parse(XmlElement xmlElement)
        {
            if (!base.Parse(xmlElement)) return false;

            _typeAttr = xmlElement.GetAttributeNode("ss:Type");
            var type = _typeAttr.Value;
            switch (type)
            {
                case "Number":
                    _type = DataType.Number;
                    break;
                case "String":
                    _type = DataType.String;
                    break;
                default:
                    _type = DataType.Invalid;
                    break;
            }

            IsEmpty = xmlElement.IsEmpty || string.IsNullOrEmpty(xmlElement.InnerText);

            IsDirty = true;

            return true;
        }

        protected override void OnCreate()
        {
            base.OnCreate();
            
            // Append属性会带来相对较大的性能损耗，而且设置_typeAttr.Value也会有较大的性能损耗
            // 所以这里先不创建属性，在需要写入时(PrepareSave函数中)再创建
            // _typeAttr = XmlElement.OwnerDocument.CreateAttribute("", "Type", NAMESPACE_URI);
            // XmlElement.Attributes.Append(_typeAttr);
            
            Type = DataType.String;
            IsEmpty = true;
        }

        internal void PrepareSave()
        {
            if (!IsDirty) return;
            
            LinkToParent();
            FormatText();
            IsDirty = false;
        }

        internal string GetText()
        {
            return XmlElement.InnerText;
        }

        internal void SetText(string text)
        {
            if (text == XmlElement.InnerText)
            {
                return;
            }

            if (_type == DataType.Number && !TryParseDouble(text, out _))
            {
                throw new InvalidOperationException($"{text} is not number");
            }

            XmlElement.InnerText = text;
            IsEmpty = XmlElement.IsEmpty;
            IsDirty = true;
        }

        private void FormatText()
        {
            if (_type == DataType.Number)
            {
                //临时代码，解决新增Number单元时会生成换行符，导致excel打不开
                XmlElement.InnerText = XmlElement.InnerText;
                return;
            }
            if (_type != DataType.String) return;
            if (IsEmpty) return;

            XmlElement.InnerText = XmlElement.InnerText
                .Replace("&", "&amp;")
                .Replace("\n", "&#10;");

            // var str2 = System.Net.WebUtility.HtmlEncode(str);
            // .Replace("-", "&#45;")
            // .Replace("+", "&#43;")
            // .Replace("*", "&#42;")
            // .Replace("/", "&#47;")
            //
            // .Replace("&", "&amp;")
            // .Replace("\"", "&quot;")
            // .Replace("\'", "&apos;")
            // .Replace("<", "&lt;")
            // .Replace(">", "&gt;");
            // var str3 = str.Replace("\n", "&#10;");
            // var str3 = str2.Replace("--", "&#45;-");
        }
        
        private void LinkToParent()
        {

            if (XmlElement.ParentNode == null)
            {
                XmlParent.AppendChild(XmlElement);
            }

            if (_typeAttr == null && XmlElement.OwnerDocument != null)
            {
                _typeAttr = XmlElement.OwnerDocument.CreateAttribute("", "Type", NAMESPACE_URI);
                _typeAttr.Value = _type == DataType.Number ? "Number" : "String";

                XmlElement.Attributes.Append(_typeAttr);
            }
        }

        private const NumberStyles NUMBER_STYLE = NumberStyles.AllowLeadingSign |
                                                  // xml格式的excel无法解析带空白的数字，所以这里不允许有空白
                                                  // NumberStyles.AllowTrailingWhite | NumberStyles.AllowLeadingWhite |
                                                  NumberStyles.AllowExponent |
                                                  NumberStyles.AllowThousands;

        internal static bool TryParseDouble(string text, out double value)
        {
            const NumberStyles DOUBLE_STYLE = NUMBER_STYLE | NumberStyles.AllowDecimalPoint;
            return double.TryParse(text, DOUBLE_STYLE, NumberFormatInfo.CurrentInfo, out value);
        }

        internal static bool TryParseInt64(string text, out long value)
        {
            return long.TryParse(text, NUMBER_STYLE, NumberFormatInfo.CurrentInfo, out value);
        }
    }
}