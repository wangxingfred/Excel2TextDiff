//**************************************************************************************
//Create By Fred on 2020/12/15.
//
//@Description excel基本元素
//**************************************************************************************

using System;
using System.Xml;
using System.Collections.Generic;

namespace XmlElements
{
    public abstract class Element
    {
        protected const string NAMESPACE_URI = "urn:schemas-microsoft-com:office:spreadsheet";
        protected XmlElement XmlParent { get; private init; }
        protected XmlElement XmlElement { get; private set; }

        public bool IsEmpty { get; protected set; } = true;

        public bool IsDirty { get; protected set; }

        internal virtual bool Parse(XmlElement xmlElement)
        {
            XmlElement = xmlElement;
            IsEmpty = xmlElement == null;
            return !IsEmpty;
        }

        protected virtual void OnCreate()
        {
        }

        protected virtual void OnDestroy()
        {
            XmlParent.RemoveChild(XmlElement);
        }

        protected string QueryAttrString(string name)
        {
            return XmlElement.GetAttribute(name);
        }

        protected bool QueryAttrBool(string name)
        {
            var attr = XmlElement.GetAttribute(name);
            return bool.TryParse(attr, out var value) && value;
        }

        protected int QueryAttrInt(string name)
        {
            var attr = XmlElement.GetAttribute(name);
            return int.TryParse(attr, out var value) ? value : 0;
        }

        protected uint QueryAttrUnsigned(string name)
        {
            var attr = XmlElement.GetAttribute(name);
            return uint.TryParse(attr, out var value) ? value : 0;
        }

        protected double QueryAttrDouble(string name)
        {
            var attr = XmlElement.GetAttribute(name);
            return double.TryParse(attr, out var value) ? value : 0.0;
        }


        internal static T ParseSingleChild<T>(XmlElement parent) where T : Element, new()
        {
            var type = typeof(T);

            var childNodes = parent.GetElementsByTagName(type.Name);
            var childCount = childNodes.Count;

            for (int i = 0; i < childCount; i++)
            {
                var childNode = childNodes[i];
                if (childNode.NodeType == XmlNodeType.Element)
                {
                    var child = new T();
                    child.Parse(childNode as XmlElement);
                    return child;
                }
            }

            return null;
        }

        internal static List<T> ParseChildren<T>(XmlElement parent, Action<T> callback = null)
            where T : Element, new()
        {
            var type = typeof(T);

            var childNodes = parent.GetElementsByTagName(type.Name);
            var childCount = childNodes.Count;
            var children = new List<T>(childCount);
            for (var i = 0; i < childCount; i++)
            {
                var childNode = childNodes[i];
                if (childNode.NodeType == XmlNodeType.Element)
                {
                    var child = new T {XmlParent = parent};
                    child.Parse(childNode as XmlElement);

                    callback?.Invoke(child);

                    children.Add(child);
                }
            }

            return children;
        }

        internal static T Create<T>(Element parent) where T : Element, new()
        {
            var type = typeof(T);

            var xmlParent = parent.XmlElement;
            var xmlElement = xmlParent.OwnerDocument.CreateElement("", type.Name, NAMESPACE_URI);
            xmlParent.AppendChild(xmlElement);

            var instance = new T
            {
                XmlParent = xmlParent,
                XmlElement = xmlElement
            };
            instance.OnCreate();

            return instance;
        }
        
        internal static T CreateWithoutLink<T>(Element parent) where T : Element, new()
        {
            var type = typeof(T);

            var xmlParent = parent.XmlElement;
            var xmlElement = xmlParent.OwnerDocument.CreateElement("", type.Name, NAMESPACE_URI);
            // AppendChild会带来相对较大的性能损耗，所以这里不做AppendChild，在需要写入的时候再做AppendChild
            // xmlParent.AppendChild(xmlElement);

            var instance = new T
            {
                XmlParent = xmlParent,
                XmlElement = xmlElement
            };
            instance.OnCreate();

            return instance;
        }

        internal static void Destroy(Element instance)
        {
            instance.OnDestroy();
        }
    }
}