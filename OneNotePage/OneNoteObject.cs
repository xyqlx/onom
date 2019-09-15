using System;
using System.Xml.Linq;

namespace Xyqlx.OneNote
{
    public class OneNoteObject
    {
        internal XElement root;
        internal XNamespace one;
        public OneNoteObject()
        {
            one = App.one;
        }
        public OneNoteObject(string xml)
        {
            root = XElement.Parse(xml);
            one = root.GetNamespaceOfPrefix("one");
        }
        public OneNoteObject(XElement element)
        {
            root = element;
            one = root.GetNamespaceOfPrefix("one");
        }
        public string ID
        {
            get
            {
                var id = root.Attribute("ID") ?? root.Attribute("objectID");
                if (id == null)
                    return "";
                return id.Value;
            }
            set
            {
                var id = root.Attribute("ID") ?? root.Attribute("objectID");
                //也许应该都
                if (id == null)
                    root.SetAttributeValue("ID", value);
                else
                    id.Value = value;
            }
        }
        public System.DateTime DateTime
        {
            get
            {
                var dateTime = root.Attribute("dateTime") ?? root.Attribute("creationTime");
                if (dateTime == null) return DateTime.MinValue;
                else return System.Xml.XmlConvert.ToDateTime(dateTime.Value, System.Xml.XmlDateTimeSerializationMode.Utc);
            }
            set
            {
                var dateTime = root.Attribute("dateTime") ?? root.Attribute("creationTime");
                if (dateTime != null)
                    dateTime.Value = value.ToString("yyyy-MM-ddTHH:mm:ss.fffZ");
            }
        }
        public System.DateTime LastModifiedTime
        {
            get
            {
                var dateTime = root.Attribute("lastModifiedTime");
                if (dateTime == null) return DateTime.MinValue;
                else return System.Xml.XmlConvert.ToDateTime(dateTime.Value, System.Xml.XmlDateTimeSerializationMode.Utc);
            }
            set
            {
                var dateTime = root.Attribute("lastModifiedTime");
                if (dateTime != null)
                    dateTime.Value = value.ToString("yyyy-MM-ddTHH:mm:ss.fffZ");
            }
        }
        public void Open(bool newWindow = false)
        {
            App.application.NavigateTo(this.ID);
        }
        public override string ToString()
        {
            return root.ToString();
        }
    }
}
