using System.Xml.Linq;
namespace Xyqlx.OneNote
{
    public class PageInfo : OneNoteObject, IHasName
    {
        public PageInfo(Page page)
        {
            root = page.root;
        }
        public PageInfo(string name, int pagelevel)
        {
            root = new XElement(one + "Page",
                new XAttribute("name",name),
                new XAttribute("pageLevel",pagelevel.ToString())
                ,new XAttribute(XNamespace.Xmlns + "one", one));
        }
        public PageInfo(string xml) : base(xml) { }
        public PageInfo(XElement element) : base(element) { }
        public string Name
        {
            get => root.Attribute("name").Value;
            set => root.Attribute("name").Value = value;
        }
        public int PageLevel
        {
            get => int.Parse(root.Attribute("pageLevel").Value);
            set => root.Attribute("pageLevel").Value = value.ToString();
        }
    }
}
