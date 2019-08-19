using System.Collections.Generic;
using System.Linq;

using System.Xml.Linq;
namespace Xyqlx.OneNote
{
    public class SectionGroup : OneNoteObject, IHasName, IHasPath
    {
        public SectionGroup(string xml) : base(xml) { }
        public SectionGroup(XElement element) : base(element) { }
        public string Name
        {
            get => root.Attribute("name").Value;
            set => root.Attribute("name").Value = value;
        }
        public string Path
        {
            get => root.Attribute("path").Value;
        }
        public bool IsRecycleBin
        {
            get => root.Attribute("isRecycleBin") != null;
        }
        public IEnumerable<Section> Sections =>
            from el in root.Elements(one + "Section")
            select App.GetSection(new Section(el).ID);
        public IEnumerable<SectionGroup> SectionGroups =>
            from el in root.Elements(one + "SectionGroup")
            select new SectionGroup(el);
        public IEnumerable<Section> AllSections
        {
            get
            {
                foreach (var s in Sections)
                    yield return s;
                foreach (var g in SectionGroups)
                    foreach (var s in g.AllSections)
                        yield return s;
            }
        }
    }
}
