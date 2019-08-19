using System.Collections.Generic;
using System.Linq;

namespace Xyqlx.OneNote
{
    public class Notebook : OneNoteObject, IHasName, IHasColor, IHasPath
    {
        public Notebook(string xml) : base(xml) { }
        public string Name
        {
            get => root.Attribute("name").Value;
            set => root.Attribute("name").Value = value;
        }
        public string NickName
        {
            get => root.Attribute("nickname").Value;
            set => root.Attribute("nickname").Value = value;
        }
        public string Path
        {
            get => root.Attribute("path").Value;
        }
        public System.Drawing.Color Color
        {
            get => System.Drawing.ColorTranslator.FromHtml(root.Attribute("color").Value);
            set => root.Attribute("color").Value = System.Drawing.ColorTranslator.ToHtml(value);
        }
        public IEnumerable<Section> Sections =>
            from el in root.Elements(one + "Section")
            select App.GetSection(new Section(el).ID);
        public IEnumerable<SectionGroup> SectionGroups =>
            (from el in root.Elements(one + "SectionGroup")
             select new SectionGroup(el)).Where(x => !x.IsRecycleBin);
        /// <summary>
        /// 包括所有节组中的节
        /// </summary>
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
