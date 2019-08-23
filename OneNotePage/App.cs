using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace Xyqlx.OneNote
{
    /// <summary>
    /// 封装了OneNote Application的静态类
    /// </summary>
    public static class App
    {
        public static readonly Microsoft.Office.Interop.OneNote.Application application;
        private static XElement root;
        public readonly static XNamespace one;
        static App()
        {
            application = new Microsoft.Office.Interop.OneNote.Application();
            application.GetHierarchy("", Microsoft.Office.Interop.OneNote.HierarchyScope.hsSections, out string xml);
            root = XElement.Parse(xml);
            one = root.GetNamespaceOfPrefix("one");
        }
        public static IEnumerable<Notebook> Notebooks =>
            from el in root.Elements(one + "Notebook")
            select new Notebook(el.ToString());
        public static Notebook GetNotebook(string id)
        {
            application.GetHierarchy(id, Microsoft.Office.Interop.OneNote.HierarchyScope.hsSections, out string xml);
            return new Notebook(xml);
        }
        public static SectionGroup GetSectionGroup(string id)
        {
            application.GetHierarchy(id, Microsoft.Office.Interop.OneNote.HierarchyScope.hsSections, out string xml);
            return new SectionGroup(xml);
        }
        public static Section GetSection(string id)
        {
            application.GetHierarchy(id, Microsoft.Office.Interop.OneNote.HierarchyScope.hsPages, out string xml);
            return new Section(xml);
        }
        public static Page GetPage(string id)
        {
            application.GetPageContent(id, out string xml, Microsoft.Office.Interop.OneNote.PageInfo.piBasic);
            return new Page(xml);
        }
        public static void UpdatePage(Page page)
        {
            try
            {
                application.UpdatePageContent(page.ToString());
            }catch(Exception ex)
            {
                page.Fix(ex);
                application.UpdatePageContent(page.ToString());
            }
            
        }
        public static void UpdateSection(Section section)
        {
            application.UpdateHierarchy(section.ToString());
        }
        public static Page AddPage(Section section)
        {
            application.CreateNewPage(section.ID, out string pageID);
            return GetPage(pageID);
        }
    }
}
