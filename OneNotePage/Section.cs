using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace Xyqlx.OneNote
{
    public class Section : OneNoteObject, IHasName, IHasPath, IHasColor
    {
        public Section(string xml) : base(xml) { }
        public Section(XElement element) : base(element) { }
        public string Name
        {
            get => root.Attribute("name").Value;
            set => root.Attribute("name").Value = value;
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
        public bool IsInRecycleBin
        {
            get => root.Attribute("isInRecycleBin") != null;
        }
        public bool IsDeletedPages
        {
            get => root.Attribute("isDeletedPages") != null;
        }
        public IEnumerable<PageInfo> PageInfos =>
            from el in root.Elements(one + "Page")
            select new PageInfo(el);
        public PageList Pages => new PageList(this);
        /// <summary>
        /// 应用结构上的更改
        /// </summary>
        public void Update() => App.UpdateSection(this);
        /// <summary>
        /// 在节的最后添加新页
        /// </summary>
        /// <returns></returns>
        public Page AddPage() => App.AddPage(this);
    }
    /// <summary>
    /// 适用于单节的页面List，它会立即应用节的结构更改以及与之相关的页面更改
    /// </summary>
    public class PageList : IList<Page>
    {
        private readonly Section section;
        /// <summary>
        /// 上传更改结构并下载
        /// </summary>
        private void Update()
        {
            App.UpdateSection(section);
            section.root = App.GetSection(section.ID).root;
        }
        /// <summary>
        /// 下载新的结构，当在类外进行了结构修改时使用
        /// </summary>
        public void Refresh()
        {
            section.root = App.GetSection(section.ID).root;
        }

        public PageList(Section section)
        {
            this.section = section;
        }

        public Page this[int index]
        {
            get => App.GetPage(section.root.Elements().Skip(index).First().Attribute("ID").Value); set
            {
                if (index == 0)
                    section.root.AddFirst(value.root);
                else section.root.Elements().Skip(index).First().AddBeforeSelf(value.root);
                Update();
                var page = this[index];
                value.ID = page.ID;
                App.UpdatePage(value);
            }
        }

        public int Count => section.root.Elements().Count();

        public bool IsReadOnly => false;

        /// <summary>
        /// 添加页面
        /// </summary>
        /// <param name="item">如果是现有的页面，建议调用其Copy()方法</param>
        public void Add(Page item)
        {
            var page = App.AddPage(section);
            item.ID = page.ID;
            try
            {
                App.UpdatePage(item);
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                if (ex.ErrorCode == (int)Microsoft.Office.Interop.OneNote.Error.hrPageObjectDoesNotExist)
                {
                    try
                    {
                        var copy = item.Copy();
                        copy.ID = item.ID;
                        App.UpdatePage(copy);
                    }
                    catch (System.Exception e)
                    {
                        throw e;
                    }
                }
                else throw ex;
            }
            Refresh();
        }

        public void Clear()
        {
            App.application.DeleteHierarchy(section.ID);
            section.root.RemoveNodes();
            Update();
        }

        /// <summary>
        /// 页面的比较依据是ID>name，但不建议使用此方法
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        public bool Contains(Page item)
        {
            string id = item.ID;
            if (id != "")
                return section.PageInfos.Any(x => x.ID == id);
            else
            {
                string name = item.Name;
                return section.PageInfos.Any(x => x.Name == name);
            }
        }

        public void CopyTo(Page[] array, int arrayIndex)
        {
            array = this.Skip(arrayIndex).ToArray();
        }

        public IEnumerator<Page> GetEnumerator()
        {
            return
                (from info in section.PageInfos
                 select App.GetPage(info.ID)).GetEnumerator();
        }

        /// <summary>
        /// 页面的比较依据是ID>name，但不建议使用此方法
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        public int IndexOf(Page item)
        {
            string id = item.ID;
            if (id != "")
                return section.PageInfos.TakeWhile(x => x.ID != id).Count();
            else
            {
                string name = item.Name;
                return section.PageInfos.TakeWhile(x => x.Name != name).Count();
            }
        }

        public void Insert(int index, Page item)
        {
            if(index == 0)
                section.root.AddFirst(new XElement(section.one + "Page",new XAttribute("name",item.Name),new XAttribute("pageLevel",item.PageLevel)));
            else
                section.root.Elements().Skip(index).First().AddBeforeSelf(new XElement(section.one + "Page", new XAttribute("name", item.Name), new XAttribute("pageLevel", item.PageLevel)));
            Update();
            item.ID = this[index].ID;
            item.Update();
        }

        /// <summary>
        /// 不会检查移除的页面是否在本节内
        /// </summary>
        /// <param name="item"></param>
        /// <returns>总是返回true</returns>
        public bool Remove(Page item)
        {
            App.application.DeleteHierarchy(item.ID);
            Refresh();
            return true;
        }

        public void RemoveAt(int index)
        {
            App.application.DeleteHierarchy(this[index].ID);
            Refresh();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
