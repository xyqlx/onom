using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace Xyqlx.OneNote
{
    public class Page : OneNoteObject, IHasName
    {
        public Page(string name, int pagelevel)
        {
            root = new XElement(one + "Page",
                new XAttribute("name", name),
                new XAttribute("pageLevel", pagelevel.ToString()),
                new XAttribute(XNamespace.Xmlns + "one", one),
                new XElement(one + "Title", new XElement(one + "OE", new XElement(one + "T", name)))
                );
        }
        public Page(string name, int pagelevel, string content) : this(name, pagelevel)
        {
            AddLastOutlineText(content);
        }
        public Page(string xml) : base(xml)
        {
        }
        public Page(XElement element) : base(element) { }
        public string Title
        {
            get => root.Element(one + "Title").Element(one + "OE").Element(one + "T").Value;
            set => root.Element(one + "Title").Element(one + "OE").Element(one + "T").Value = value;
        }
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
        
        /// <summary>
        /// 返回去除了所有ID的副本
        /// </summary>
        /// <returns></returns>
        public Page Copy()
        {
            var el = new XElement(root);
            foreach (var d in el.DescendantsAndSelf())
            {
                var id = d.Attribute("ID") ?? d.Attribute("objectID");
                if (id != null)
                    id.Remove();
            }
            return new Page(el);
        }
        /// <summary>
        /// 在最后一行添加文本
        /// </summary>
        /// <param name="text"></param>
        public void AddLastOutlineText(string text)
        {
            var lastOutline = root.Elements(one + "Outline").LastOrDefault();
            if (lastOutline == null)
                root.Add(new XElement(one + "Outline", new XElement(one + "OEChildren")));
            lastOutline = root.Elements(one + "Outline").Last();
            lastOutline.Elements(one + "OEChildren").Last().Add(new XElement(one + "OE", new XElement(one + "T", text)));
        }
        /// <summary>
        /// 返回以纯文本形式存在的文字
        /// </summary>
        public string PlainText
        {
            get
            {
                List<string> list = new List<string>();
                foreach (var outline in root.Elements(one + "Outline"))
                    foreach (var oe in outline.Element(one + "OEChildren").Elements(one + "OE"))
                    {
                        foreach (var t in oe.Descendants(one + "T"))
                            list.Add(Regex.Replace(t.Value.Replace('\n', ' '), "<[^>]*>", string.Empty));
                        list.Add("\n");
                    }
                return string.Join(string.Empty, list);
            }
        }
        /// <summary>
        /// 试着修复一些内容上的错误
        /// </summary>
        public void Fix(Exception exception)
        {
            if(exception.HResult == (int)Microsoft.Office.Interop.OneNote.Error.hrInvalidXML)
            {
                foreach (var outline in root.Elements(one + "Outline"))
                {
                    var indents = outline.Element(one + "Indents");
                    if(indents != null)
                    {
                        foreach(var indent in indents.Elements())
                        {
                            string s = indent.Attribute("indent").Value;
                            if (s.Contains("E"))
                                indent.Attribute("indent").Value = "0.0";
                        }
                    }
                }
            }
        }
        /// <summary>
        /// 应用内容的更改
        /// </summary>
        public void Update() => App.UpdatePage(this);
        public void Remove() => App.application.DeleteHierarchy(this.ID);
    }
    /// <summary>
    /// 对页面MetaData的Lookup
    /// </summary>
    public class PageMetaDataLookup : ILookup<string, string>
    {
        private readonly Page page;

        Lookup<string, string> lookup;

        public PageMetaDataLookup(Page page)
        {
            this.page = page;
            lookup = (Lookup<string, string>)page.root.Elements(page.one + "Meta").ToLookup(x => x.Attribute("name").Value, x => x.Attribute("content").Value);
        }

        public IEnumerable<string> this[string key] => lookup[key];

        public int Count => lookup.Count;

        public bool Contains(string key)
        {
            return lookup.Contains(key);
        }

        public IEnumerator<IGrouping<string, string>> GetEnumerator()
        {
            return lookup.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
    /// <summary>
    /// 页面MetaData的Dictionary（只允许name出现一次）
    /// </summary>
    public class PageMetaDataDictionary : IDictionary<string, string>
    {
        private readonly Page page;

        Dictionary<string, string> data;

        public PageMetaDataDictionary(Page page)
        {
            this.page = page;
            data = page.root.Elements(page.one + "Meta").ToDictionary(x => x.Attribute("name").Value, x => x.Attribute("content").Value);
        }

        public string this[string key]
        {
            get => data[key]; set
            {
                if (data.ContainsKey(key))
                {
                    page.root.Elements(page.one + "Meta").Where(x => x.Attribute("name").Value == key).Last().Attribute("content").Value = value;
                    data[key] = value;
                }
                else this.Add(key, value);
            }
        }

        public ICollection<string> Keys => data.Keys;

        public ICollection<string> Values => data.Values;

        public int Count => data.Count;

        public bool IsReadOnly => false;

        public void Add(string key, string value)
        {
            data.Add(key, value);
            var place =
                (from el in page.root.Elements()
                 where el.Name == page.one + "TagDef"
                 || el.Name == page.one + "QuickStyleDef"
                 select el).LastOrDefault();
            var meta = new XElement(page.one + "Meta", new XAttribute("name", key), new XAttribute("content", value));
            if (place == null)
                page.root.AddFirst(meta);
            else
                place.AddAfterSelf(meta);
        }

        public void Add(KeyValuePair<string, string> item)
        {
            this.Add(item.Key, item.Value);
        }

        public void Clear()
        {
            page.root.Elements(page.one + "Meta").Remove();
            data.Clear();
        }

        public bool Contains(KeyValuePair<string, string> item)
        {
            return data.Contains(item);
        }

        public bool ContainsKey(string key)
        {
            return data.ContainsKey(key);
        }

        public void CopyTo(KeyValuePair<string, string>[] array, int arrayIndex)
        {
            throw new System.NotImplementedException();
        }

        public IEnumerator<KeyValuePair<string, string>> GetEnumerator()
        {
            throw new System.NotImplementedException();
        }

        public bool Remove(string key)
        {
            if (!data.ContainsKey(key))
                return false;
            page.root.Elements(page.one + "Meta").Where(x => x.Attribute("name").Value == key).Remove();
            return data.Remove(key);
        }

        public bool Remove(KeyValuePair<string, string> item)
        {
            if (data.Contains(item))
            {
                this.Remove(item.Key);
                return true;
            }
            else return false;
        }

        public bool TryGetValue(string key, out string value)
        {
            return data.TryGetValue(key, out value);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            throw new System.NotImplementedException();
        }
    }
}
