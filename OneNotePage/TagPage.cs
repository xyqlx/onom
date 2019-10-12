using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace Xyqlx.OneNote
{
    /// <summary>
    /// 带有Tag的页面，注意含义与OneNote中的标记语义不同
    /// </summary>
    public class TagPage : Page
    {
        public TagPage(Page page) : base(page.root)
        {
            Init();
        }

        public TagPage(string xml) : base(xml)
        {
            Init();
        }

        public TagPage(XElement element) : base(element)
        {
            Init();
        }

        public TagPage(string name, int pagelevel) : base(name, pagelevel)
        {
            Init();
        }

        public TagPage(string name, int pagelevel, string content) : base(name, pagelevel, content)
        {
            Init();
        }

        private void Init()
        {
            metaData = new PageMetaDataDictionary(this);
            tags = new Tags(this);
        }

        public PageMetaDataDictionary metaData;
        public Tags tags;

    }
    /// <summary>
    /// 适用于TagPage的标签集合
    /// </summary>
    public class Tags : ISet<string>
    {
        private readonly TagPage page;
        HashSet<string> tags;

        public Tags(TagPage page)
        {
            this.page = page;
            if (!page.metaData.ContainsKey("tags"))
                page.metaData.Add("tags", "");
            tags = new HashSet<string>(page.metaData["tags"].Split(';').Where(x => !String.IsNullOrEmpty(x)));
        }

        private void Update()
        {
            page.metaData["tags"] = String.Join(";", tags);
        }

        public int Count => tags.Count;

        public bool IsReadOnly => false;

        public bool Add(string item)
        {
            if (tags.Contains(item))
                return false;
            tags.Add(item);
            Update();
            return true;
        }

        public void Clear()
        {
            tags.Clear();
            Update();
        }

        public bool Contains(string item)
        {
            return tags.Contains(item);
        }

        public void CopyTo(string[] array, int arrayIndex)
        {
            tags.CopyTo(array, arrayIndex);
        }

        public void ExceptWith(IEnumerable<string> other)
        {
            tags.ExceptWith(other);
            Update();
        }

        public IEnumerator<string> GetEnumerator()
        {
            return tags.GetEnumerator();
        }

        public void IntersectWith(IEnumerable<string> other)
        {
            tags.IntersectWith(other);
            Update();
        }

        public bool IsProperSubsetOf(IEnumerable<string> other)
        {
            return tags.IsProperSubsetOf(other);
        }

        public bool IsProperSupersetOf(IEnumerable<string> other)
        {
            return tags.IsProperSupersetOf(other);
        }

        public bool IsSubsetOf(IEnumerable<string> other)
        {
            return tags.IsSubsetOf(other);
        }

        public bool IsSupersetOf(IEnumerable<string> other)
        {
            return tags.IsSupersetOf(other);
        }

        public bool Overlaps(IEnumerable<string> other)
        {
            return tags.Overlaps(other);
        }

        public bool Remove(string item)
        {
            if (!item.Contains(item))
                return false;
            tags.Remove(item);
            Update();
            return true;
        }

        public bool SetEquals(IEnumerable<string> other)
        {
            return tags.SetEquals(other);
        }

        public void SymmetricExceptWith(IEnumerable<string> other)
        {
            tags.SymmetricExceptWith(other);
            Update();
        }

        public void UnionWith(IEnumerable<string> other)
        {
            tags.UnionWith(other);
            Update();
        }

        void ICollection<string>.Add(string item)
        {
            tags.Add(item);
            Update();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
