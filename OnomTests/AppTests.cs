using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;



namespace Xyqlx.OneNote.Tests
{
    [TestClass()]
    public class AppTests
    {
        public List<Notebook> Notebooks { get; set; }
        public List<Section> Sections { get; set; }

        [TestMethod]
        public void GetNotebooksTest()
        {
            Notebooks = App.Notebooks?.ToList();
        }

        [TestMethod]
        public void GetSectionsTest()
        {
            GetNotebooksTest();
            Sections = Notebooks?.Where(x=>x.Name=="xyals")?.FirstOrDefault()?.Sections?.ToList();
        }

        [TestMethod]
        public void CreatePageTest()
        {
            GetSectionsTest();
            Section section = Sections?.Where(x=>x.Name=="test")?.FirstOrDefault();
            var startPageNums = section?.PageInfos?.Count();
            for (var i = 0; i < 100; ++i)
            {
                section?.AddPage();
            }
            GetSectionsTest();
            section = Sections?.Where(x => x.Name == "test")?.FirstOrDefault();
            var endPageNums = section?.PageInfos?.Count();
            Assert.AreEqual(100, endPageNums - startPageNums);
        }

        [TestMethod]
        public void DeletePageTest()
        {
            GetSectionsTest();
            var section = Sections?.Where(x => x.Name == "test")?.FirstOrDefault();
            section.Pages.Clear();
            Assert.AreEqual(0, Sections?.Where(x => x.Name == "test")?.FirstOrDefault()?.PageInfos?.Count());
        }
    }
}
