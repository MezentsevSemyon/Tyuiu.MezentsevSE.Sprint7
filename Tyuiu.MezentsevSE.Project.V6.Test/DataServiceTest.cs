using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
namespace Tyuiu.MezentsevSE.Project.V6.Test
{
    [TestClass]
    public class DataServiceTest
    {
        [TestMethod]
        public void TestMethod1()
        {
            string path = @"C:\source\repos\Tyuiu.MezentsevSE.Sprint7\Files\OutPutDataFile.csv";

            FileInfo fileinfo = new FileInfo(path);
            bool fileExists = fileinfo.Exists;
            bool wait = true;
            Assert.AreEqual(wait, fileExists);
        }
    }
}
