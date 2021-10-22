using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;


namespace UnitTestProject
{
    [TestClass]
    public class UnitTest1
    {
        [TestClass]
        public class DummyTest
        {
            private const string Expected = "Hello World!";
            [TestMethod]
            public void TestMethod1()
            {
                using (var sw = new StringWriter())
                {
                    Console.SetOut(sw);
                    UserUpdater.Class1.Main();

                    var result = sw.ToString().Trim();
                    Assert.AreEqual(Expected, result);
                }
            }
        }
    }
}
