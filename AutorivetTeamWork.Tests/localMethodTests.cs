using Microsoft.VisualStudio.TestTools.UnitTesting;
using GoumangToolKit;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.CSharp;

namespace GoumangToolKit.Tests
{
    [TestClass()]
    public class localMethodTests
    {
        [TestMethod()]
        public void GetConfigValueTest()
        {
          dynamic CPNameDic = localMethod.GetConfigValue("GetCouponName", "CouponCfg.py");

            string fd = CPNameDic("l1");
            string dd2= CPNameDic("");
            Assert.AreEqual(dd2, null);
            Assert.AreEqual(fd, "SKIN-70/FRM-50");
           

            //  Assert.Fail();
        }
        [TestMethod()]
        public void GetConfigValueTest2()
        {
            dynamic CPNameDic = localMethod.GetConfigValue("InfoPath");

            string fd = CPNameDic;
            Assert.AreEqual(fd, @"\\192.168.3.32\Autorivet\Prepare\INFO\");

            //  Assert.Fail();
        }
    }
}