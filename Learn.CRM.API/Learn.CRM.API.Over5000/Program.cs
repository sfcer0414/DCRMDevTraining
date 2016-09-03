﻿using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Learn.CRM.API.Over5000
{
    class Program
    {
        static void Main(string[] args)
        {
            //記得引用 System.Configuration.dll 系統元件，才能透過此方法取得在.config中的連線字串
            var connString = ConfigurationManager.ConnectionStrings["MyCRMServer"].ConnectionString;

            var over5000SampleCode = new Over5000SampleCode(connString);
            over5000SampleCode.onLog += SampleCode_onLog;

            over5000SampleCode.Init1wAccountData();
            over5000SampleCode.Query1wAccountData();
            over5000SampleCode.Remove1wAccountData();

#if DEBUG
            System.Console.Write("press any key to finish...");
            var inChar = System.Console.ReadKey();
#endif
        }

        private static void SampleCode_onLog(string message)
        {
            System.Console.WriteLine(message);
        }
    }
}
