using System;
using System.Configuration;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Learn.CRM.API.Workflow
{
    class Program
    {
        static void Main(string[] args)
        {
            //記得引用 System.Configuration.dll 系統元件，才能透過此方法取得在.config中的連線字串
            var connString = ConfigurationManager.ConnectionStrings["MyCRMServer"].ConnectionString;

            var workflowSampleCode = new WorkflowSampleCode(connString);
            workflowSampleCode.onLog += SampleCode_onLog;
            workflowSampleCode.ExecuteSampleWorkflow();
            workflowSampleCode.ExecuteSampleAction();

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
