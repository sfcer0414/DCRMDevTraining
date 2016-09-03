using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Learn.CRM.API.EarlyBound
{
    class Program
    {
        static void Main(string[] args)
        {
            //記得引用 System.Configuration.dll 系統元件，才能透過此方法取得在.config中的連線字串
            var connString = ConfigurationManager.ConnectionStrings["MyCRMServer"].ConnectionString;

            var earlyBoundSampleCode = new EarlyBoundSampleCode(connString);
            earlyBoundSampleCode.onLog += SampleCode_onLog;

            earlyBoundSampleCode.SingleMultiple();
            earlyBoundSampleCode.RetrieveMultiple();
            earlyBoundSampleCode.AddProduct();
            earlyBoundSampleCode.UpdateProduct();
            earlyBoundSampleCode.Associate_Disassociate();
            earlyBoundSampleCode.PublishProduct();
            earlyBoundSampleCode.PublishSingleProduct();
            earlyBoundSampleCode.RetrieveProductInfo_QueryExpression();
            earlyBoundSampleCode.RetrieveProductInfo_FetchXML();
            earlyBoundSampleCode.UpdatePriceLevelState();
            earlyBoundSampleCode.DeleteProductInfo();


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
