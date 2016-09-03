using Learn.CRM.API.EntityLib;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Learn.CRM.API.LINQ
{
    public class LINQSampleCode
    {
        private OrganizationServiceProxy _orgService = null;

        public LINQSampleCode(string connString)
        {
            // Connect to the CRM web service using a connection string.
            CrmServiceClient conn = new CrmServiceClient(connString);
            _orgService = conn.OrganizationServiceProxy;
            _orgService.EnableProxyTypes();
            _context = new CRMEntities(_orgService);
        }

        public event Action<string> onLog;

        private CRMEntities _context;

        internal void EarlyBound_SingleMultiple()
        {
            var account = (from a in _context.AccountSet
                           where a.Id == new Guid("2123455d-be62-e411-80d6-b4b52f567ec8")
                           select a).First();

            onLog(string.Format("{0}\t{1}\t{2}", account.Id, account.Name, account.EMailAddress1));
        }

        internal void EarlyBound_RetrieveMultiple()
        {
            var accounts = from a in _context.AccountSet
                           select a;

            foreach (Account acc in accounts)
            {
                onLog(string.Format("{0}\t{1}\t{2}", acc.Id, acc.Name, acc.EMailAddress1));
            }
        }

        internal void LateBound_SingleMultiple()
        {
            var entity = (from a in _context.CreateQuery("account")
                          where a.GetAttributeValue<Guid>("accountid") == new Guid("2123455d-be62-e411-80d6-b4b52f567ec8")
                          select new
                          {
                              Id = a.Id,
                              Name = a.GetAttributeValue<string>("name"),
                              EmailAddress1 = a.GetAttributeValue<string>("emailaddress1")
                          }).First();

            onLog(string.Format("{0}\t{1}\t{2}", entity.Id, entity.Name, entity.EmailAddress1));
        }

        internal void LateBound_RetrieveMultiple()
        {
            var enties = from a in _context.CreateQuery("account")
                         select new
                         {
                             Id = a.Id,
                             Name = a.GetAttributeValue<string>("name"),
                             EmailAddress1 = a.GetAttributeValue<string>("emailaddress1")
                         };

            foreach (var entity in enties)
            {
                onLog(string.Format("{0}\t{1}\t{2}", entity.Id, entity.Name, entity.EmailAddress1));
            }
        }
    }
}
