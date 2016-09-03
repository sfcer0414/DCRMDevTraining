using Learn.CRM.API.EntityLib;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Learn.CRM.API.Over5000
{
    public class Over5000SampleCode
    {
        private OrganizationServiceProxy _orgService = null;


        public Over5000SampleCode(string connString)
        {
            // Connect to the CRM web service using a connection string.
            CrmServiceClient conn = new CrmServiceClient(connString);
            _orgService = conn.OrganizationServiceProxy;
            _orgService.EnableProxyTypes();
        }

        public event Action<string> onLog;

        public void Init1wAccountData()
        {
            var success = 0;
            var fail = 0;

            for (int batchIndex = 0; batchIndex < 10; batchIndex++)
            {
                var requestWithResults = new ExecuteMultipleRequest()
                {
                    // Assign settings that define execution behavior: continue on error, return responses. 
                    Settings = new ExecuteMultipleSettings()
                    {
                        ContinueOnError = true,
                        ReturnResponses = true
                    },
                    // Create an empty organization request collection.
                    Requests = new OrganizationRequestCollection()
                };

                for (int reqIndex = 1; reqIndex <= 1000; reqIndex++)
                {
                    var currentIndex = (batchIndex * 1000) + reqIndex;
                    var newAcc = new Account();
                    newAcc.Name = "Over5000SampleData " + currentIndex.ToString().PadLeft(5, '0');
                    requestWithResults.Requests.Add(
                        new CreateRequest()
                        {
                            Target = newAcc
                        });
                }

                var result = (ExecuteMultipleResponse)_orgService.Execute(requestWithResults);

                foreach (var resp in result.Responses)
                {
                    if (resp.Response != null)
                    {
                        success++;
                    }
                    else if (resp.Fault != null)
                    {
                        fail++;
                    }
                }

                onLog(string.Format("Create Account Batch Done!! BatchIndex=>{0}", batchIndex));
            }

            onLog(string.Format("Create Account Done!! Success=>{0}, Fail=>{1}", success, fail));
        }

        public void Remove1wAccountData()
        {
            var removeAccountQuery = new QueryExpression()
            {
                EntityName = Account.EntityLogicalName,
                ColumnSet = new ColumnSet("accountid", "name")
            };

            removeAccountQuery.Criteria.AddCondition("name", ConditionOperator.BeginsWith, "Over5000SampleData");
            removeAccountQuery.PageInfo.PageNumber = 1;
            removeAccountQuery.PageInfo.PagingCookie = null;
            EntityCollection accounts;

            var requestWithResults = new ExecuteMultipleRequest()
            {
                // Assign settings that define execution behavior: continue on error, return responses. 
                Settings = new ExecuteMultipleSettings()
                {
                    ContinueOnError = true,
                    ReturnResponses = true
                },
                // Create an empty organization request collection.
                Requests = new OrganizationRequestCollection()
            };

            do
            {
                accounts = _orgService.RetrieveMultiple(removeAccountQuery);

                var requestCount = 0;
                foreach (Account acc in accounts.Entities)
                {
                    requestWithResults.Requests.Add(
                        new DeleteRequest()
                        {
                            Target = new EntityReference(Account.EntityLogicalName, acc.Id)
                        });
                    requestCount++;

                    if (requestCount == 1000)
                    {
                        var result = (ExecuteMultipleResponse)_orgService.Execute(requestWithResults);

                        var success = 0;
                        var fail = 0;
                        foreach (var resp in result.Responses)
                        {
                            if (resp.Response != null)
                            {
                                success++;
                            }
                            else if (resp.Fault != null)
                            {
                                fail++;
                            }
                        }

                        onLog(string.Format("Delete Account Batch Done!! success:{0}, Fail:{1}", success, fail));

                        requestWithResults.Requests.Clear();
                        requestCount = 0;
                    }
                }

                removeAccountQuery.PageInfo.PageNumber++;
                removeAccountQuery.PageInfo.PagingCookie = accounts.PagingCookie;
            } while (accounts.MoreRecords);

        }

        public void Query1wAccountData()
        {
            var removeAccountQuery = new QueryExpression()
            {
                EntityName = Account.EntityLogicalName,
                ColumnSet = new ColumnSet("accountid", "name")
            };

            removeAccountQuery.Criteria.AddCondition("name", ConditionOperator.BeginsWith, "Over5000SampleData");
            removeAccountQuery.PageInfo.PageNumber = 1;
            removeAccountQuery.PageInfo.PagingCookie = null;
            EntityCollection accounts;
            do
            {
                accounts = _orgService.RetrieveMultiple(removeAccountQuery);

                foreach (Account acc in accounts.Entities)
                {
                    onLog(string.Format("Account Name:{0}", acc.Name));
                }

                removeAccountQuery.PageInfo.PageNumber++;
                removeAccountQuery.PageInfo.PagingCookie = accounts.PagingCookie;
            } while (accounts.MoreRecords);
        }
    }
}
