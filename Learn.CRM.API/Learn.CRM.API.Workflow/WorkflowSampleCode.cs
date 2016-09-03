using Learn.CRM.API.EntityLib;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Learn.CRM.API.Workflow
{
    public class WorkflowSampleCode
    {
        private OrganizationServiceProxy _orgService = null;


        public WorkflowSampleCode(string connString)
        {
            // Connect to the CRM web service using a connection string.
            CrmServiceClient conn = new CrmServiceClient(connString);        
            _orgService = conn.OrganizationServiceProxy;
            _orgService.EnableProxyTypes();
        }

        public event Action<string> onLog;

        private ParameterCollection ExecuteAction(string actionName, Dictionary<string, object> inputArguments)
        {
            OrganizationRequest req = new OrganizationRequest(actionName);
            foreach (string key in inputArguments.Keys)
            {
                req[key] = inputArguments[key];
            }

            //execute the request
            OrganizationResponse response = _orgService.Execute(req);
            return response.Results;
        }

        private void ExecuteAction(string actionName, string targetEntityName, Guid EntityId)
        {
            var inputs = new Dictionary<string, object>();
            var entityReferenct = new EntityReference(targetEntityName, EntityId);
            inputs.Add("target", entityReferenct);
            this.ExecuteAction(actionName, inputs);
        }

        public void ExecuteSampleWorkflow()
        {
            //取得工作流程Id
            var workflowName = "VisitedAccount";
            QueryExpression objQueryExpression = new QueryExpression(Learn.CRM.API.EntityLib.Workflow.EntityLogicalName);
            objQueryExpression.ColumnSet = new ColumnSet(true);
            objQueryExpression.Criteria.AddCondition(new ConditionExpression("name", ConditionOperator.Equal, workflowName));
            objQueryExpression.Criteria.AddCondition(new ConditionExpression("parentworkflowid", ConditionOperator.Null));
            EntityCollection entColWorkflows = _orgService.RetrieveMultiple(objQueryExpression);

            //取得工作流程對應的客戶實體資料
            var accounts = _orgService.RetrieveMultiple(new QueryExpression()
            {
                EntityName = Account.EntityLogicalName,
                ColumnSet = new ColumnSet("accountid", "name"),
                PageInfo = new PagingInfo
                {
                    Count = 1,
                     PageNumber = 1,
                }
            });

            if (entColWorkflows != null && entColWorkflows.Entities.Count > 0)
            {
                var wfId = entColWorkflows.Entities[0].Id;
                ExecuteWorkflowRequest request = new ExecuteWorkflowRequest()
                {
                    WorkflowId = wfId,
                    EntityId = accounts.Entities.First().Id,
                };

                // 執行工作流程
                ExecuteWorkflowResponse response = _orgService.Execute(request) as ExecuteWorkflowResponse;
            }
            else
            {
                onLog(string.Format("找不到對應的工作流程。WorkflowName:{0}", workflowName));
            }
        }

        public void ExecuteSampleAction()
        {
            //執行Action

            var actionName = "new_CreateAccount";
            var inputs = new Dictionary<string, object>();
            inputs.Add("name", "JeremyLin");
            var results = ExecuteAction(actionName, inputs);

            onLog("Response Results:");
            foreach (var result in results)
            {
                onLog(string.Format("{0}:{1}", result.Key, result.Value));
            }
        }
    }
}
