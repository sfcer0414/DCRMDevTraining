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

namespace Learn.CRM.API.LateBound
{
    public class LateBoundSampleCode
    {
        private OrganizationServiceProxy _orgService = null;
        public LateBoundSampleCode(string connString)
        {
            // Connect to the CRM web service using a connection string.
            CrmServiceClient conn = new CrmServiceClient(connString);
            _orgService = conn.OrganizationServiceProxy;
        }

        public event Action<string> onLog;

        internal void SingleMultiple()
        {
            var account = _orgService.Retrieve("account"
                , new Guid("2123455d-be62-e411-80d6-b4b52f567ec8")
                , new ColumnSet("accountid", "name", "emailaddress1"));

            var id = account.Attributes["accountid"];
            var name = account.Attributes["name"];
            var emailAddress = account.Attributes.ContainsKey("emailaddress1") ? account.Attributes["emailaddress1"] : "";

            onLog(string.Format("{0}\t{1}\t{2}", id, name, emailAddress));
        }

        internal void RetrieveMultiple()
        {
            var qe = new QueryExpression("account");
            qe.ColumnSet = new ColumnSet("accountid", "name", "emailaddress1");

            var accounts = _orgService.RetrieveMultiple(qe);
            foreach (var account in accounts.Entities)
            {
                var id = account.Attributes["accountid"];
                var name = account.Attributes["name"];
                var emailAddress = account.Attributes.ContainsKey("emailaddress1") ? account.Attributes["emailaddress1"] : "";

                onLog(string.Format("{0}\t{1}\t{2}", id, name, emailAddress));

            }
        }

        internal void AddProduct()
        {
            //建立單位群組
            var newUnitGroup = new Entity("uomschedule");
            newUnitGroup.Attributes["name"] = "Example Unit Group";
            newUnitGroup.Attributes["baseuomname"] = "Example Primary Unit";
            var unitGroupId = _orgService.Create(newUnitGroup);

            onLog(string.Format("Create Unit Group => {0}", unitGroupId));

            //取得預設單位
            QueryExpression unitQuery = new QueryExpression
            {
                EntityName = "uom",
                ColumnSet = new ColumnSet("uomid", "name"),
                Criteria = new FilterExpression(),
                PageInfo = new PagingInfo
                {
                    PageNumber = 1,
                    Count = 1
                }
            };
            unitQuery.Criteria.AddCondition("uomscheduleid", ConditionOperator.Equal, unitGroupId);

            var unit = _orgService.RetrieveMultiple(unitQuery).Entities[0];
            var unitId = (Guid)unit.Attributes["uomid"];

            onLog(string.Format("Get Default Unit Of Unit Group => {0}", unitId));

            /*
            Product.ProductStructure:
                1 表示建立產品
                2 表示建立產品系列
                3 表示建立搭售方案
            */


            //Create Product Family -> 電子連接器高溫工程塑膠
            var productFamily = new Entity("product");
            productFamily.Attributes["name"] = "電子連接器高溫工程塑膠";
            productFamily.Attributes["productnumber"] = "CPA";
            productFamily.Attributes["productstructure"] = new OptionSetValue(2);

            var productCategoryId = _orgService.Create(productFamily);

            onLog(string.Format("Create Product Family => {0}", productCategoryId));
            //Create Product -> PA9T, LCP, PBT                      

            //PA9T
            var prod1 = new Entity("product");
            prod1.Attributes["name"] = "PA9T";
            prod1.Attributes["productnumber"] = "CPA-PA9T";
            prod1.Attributes["productstructure"] = new OptionSetValue(1);
            prod1.Attributes["parentproductid"] = new EntityReference("product", productCategoryId);
            prod1.Attributes["quantitydecimal"] = 2;
            prod1.Attributes["defaultuomscheduleid"] = new EntityReference("uomschedule", unitGroupId);
            prod1.Attributes["defaultuomid"] = new EntityReference("uom", unitId);

            var prod1Id = _orgService.Create(prod1);
            onLog(string.Format("Create Product 1=> {0}", prod1Id));

            //LCP
            var prod2 = new Entity("product");
            prod2.Attributes["name"] = "LCP";
            prod2.Attributes["productnumber"] = "CPA-LCP";
            prod2.Attributes["productstructure"] = new OptionSetValue(1);
            prod2.Attributes["parentproductid"] = new EntityReference("product", productCategoryId);
            prod2.Attributes["quantitydecimal"] = 2;
            prod2.Attributes["defaultuomscheduleid"] = new EntityReference("uomschedule", unitGroupId);
            prod2.Attributes["defaultuomid"] = new EntityReference("uom", unitId);

            var prod2Id = _orgService.Create(prod2);
            onLog(string.Format("Create Product 2=> {0}", prod2Id));

            //PBT
            var prod3 = new Entity("product");
            prod3.Attributes["name"] = "PBT";
            prod3.Attributes["productnumber"] = "CPA-PBT";
            prod3.Attributes["productstructure"] = new OptionSetValue(1);
            prod3.Attributes["parentproductid"] = new EntityReference("product", productCategoryId);
            prod3.Attributes["quantitydecimal"] = 2;
            prod3.Attributes["defaultuomscheduleid"] = new EntityReference("uomschedule", unitGroupId);
            prod3.Attributes["defaultuomid"] = new EntityReference("uom", unitId);

            var prod3Id = _orgService.Create(prod3);
            onLog(string.Format("Create Product => {0}", prod3Id));
        }

        internal void UpdateProduct()
        {
            QueryExpression productsQuery = new QueryExpression
            {
                EntityName = "product",
                ColumnSet = new ColumnSet("productid", "name"),
            };

            productsQuery.Criteria.AddCondition("name", ConditionOperator.Equal, "電子連接器高溫工程塑膠");

            var product = _orgService.RetrieveMultiple(productsQuery).Entities[0];
            product.Attributes["productnumber"] = "CPA-00";
            _orgService.Update(product);

            onLog(string.Format("Update Product Success"));
        }

        internal void PublishProduct()
        {
            //Get Product Family Id
            QueryExpression productsQuery = new QueryExpression
            {
                EntityName = "product",
                ColumnSet = new ColumnSet("productid", "name"),
            };

            productsQuery.Criteria.AddCondition("name", ConditionOperator.Equal, "電子連接器高溫工程塑膠");

            var product = _orgService.RetrieveMultiple(productsQuery).Entities[0];
            var productId = (Guid)product.Attributes["productid"];
            onLog(string.Format("Get Publish Product Family Id => {0}", productId));

            //Publish Product Family
            PublishProductHierarchyRequest publishReq = new PublishProductHierarchyRequest
            {
                Target = new EntityReference("product", productId)
            };
            PublishProductHierarchyResponse published = (PublishProductHierarchyResponse)_orgService.Execute(publishReq);
            if (published.Results != null)
            {
                onLog("Published the product records");
            }
        }

        internal void RetrieveProductInfo()
        {
            QueryExpression productsQuery = new QueryExpression
            {
                EntityName = "product",
                ColumnSet = new ColumnSet("productid", "name", "productnumber"),
            };

            var linkEntity = productsQuery.AddLink("product", "parentproductid", "productid");
            linkEntity.Columns = new ColumnSet("productid", "name", "productnumber");
            linkEntity.EntityAlias = "parentproduct";
            linkEntity.LinkCriteria.AddCondition("name", ConditionOperator.Equal, "電子連接器高溫工程塑膠");

            var products = _orgService.RetrieveMultiple(productsQuery).Entities;

            var categoryPN = (AliasedValue)products[0].Attributes["parentproduct.productnumber"];
            var categoryName = (AliasedValue)products[0].Attributes["parentproduct.name"];
            onLog(string.Format("Product Family => ({0}) {1}", categoryPN.Value, categoryName.Value));

            foreach (var prod in products)
            {
                onLog(string.Format("Product => ({0}) {1}", prod.Attributes["productnumber"], prod.Attributes["name"]));
            }
        }

        internal void DeleteProductInfo()
        {

            //Get Delete Product Id
            QueryExpression productsQuery = new QueryExpression
            {
                EntityName = "product",
                ColumnSet = new ColumnSet("productid", "name"),
            };

            productsQuery.Criteria.AddCondition("name", ConditionOperator.Equal, "電子連接器高溫工程塑膠");

            var product = _orgService.RetrieveMultiple(productsQuery).Entities[0];
            var productId = (Guid)product.Attributes["productid"];

            onLog(string.Format("Get Delete Product Family Id => {0}", productId));
            _orgService.Delete("product", productId);
            onLog("Delete Product of Product Family Success :) ");

            //Get Delete UoMSchedule Id
            QueryExpression uomsQuery = new QueryExpression
            {
                EntityName = "uomschedule",
                ColumnSet = new ColumnSet("uomscheduleid", "name"),
            };
            uomsQuery.Criteria.AddCondition("name", ConditionOperator.Equal, "Example Unit Group");

            var uoms = _orgService.RetrieveMultiple(uomsQuery).Entities[0];
            var uomsId = (Guid)uoms.Attributes["uomscheduleid"];

            onLog(string.Format("Get Delete UoMSchedule => {0}", uomsId));
            _orgService.Delete("uomschedule", uomsId);
            onLog("Delete UoMSchedule Success :) ");
        }
    }
}
