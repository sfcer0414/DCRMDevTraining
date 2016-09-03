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

namespace Learn.CRM.API.EarlyBound
{
    public class EarlyBoundSampleCode
    {
        private OrganizationServiceProxy _orgService = null;

        public EarlyBoundSampleCode(string connString)
        {
            // Connect to the CRM web service using a connection string.
            CrmServiceClient conn = new CrmServiceClient(connString);
            _orgService = conn.OrganizationServiceProxy;
            _orgService.EnableProxyTypes();

        }

        public event Action<string> onLog;

        internal void SingleMultiple()
        {
            var account = (Account)_orgService.Retrieve(Account.EntityLogicalName
                , new Guid("2123455d-be62-e411-80d6-b4b52f567ec8")
                , new ColumnSet("accountid", "name", "emailaddress1"));

            onLog(string.Format("{0}\t{1}\t{2}", account.Id, account.Name, account.EMailAddress1));
        }

        internal void RetrieveMultiple()
        {
            var qe = new QueryExpression(Account.EntityLogicalName);
            qe.ColumnSet = new ColumnSet("accountid", "name", "emailaddress1");

            var accounts = _orgService.RetrieveMultiple(qe);

            foreach (Account acc in accounts.Entities)
            {
                onLog(string.Format("{0}\t{1}\t{2}", acc.Id, acc.Name, acc.EMailAddress1));
            }
        }

        internal void AddProduct()
        {
            //建立單位群組
            var newUnitGroup = new UoMSchedule();
            newUnitGroup.Name = "觸控面板單位";
            newUnitGroup.BaseUoMName = "觸控面板基本單位";
            var unitGroupId = _orgService.Create(newUnitGroup);

            onLog(string.Format("Create Unit Group => {0}", unitGroupId));

            //取得單位群組的預設單位
            QueryExpression unitQuery = new QueryExpression
            {
                EntityName = UoM.EntityLogicalName,
                ColumnSet = new ColumnSet("uomid", "name"),
                Criteria = new FilterExpression(),
                PageInfo = new PagingInfo
                {
                    PageNumber = 1,
                    Count = 1
                }
            };
            unitQuery.Criteria.AddCondition("uomscheduleid", ConditionOperator.Equal, unitGroupId);

            var unit = (UoM)_orgService.RetrieveMultiple(unitQuery).Entities[0];

            onLog(string.Format("Get Default Unit Of Unit Group => {0}", unit.Id));

            /*
           Product.ProductStructure:
               1 表示建立產品
               2 表示建立產品系列
               3 表示建立搭售方案
           */

            //建立產品系列 -> 觸控面板烤爐
            var productFamily = new Product();
            productFamily.Name = "觸控面板烤爐";
            productFamily.ProductNumber = "HPCP";
            productFamily.ProductStructure = new OptionSetValue(2);

            var productCategoryId = _orgService.Create(productFamily);

            onLog(string.Format("Create Product Family => {0}", productCategoryId));                             

            //建立產品 -> Dehydration-Bake
            var prod1 = new Product();
            prod1.Name = "Dehydration-Bake";
            prod1.ProductNumber = "HPCP-DB";
            prod1.ProductStructure = new OptionSetValue(1);
            prod1.ParentProductId = new EntityReference(Product.EntityLogicalName, productCategoryId);
            prod1.QuantityDecimal = 2;
            prod1.DefaultUoMScheduleId = new EntityReference(UoMSchedule.EntityLogicalName, unitGroupId);
            prod1.DefaultUoMId = new EntityReference(UoM.EntityLogicalName, unit.Id);

            var prod1Id = _orgService.Create(prod1);
            onLog(string.Format("Create Product 1=> {0}", prod1Id));

            //建立產品 -> Pre-Bake
            var prod2 = new Product();
            prod2.Name = "Pre-Bake";
            prod2.ProductNumber = "HPCP-PreB";
            prod2.ProductStructure = new OptionSetValue(1);
            prod2.ParentProductId = new EntityReference(Product.EntityLogicalName, productCategoryId);
            prod2.QuantityDecimal = 2;
            prod2.DefaultUoMScheduleId = new EntityReference(UoMSchedule.EntityLogicalName, unitGroupId);
            prod2.DefaultUoMId = new EntityReference(UoM.EntityLogicalName, unit.Id);

            var prod2Id = _orgService.Create(prod2);
            onLog(string.Format("Create Product 2=> {0}", prod2Id));

            //建立產品 -> Post-Bake
            var prod3 = new Product();
            prod3.Name = "Post-Bake";
            prod3.ProductNumber = "HPCP-PostB";
            prod3.ProductStructure = new OptionSetValue(1);
            prod3.ParentProductId = new EntityReference(Product.EntityLogicalName, productCategoryId);
            prod3.QuantityDecimal = 2;
            prod3.DefaultUoMScheduleId = new EntityReference(UoMSchedule.EntityLogicalName, unitGroupId);
            prod3.DefaultUoMId = new EntityReference(UoM.EntityLogicalName, unit.Id);

            var prod3Id = _orgService.Create(prod3);

            onLog(string.Format("Create Product 3=> {0}", prod3Id));

            //建立價目表
            PriceLevel newPriceList = new PriceLevel
            {
                Name = "觸控面板單位官網定價"
            };
            var priceLevelId = _orgService.Create(newPriceList);

            //建立價目表項目 (根據剛剛建立產品的Id來設定)

            //建立第一個產品價目表項目
            ProductPriceLevel newPriceListItem1 = new ProductPriceLevel
            {
                PriceLevelId = new EntityReference(PriceLevel.EntityLogicalName, priceLevelId),
                ProductId = new EntityReference(Product.EntityLogicalName, prod1Id),
                UoMId = new EntityReference(UoM.EntityLogicalName, unit.Id),
                Amount = new Money(20)
            };
            var priceListItem1Id = _orgService.Create(newPriceListItem1);

            Console.WriteLine("Created price list for {0}", prod1.Name);

            //建立第二個產品價目表項目
            ProductPriceLevel newPriceListItem2 = new ProductPriceLevel
            {
                PriceLevelId = new EntityReference(PriceLevel.EntityLogicalName, priceLevelId),
                ProductId = new EntityReference(Product.EntityLogicalName, prod2Id),
                UoMId = new EntityReference(UoM.EntityLogicalName, unit.Id),
                Amount = new Money(1000)
            };
            var priceListItem2Id = _orgService.Create(newPriceListItem2);

            Console.WriteLine("Created price list for {0}", prod2.Name);

            //建立第三個產品價目表項目
            ProductPriceLevel newPriceListItem3 = new ProductPriceLevel
            {
                PriceLevelId = new EntityReference(PriceLevel.EntityLogicalName, priceLevelId),
                ProductId = new EntityReference(Product.EntityLogicalName, prod3Id),
                UoMId = new EntityReference(UoM.EntityLogicalName, unit.Id),
                Amount = new Money(99)
            };
            var priceListItem3Id = _orgService.Create(newPriceListItem3);

            Console.WriteLine("Created price list for {0}", prod3.Name);
        }

        internal void PublishSingleProduct()
        {
            //取得單位群組
            QueryExpression uomsQuery = new QueryExpression
            {
                EntityName = UoMSchedule.EntityLogicalName,
                ColumnSet = new ColumnSet("uomscheduleid", "name"),
            };
            uomsQuery.Criteria.AddCondition("name", ConditionOperator.Equal, "觸控面板單位");

            var unitGroup = (UoMSchedule)_orgService.RetrieveMultiple(uomsQuery).Entities[0];

            //取得預設單位
            QueryExpression unitQuery = new QueryExpression
            {
                EntityName = UoM.EntityLogicalName,
                ColumnSet = new ColumnSet("uomid", "name"),
                Criteria = new FilterExpression(),
                PageInfo = new PagingInfo
                {
                    PageNumber = 1,
                    Count = 1
                }
            };
            unitQuery.Criteria.AddCondition("uomscheduleid", ConditionOperator.Equal, unitGroup.Id);
            var unit = (UoM)_orgService.RetrieveMultiple(unitQuery).Entities[0];

            //建立沒有產品系列的產品項目
            var prod1 = new Product();
            prod1.Name = "Single Product";
            prod1.ProductNumber = "SP-001";
            prod1.ProductStructure = new OptionSetValue(1);
            prod1.QuantityDecimal = 2;
            prod1.DefaultUoMScheduleId = new EntityReference(UoMSchedule.EntityLogicalName, unitGroup.Id);
            prod1.DefaultUoMId = new EntityReference(UoM.EntityLogicalName, unit.Id);

            var prod1Id = _orgService.Create(prod1);
            onLog(string.Format("Create Single Product => {0}", prod1Id));

            //透過更新狀態，來執行產品發布動作。
            prod1.Id = prod1Id;
            prod1.StateCode = ProductState.Active;
            _orgService.Update(prod1);
            onLog(string.Format("Publish Single Product Success"));
        }

        internal void Associate_Disassociate()
        {
            //建立等等要關聯用的客戶資料
            var account1 = new Account() { Name = "TestAccount" };
            var acc1Id = _orgService.Create(account1);

            var account2 = new Account() { Name = "TestAccount2" };
            var acc2Id = _orgService.Create(account2);
            //建立等等要關聯用的聯絡人資料
            var contact1 = new Contact() { FirstName = "Lin", LastName = "Jeremy" };
            var contact1Id = _orgService.Create(contact1);

            //建立客戶主要聯絡人的實體關聯物件
            Relationship rel = new Relationship("account_primary_contact");

            //把要被設定主要聯絡人的客戶加入到集合中
            var accounts = new EntityReferenceCollection();
            accounts.Add(new EntityReference(Account.EntityLogicalName, acc1Id));
            accounts.Add(new EntityReference(Account.EntityLogicalName, acc2Id));            

            //從連絡人執行關聯到客戶
            _orgService.Associate(Contact.EntityLogicalName, contact1Id, rel, accounts);

            onLog(string.Format("Associate Product Success"));

            //從聯絡人解除跟客戶的關聯
            _orgService.Disassociate(Contact.EntityLogicalName, contact1Id, rel, accounts);

            onLog(string.Format("Disassociate Product Success"));

            //把要設定主要聯絡人的聯絡人加入到集合中
            var contacts = new EntityReferenceCollection();
            contacts.Add(new EntityReference(Contact.EntityLogicalName, contact1Id));

            //從客戶執行關聯到聯絡人
            _orgService.Associate(Account.EntityLogicalName, acc1Id, rel, contacts);
            _orgService.Associate(Account.EntityLogicalName, acc2Id, rel, contacts);

            onLog(string.Format("Other Way Associate Product Success"));
        }

        internal void UpdateProduct()
        {
            QueryExpression productsQuery = new QueryExpression
            {
                EntityName = "product",
                ColumnSet = new ColumnSet("productid", "name"),
            };

            productsQuery.Criteria.AddCondition("name", ConditionOperator.Equal, "觸控面板烤爐");

            var product = (Product)_orgService.RetrieveMultiple(productsQuery).Entities[0];
            product.ProductNumber = "HPCP-00";
            _orgService.Update(product);

            onLog(string.Format("Update Product Success"));
        }

        internal void PublishProduct()
        {
            //取得產品系列 Id
            QueryExpression productsQuery = new QueryExpression
            {
                EntityName = Product.EntityLogicalName,
                ColumnSet = new ColumnSet("productid", "name"),
            };
            productsQuery.Criteria.AddCondition("name", ConditionOperator.Equal, "觸控面板烤爐");

            var product = (Product)_orgService.RetrieveMultiple(productsQuery).Entities[0];
            onLog(string.Format("Get Publish Product Family Id => {0}", product.Id));

            //透過產品系列，執行階層發布
            PublishProductHierarchyRequest publishReq = new PublishProductHierarchyRequest
            {
                Target = new EntityReference(Product.EntityLogicalName, product.Id)
            };
            PublishProductHierarchyResponse published = (PublishProductHierarchyResponse)_orgService.Execute(publishReq);
            if (published.Results != null)
            {
                onLog("Published the product records");
            }
        }

        internal void RetrieveProductInfo_QueryExpression()
        {
            //建立JOIN查詢
            QueryExpression productsQuery = new QueryExpression
            {
                EntityName = Product.EntityLogicalName,
                ColumnSet = new ColumnSet("productid", "name", "productnumber"),
            };

            var linkEntity = productsQuery.AddLink(Product.EntityLogicalName, "parentproductid", "productid", JoinOperator.Natural);
            linkEntity.Columns = new ColumnSet("productid", "name", "productnumber");
            linkEntity.EntityAlias = "parentproduct";
            linkEntity.LinkCriteria.AddCondition("name", ConditionOperator.Equal, "觸控面板烤爐");

            var products = _orgService.RetrieveMultiple(productsQuery).Entities;

            //取得查詢實體的欄位資料
            var categoryPN = (AliasedValue)products[0].Attributes["parentproduct.productnumber"];
            var categoryName = (AliasedValue)products[0].Attributes["parentproduct.name"];
            onLog(string.Format("Product Family => ({0}) {1}", categoryPN.Value, categoryName.Value));

            foreach (Product prod in products)
            {
                onLog(string.Format("Product => ({0}) {1}", prod.ProductNumber, prod.Name));
            }
        }

        internal void RetrieveProductInfo_FetchXML()
        {
            //建立JOIN查詢
            FetchExpression productsFetchQuery = new FetchExpression(@"
            <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                <entity name='product'>
                <attribute name='name'/>
                <attribute name='productnumber'/>
                <attribute name='statecode'/>
                <attribute name='productstructure'/>
                <order attribute='productnumber' descending='false'/>
                <link-entity name='product' from='productid' to='parentproductid' alias='ae'>
                    <attribute name='name'/>
                    <attribute name='productnumber'/>
                    <filter type='and'>
                    <condition attribute='name' operator='eq' value='觸控面板烤爐'/>
                    </filter>
                </link-entity>
                </entity>
            </fetch>
            ");

            var products = _orgService.RetrieveMultiple(productsFetchQuery).Entities;

            //取得查詢實體的欄位資料
            var categoryPN = (AliasedValue)products[0].Attributes["ae.productnumber"];
            var categoryName = (AliasedValue)products[0].Attributes["ae.name"];
            onLog(string.Format("Product Family => ({0}) {1}", categoryPN.Value, categoryName.Value));

            foreach (Product prod in products)
            {
                onLog(string.Format("Product => ({0}) {1}", prod.ProductNumber, prod.Name));
            }
        }

        internal void UpdatePriceLevelState()
        {
            //透過更新售價表狀態，來執行啟用動作
            var query = new QueryExpression()
            {
                EntityName = PriceLevel.EntityLogicalName,
                ColumnSet = new ColumnSet("pricelevelid", "name"),
            };
            query.Criteria.AddCondition("name", ConditionOperator.Equal, "觸控面板單位官網定價");
            var priceLevels = _orgService.RetrieveMultiple(query);

            var pl = (PriceLevel)priceLevels.Entities.First();
            pl.StateCode = PriceLevelState.Active;
            _orgService.Update(pl);
        }

        internal void DeleteProductInfo()
        {
            //取要刪除的產品系列與無產品系列的產品            
            QueryExpression productsQuery = new QueryExpression
            {
                EntityName = "product",
                ColumnSet = new ColumnSet("productid", "name"),
            };

            productsQuery.Criteria.AddCondition("name", ConditionOperator.Equal, "觸控面板烤爐");
            productsQuery.Criteria.AddCondition("name", ConditionOperator.Equal, "Single Product");
            productsQuery.Criteria.FilterOperator = LogicalOperator.Or;

            var products = _orgService.RetrieveMultiple(productsQuery).Entities;

            //刪除產品
            foreach(Product product in products)
            {
                onLog(string.Format("Get Delete Product Id => {0}", product.Id));
                _orgService.Delete(Product.EntityLogicalName, product.Id);
                onLog("Delete Product of Product Family Success :) ");
            }

            //取得要刪除的單位群組
            QueryExpression uomsQuery = new QueryExpression
            {
                EntityName = UoMSchedule.EntityLogicalName,
                ColumnSet = new ColumnSet("uomscheduleid", "name"),
            };
            uomsQuery.Criteria.AddCondition("name", ConditionOperator.Equal, "觸控面板單位");

            var uoms = (UoMSchedule)_orgService.RetrieveMultiple(uomsQuery).Entities[0];

            //刪除單位群組
            onLog(string.Format("Get Delete UoMSchedule => {0}", uoms.Id));
            _orgService.Delete(UoMSchedule.EntityLogicalName, uoms.Id);
            onLog("Delete UoMSchedule Success :) ");

            //取得要刪除的售價表
            QueryExpression priceLevelQuery = new QueryExpression
            {
                EntityName = PriceLevel.EntityLogicalName,
                ColumnSet = new ColumnSet("pricelevelid", "name"),
            };
            priceLevelQuery.Criteria.AddCondition("name", ConditionOperator.Equal, "觸控面板單位官網定價");

            var pricelevel = (PriceLevel)_orgService.RetrieveMultiple(priceLevelQuery).Entities[0];

            //刪除售價表
            onLog(string.Format("Get Delete PriceLevel => {0}", pricelevel.Id));
            _orgService.Delete(PriceLevel.EntityLogicalName, pricelevel.Id);
            onLog("Delete PriceLevel Success :) ");
        }
    }
}
