using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;

namespace DNLConsole365.Projects
{
    public class LegalServices
    {
        private class EntityComparer : IEqualityComparer<Entity>
        {           

            public bool Equals(Entity x, Entity y)
            {
                return x.Id.Equals(y.Id);                
            }

            public int GetHashCode(Entity obj)
            {
                return obj.GetHashCode();
            }
        }

        public enum TransactionsStatus
        {
            Draft = 1,
            Reconciled = 267070000
        }

        /*public static EntityCollection GetAssociatedRecords(IOrganizationService service, EntityReference parentRecord, string sourceRecordName)
        {
            var query = new QueryExpression(sourceRecordName);

            query.ColumnSet = new ColumnSet(false);
            query.Criteria.AddCondition(parentRecord.LogicalName, ConditionOperator.Equal, parentRecord.Id);

            return service.RetrieveMultiple(query);
            
            /*switch(sourceRecordName)
            {
                case "fw_timeenty":
                    {
                        query.ColumnSet = new ColumnSet("")
                        break;
                    }
            }
        }*/

        public static bool CheckUserHasRole(IOrganizationService service, Guid userId, string roleName)
        {
            // Get Role by its name
            var query = new QueryExpression
            {
                EntityName = "role",
                ColumnSet = new ColumnSet("roleid"),
                Criteria = new FilterExpression
                {
                    Conditions =
                                {

                                    new ConditionExpression
                                    {
                                        AttributeName = "name",
                                        Operator = ConditionOperator.Equal,
                                        Values = { roleName }
                                    }
                                }
                }
            };

            // Get the role.
            var givenRoles = service.RetrieveMultiple(query);

            // Check if role exists
            if (givenRoles.Entities.Count > 0)
            {
                var givenRole = givenRoles.Entities.FirstOrDefault();

                // Establish a SystemUser link for a query.
                var systemUserLink = new LinkEntity()
                {
                    LinkFromEntityName = "systemuserroles",
                    LinkFromAttributeName = "systemuserid",
                    LinkToEntityName = "systemuser",
                    LinkToAttributeName = "systemuserid",
                    LinkCriteria =
                                {
                                    Conditions =
                                    {
                                        new ConditionExpression(
                                            "systemuserid", ConditionOperator.Equal, userId)
                                    }
                                }
                };

                // Build the query.
                var linkQuery = new QueryExpression()
                {
                    EntityName = "role",
                    ColumnSet = new ColumnSet("roleid"),
                    LinkEntities =
                                {
                                    new LinkEntity()
                                    {
                                        LinkFromEntityName = "role",
                                        LinkFromAttributeName = "roleid",
                                        LinkToEntityName = "systemuserroles",
                                        LinkToAttributeName = "roleid",
                                        LinkEntities = {systemUserLink}
                                    }
                                },
                                Criteria =
                                {
                                    Conditions =
                                    {
                                        new ConditionExpression("roleid", ConditionOperator.Equal, givenRole.Id)
                                    }
                                }
                };

                // Retrieve matching roles for selected user
                var matchEntities = service.RetrieveMultiple(linkQuery);

                // if an entity is returned then the user is a member
                // of the role
                return (matchEntities.Entities.Count > 0);
            }

            return false;
        }

        public static Entity GetDefaultPriceList(IOrganizationService service)
        {
            var query = new QueryExpression("pricelevel");
            query.Criteria.AddCondition("name", ConditionOperator.Equal, "Default");
            query.ColumnSet = new ColumnSet(false);

            return service.RetrieveMultiple(query).Entities.FirstOrDefault();
        }

        public static void AssociateRecords(IOrganizationService service, EntityReference associateTarget, string relationshipName, List<Entity> associatedEntities)
        {
            // Creating EntityReferenceCollection for the Contact
            var relatedEntities = new EntityReferenceCollection();

            // Add related entities to list
            foreach (var entity in associatedEntities)
            {
                relatedEntities.Add(new EntityReference(entity.LogicalName, entity.Id));
            }

            // Add the relationship using schema name
            var relationship = new Relationship(relationshipName);

            // Associate the records to Target entity
            service.Associate(associateTarget.LogicalName, associateTarget.Id, relationship, relatedEntities);
        }

        public static Guid CreateRecord(IOrganizationService service, string logicalName, Dictionary<string, object> attributes)
        {
            var entity = new Entity(logicalName);

            foreach (var attr in attributes.Keys)
            {
                entity[attr] = attributes[attr];
            }

            return service.Create(entity);
        }

        public static Entity RetrieveRecord(IOrganizationService service, string logicalName, Guid id, string[] cols)
        {
            return service.Retrieve(logicalName, id, new ColumnSet(cols));
        }

        public static void UpdateEntityStatus(IOrganizationService service, Entity entity, int stateCode, int statusCode)
        {
            var attributes = new Dictionary<string, object>();

            attributes.Add("statecode", new OptionSetValue(stateCode));
            attributes.Add("statuscode", new OptionSetValue(statusCode));

            UpdateRecord(service, entity, attributes);
        }

        public static void UpdateRecord(IOrganizationService service, Entity entity, Dictionary<string, object> attributes)
        {
            foreach (var attr in attributes.Keys)
            {
                entity[attr] = attributes[attr];
            }

            service.Update(entity);
        }

        public static List<Entity> GetAssociatedRecords(IOrganizationService service, string sourceRecordName, Guid id, bool allExpenses = true)
        {
            var relationshipEntityName = "fw_invoice_" + sourceRecordName;
            var query = new QueryExpression(sourceRecordName);

            query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
            query.ColumnSet = new ColumnSet(false);

            // Look for time sheets or expenses
            var linkEntity = new LinkEntity(sourceRecordName,
                                                relationshipEntityName,
                                                sourceRecordName + "id",
                                                sourceRecordName + "id",
                                                JoinOperator.Inner);
            // Link invoices
            var linkInvoice = new LinkEntity(relationshipEntityName,
                                                "invoice",
                                                "invoiceid",
                                                "invoiceid",
                                                JoinOperator.Inner);
            // Add invoice alias to cut paid expenses and time sheets
            linkInvoice.EntityAlias = "invoice";
            linkInvoice.Columns = new ColumnSet("invoiceid", "statecode");
            // Select only related to invoice expenses

//            linkInvoice.LinkCriteria = new FilterExpression();
  //          linkInvoice.LinkCriteria.AddCondition(new ConditionExpression("invoiceid",
    //                                                            allExpenses ? ConditionOperator.NotEqual : ConditionOperator.Equal, 
      //                                                          id));
            linkEntity.LinkEntities.Add(linkInvoice);
            query.LinkEntities.Add(linkEntity);
            
            // Return only active
            return service.RetrieveMultiple(query).Entities.Where(e => e.Contains("invoice.invoiceid") &&
                                                                     e.Contains("invoice.statecode") &&
                                                                    ((OptionSetValue)((AliasedValue)e["invoice.statecode"]).Value).Value == 0).ToList();            
        }

        public static List<Entity> GetRecords(IOrganizationService service, string entityName, string[] columns)
        {
            var query = new QueryExpression(entityName);
            query.ColumnSet = new ColumnSet(columns);

            return service.RetrieveMultiple(query).Entities.ToList();
        }

        public static List<Entity> GetAssociatedRecords(IOrganizationService service, EntityReference parentRecord, string sourceRecordName)
        {   
            // Setup relationship name         
            var relationshipEntityName = "fw_invoice_" + sourceRecordName;

            var query = new QueryExpression(sourceRecordName);

            query.ColumnSet = new ColumnSet(false);
            // Look for records connected to parent record
            query.Criteria.AddCondition(parentRecord.LogicalName, ConditionOperator.Equal, parentRecord.Id);

            // Look for time sheets or expenses
            var linkEntity = new LinkEntity(sourceRecordName, 
                                                relationshipEntityName,
                                                sourceRecordName +"id",
                                                sourceRecordName+ "id",
                                                JoinOperator.LeftOuter);
            // Link invoices
            var linkInvoice = new LinkEntity(relationshipEntityName, 
                                                "invoice", 
                                                "invoiceid", 
                                                "invoiceid", 
                                                JoinOperator.LeftOuter);
            // Add invoice alias to cut paid expenses and time sheets
            linkInvoice.EntityAlias = "invoice";
            linkInvoice.Columns = new ColumnSet("invoiceid", "statecode");

            linkEntity.LinkEntities.Add(linkInvoice);
            query.LinkEntities.Add(linkEntity);

            // Check if invoice is in "Paid" status
            //linkInvoice.LinkCriteria = new FilterExpression();
            //linkInvoice.LinkCriteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.NotEqual, 2));
            //linkInvoice.LinkCriteria.FilterOperator = LogicalOperator.Or;
            //linkInvoice.LinkCriteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.NotEqual, 2));

            var returnExpensesList = new List<Entity>();
            // Get all epxenses for Matter
            var expensesCollection = service.RetrieveMultiple(query);

            if (expensesCollection.Entities.Count > 0)
            {
                // Select distinct expense
                var distinctExpensesList = expensesCollection.Entities.GroupBy(x => new { x.Id }).Select(g => g.First()).ToList();

                  // Add to return list only expenses which do not have Paid Status
                  foreach (var dex in distinctExpensesList)
                  {
                      if (expensesCollection.Entities.Any(e => e.Id == dex.Id && e.Contains("invoice.statecode") && ((OptionSetValue)((AliasedValue)e["invoice.statecode"]).Value).Value == 2))
                      {
                          continue;
                      }
                      else
                      {
                          returnExpensesList.Add(dex);
                      }
                  }                
             }            
             return returnExpensesList;
         }
    }
}
