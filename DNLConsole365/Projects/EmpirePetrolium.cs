using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DNLConsole365.Projects
{
    public class EmpirePetrolium
    {
        public class SharepointFolder
        {
            public string Name { get; set; }
            public string RelUrl { get; set; }
            public int ItemsCount { get; set; }
        }

        public static void RemoveComponentFromSolution(IOrganizationService service, Guid componentId, int componentType, string solutionName)
        {
            var removeRequest = new RemoveSolutionComponentRequest()
            {
                // this is the Guid you have found within your Dynamics 365 trace files
                ComponentId = componentId,
                ComponentType = componentType,
                // This is the unique name, not the display name of the solution you are trying to export
                SolutionUniqueName = solutionName
            };

            var response = service.Execute(removeRequest);
        }

        public static List<Entity> GetCrmEntities(IOrganizationService service, string logicalName, string[] cols = null)
        {
            var query = new QueryExpression(logicalName);

            query.ColumnSet = cols != null ? new ColumnSet(cols) : new ColumnSet(true);

            //query.ColumnSet = new ColumnSet("ownerid", "regardingobjectid");
            query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
            

            return service.RetrieveMultiple(query).Entities.ToList();
        }

        public static void DisableLocation(IOrganizationService service, Guid id)
        {
            service.Update(new Entity("sharepointdocumentlocation", id)
            {
                Attributes =
                {
                    new KeyValuePair<string, object>("statecode", new OptionSetValue(1)),
                    new KeyValuePair<string, object>("statuscode", new OptionSetValue(2)),
                }
            });
        }

        public static void UpdateNameAndPath(IOrganizationService service, Guid id, string pcnumber)
        {
            service.Update(new Entity("sharepointdocumentlocation", id)
            {
                Attributes =
                {
                    new KeyValuePair<string, object>("name", pcnumber),
                    new KeyValuePair<string, object>("relativeurl", pcnumber),
                }
            });
        }

        public static void CreateDocumentLocation(IOrganizationService service, Guid opportunityid, string pcnumber)
        {
            service.Create(new Entity("sharepointdocumentlocation")
            {
                Attributes =
                {
                    new KeyValuePair<string, object>("name", pcnumber),
                    new KeyValuePair<string, object>("relativeurl", pcnumber),
                    new KeyValuePair<string, object>("regardingobjectid", new EntityReference("opportunity", opportunityid)),
                    new KeyValuePair<string, object>("parentsiteorlocation", new EntityReference("sharepointdocumentlocation", new Guid("686F69E6-527F-E511-80E3-3863BB349E38"))),

                }
            });
        }

        public static void DisableSomeLocations(IOrganizationService service, List<Entity> locations)
        {
            var request = new ExecuteMultipleRequest()
            {
                Requests = new OrganizationRequestCollection(),
                Settings = new ExecuteMultipleSettings
                {
                    ContinueOnError = false,
                    ReturnResponses = true
                }
            };

            foreach(var l in locations)
            {
                l["statecode"] = new OptionSetValue(1);
                l["statuscode"] = new OptionSetValue(2);
                                
                request.Requests.Add(new UpdateRequest()
                {
                    Target = l
                });
            }

            Console.WriteLine("Disable start");
            var response = (ExecuteMultipleResponse)service.Execute(request);
            Console.WriteLine("Disable end");
        }

        public static List<Entity> GetAndDisableDocumentsLocations(IOrganizationService service)
        {
            var query = new QueryExpression("sharepointdocumentlocation");

            query.ColumnSet = new ColumnSet("absoluteurl", "locationtype", "relativeurl", "servicetype", "name", "regardingobjectid", "parentsiteorlocation");
            query.Criteria.AddCondition("regardingobjectid", ConditionOperator.NotNull);
            query.Criteria.AddCondition("name", ConditionOperator.In, new string[] { "Credit",
                            "RFE",
                            "Dealer",
                            "Legal",
                            "Property",
                            "Environmental",
                            "Tanks"});
            query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
            query.Criteria.AddCondition("locationtype", ConditionOperator.Equal, 0);
            query.Criteria.AddCondition("servicetype", ConditionOperator.Equal, 0);

            // Set initial page number
            int pageNumber = 1;
            // Collections for entities
            var entityCollection = new EntityCollection();
            var tempResult = new EntityCollection();

            // Select records using paging
            do
            {
                Console.WriteLine("Page - " + pageNumber);

                query.PageInfo.Count = 1000;
                query.PageInfo.PageNumber = pageNumber++;

                tempResult = service.RetrieveMultiple(query);

                DisableSomeLocations(service, tempResult.Entities.ToList());

                entityCollection.Entities.AddRange(tempResult.Entities);
            }
            while (tempResult.MoreRecords);

            return entityCollection.Entities.ToList(); ;
        }

        public static List<Entity> GetDocumentsLocations(IOrganizationService service)
        {
            var query = new QueryExpression("sharepointdocumentlocation");

            query.ColumnSet = new ColumnSet("absoluteurl", "locationtype", "relativeurl", "servicetype", "name", "regardingobjectid", "parentsiteorlocation");
            query.Criteria.AddCondition("regardingobjectid", ConditionOperator.NotNull);
            query.Criteria.AddCondition("name", ConditionOperator.NotIn, new string[] { "Credit",
                            "RFE",
                            "Dealer",
                            "Legal",
                            "Property",
                            "Environmental",
                            "Tanks"});
            query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
            query.Criteria.AddCondition("locationtype", ConditionOperator.Equal, 0);
            query.Criteria.AddCondition("servicetype", ConditionOperator.Equal, 0);

            // Set initial page number
            int pageNumber = 1;
            // Collections for entities
            var entityCollection = new EntityCollection();
            var tempResult = new EntityCollection();

            // Select records using paging
            do
            {
                Console.WriteLine("Page - " + pageNumber);

                query.PageInfo.Count = 5000;
                query.PageInfo.PageNumber = pageNumber++;

                tempResult = service.RetrieveMultiple(query);

                entityCollection.Entities.AddRange(tempResult.Entities);
            }
            while (tempResult.MoreRecords);

            return entityCollection.Entities.ToList(); ;
        }

    }
}
