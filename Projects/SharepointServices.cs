using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace DNLConsole365.Projects
{
    public class SharepointServices
    {
        public static List<Entity> GetDocuments(IOrganizationService service)
        {
            var query = new QueryExpression("sharepointdocument");
            query.ColumnSet = new ColumnSet(true);
            
            return service.RetrieveMultiple(query).Entities.ToList();
        }

        public static List<Entity> GetDocumentsLocations(IOrganizationService service)
        {
            var query = new QueryExpression("sharepointdocumentlocation");
            query.ColumnSet = new ColumnSet("absoluteurl", "locationtype", "relativeurl", "servicetype", "name", "regardingobjectid", "parentsiteorlocation");

            return service.RetrieveMultiple(query).Entities.ToList();
        }

        public static List<Entity> GetSites(IOrganizationService service)
        {
            var query = new QueryExpression("sharepointsite");
            query.ColumnSet = new ColumnSet(true);

            return service.RetrieveMultiple(query).Entities.ToList();
        }

        public static IEnumerable<Entity> GetEntitesByField(IOrganizationService service, string entityName, string attributeKey, string attributeValue, params string[] attributes)
        {
            QueryExpression queryExpression = new QueryExpression(entityName)
            {
                Criteria =
                {
                    Conditions =
                    {
                        new ConditionExpression(attributeKey, ConditionOperator.Equal, attributeValue)
                    }
                }
            };
            if (attributes != null)
                queryExpression.ColumnSet = new ColumnSet(attributes);
            return service.RetrieveMultiple(queryExpression).Entities;
        }
    }
}
