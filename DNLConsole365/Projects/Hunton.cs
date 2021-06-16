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
    public class Hunton
    {
        public static List<Entity> GetActiveUsers(IOrganizationService service)
        {
            var query = new QueryExpression("systemuser");
            query.ColumnSet = new ColumnSet("firstname", "lastname", "fullname", "new_division");
            query.Criteria.AddCondition("isdisabled", ConditionOperator.Equal, false);

            return service.RetrieveMultiple(query).Entities.ToList();
        }

        public static List<Entity> GetActiveUsersWithSourceId(IOrganizationService service)
        {
            var query = new QueryExpression("systemuser");
            query.ColumnSet = new ColumnSet("firstname", "lastname", "fullname", "new_sourceid");
            query.Criteria.AddCondition("new_sourceid", ConditionOperator.NotNull);
            query.Criteria.AddCondition("isdisabled", ConditionOperator.Equal, false);

            return service.RetrieveMultiple(query).Entities.ToList();
        }        

        public static List<Entity> GetUserDashboards(IOrganizationService service)
        {
            var query = new QueryExpression("userform");
            query.ColumnSet = new ColumnSet(true);// "name", "ownerid");           

            return service.RetrieveMultiple(query).Entities.ToList();
        }

        public static Guid CreateUserView(IOrganizationService service, Entity view)
        {
            return service.Create(view);
        }

        public static List<Entity> GetUserViews(IOrganizationService service)
        {
            var query = new QueryExpression("userquery");
            query.ColumnSet = new ColumnSet(true);// "name", "ownerid");           
            query.Criteria.AddCondition("querytype", ConditionOperator.Equal, 0);

            return service.RetrieveMultiple(query).Entities.ToList();
        }

        public static List<Entity> GetUserViewsByOwnerId(IOrganizationService service, Guid ownerId)
        {
            var query = new QueryExpression("userquery");
            query.ColumnSet = new ColumnSet(true);// "name", "ownerid");           
            query.Criteria.AddCondition("ownerid", ConditionOperator.Equal, ownerId);
            query.Criteria.AddCondition("querytype", ConditionOperator.Equal, 0);

            return service.RetrieveMultiple(query).Entities.ToList();
        }

        public static List<Entity> GetSystemDashboards(IOrganizationService service)
        {
            var query = new QueryExpression("systemform");
            query.ColumnSet = new ColumnSet(true);// "name", "ownerid");           
            query.Criteria.AddCondition("typename", ConditionOperator.Equal, "Dashboard");

            return service.RetrieveMultiple(query).Entities.ToList();
        }    

        public static List<Entity> GetUserDashboardsByOwnerId(IOrganizationService service, Guid ownerid)
        {
            var query = new QueryExpression("userform");
            query.ColumnSet = new ColumnSet(true);// "name", "ownerid");           
            query.Criteria.AddCondition("ownerid", ConditionOperator.Equal, ownerid);

            return service.RetrieveMultiple(query).Entities.ToList();
        }

        public static RetrieveAllEntitiesResponse RetrieveEntities(IOrganizationService service)
        {
            RetrieveAllEntitiesRequest req = new RetrieveAllEntitiesRequest();
            req.EntityFilters = EntityFilters.Entity;
            req.RetrieveAsIfPublished = true;

            RetrieveAllEntitiesResponse response = (RetrieveAllEntitiesResponse)service.Execute(req);
            return response;            
        }
    }
}
