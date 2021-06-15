using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;

namespace DNLConsole365.Projects
{
    public class BirchEquipments
    {

        public static List<Entity> GetCrmEntiesUsingFetch(IOrganizationService service, string fetchString)
        {
            return service.RetrieveMultiple(new FetchExpression(fetchString)).Entities.ToList();
        }

        public static List<Entity> GetCrmEntities(IOrganizationService service, string logicalName, string[] cols = null)
        {
            var query = new QueryExpression(logicalName);

            query.ColumnSet = cols != null ? new ColumnSet(cols) : new ColumnSet(true);

            //query.ColumnSet = new ColumnSet("ownerid", "regardingobjectid");
            query.Criteria.AddCondition("createdon", ConditionOperator.Last7Days);

            return service.RetrieveMultiple(query).Entities.ToList();
        }


        public static List<Entity> GetPrincipalObjectAccess(IOrganizationService service)
        {
            var query = new QueryExpression("principalobjectaccess");
            query.ColumnSet = new ColumnSet("principalid", "inheritedaccessrightsmask", "accessrightsmask", "objectid");

            var linkEntity = new LinkEntity()
            {
                LinkFromEntityName = "principalobjectaccess",
                LinkToEntityName = "contact",
                LinkFromAttributeName = "objectid",
                LinkToAttributeName = "contactid"
            };

            query.Criteria.AddCondition("inheritedaccessrightsmask", ConditionOperator.Equal, 134217729);
            query.Criteria.AddCondition("accessrightsmask", ConditionOperator.Equal, 0);

            linkEntity.EntityAlias = "ct";
            linkEntity.Columns = new ColumnSet("fullname", "contactid");

            query.LinkEntities.Add(linkEntity);

            return service.RetrieveMultiple(query).Entities.ToList();
        }

        public static List<Entity> GetAccessRightsForContacts(IOrganizationService service, Guid userId)
        {
            var query = new QueryExpression("contact");

            query.ColumnSet = new ColumnSet("fullname", "ownerid");
            //query.Criteria.AddCondition("contactid", ConditionOperator.Equal, appId);

            var linkEntity = new LinkEntity()
            {
                LinkFromEntityName = "contact",
                LinkToEntityName = "principalobjectaccess",
                LinkFromAttributeName = "contactid",
                LinkToAttributeName = "objectid",
            };

            linkEntity.Columns = new ColumnSet("inheritedaccessrightsmask", "accessrightsmask", "principalobjectaccessid");
            linkEntity.EntityAlias = "RIGHTS";
            query.LinkEntities.Add(linkEntity);

            linkEntity.LinkCriteria.AddCondition("principalid", ConditionOperator.Equal, userId);
            // linkEntity.LinkCriteria.AddCondition("accessrightsmask", ConditionOperator.Equal, 0);

            return service.RetrieveMultiple(query).Entities.ToList();
        }


        public static List<Entity> GetAccessRightsForAppoitments(IOrganizationService service, Guid appId, Guid userId)
        {
            var query = new QueryExpression("appointment");

            query.ColumnSet = new ColumnSet("ownerid", "regardingobjectid");
            query.Criteria.AddCondition("activityid", ConditionOperator.Equal, appId);

            var linkEntity = new LinkEntity()
            {
                LinkFromEntityName = "appointment",
                LinkToEntityName = "principalobjectaccess",
                LinkFromAttributeName = "activityid",
                LinkToAttributeName = "objectid",                
            };

            linkEntity.Columns = new ColumnSet(true);
            linkEntity.EntityAlias = "RIGHTS";
            query.LinkEntities.Add(linkEntity);

            linkEntity.LinkCriteria.AddCondition("principalid", ConditionOperator.Equal, userId);
           // linkEntity.LinkCriteria.AddCondition("accessrightsmask", ConditionOperator.Equal, 0);

            return service.RetrieveMultiple(query).Entities.ToList();
        }


        public static void GetAccessRightForUser(IOrganizationService service, Guid userId)
        {
            QueryExpression query = new QueryExpression();
            query.EntityName = "role";
            query.ColumnSet = new ColumnSet("name");

            LinkEntity systemUseRole = new LinkEntity();
            systemUseRole.LinkFromEntityName = "role";
            systemUseRole.LinkFromAttributeName = "roleid";
            systemUseRole.LinkToEntityName = "systemuserroles";
            systemUseRole.LinkToAttributeName = "roleid";
            systemUseRole.JoinOperator = JoinOperator.Inner;
            systemUseRole.EntityAlias = "SUR";

            LinkEntity userRoles = new LinkEntity();
            userRoles.LinkFromEntityName = "systemuserroles";
            userRoles.LinkFromAttributeName = "systemuserid";
            userRoles.LinkToEntityName = "systemuser";
            userRoles.LinkToAttributeName = "systemuserid";
            userRoles.JoinOperator = JoinOperator.Inner;
            userRoles.EntityAlias = "SU";
            userRoles.Columns = new ColumnSet("fullname");

            LinkEntity rolePrivileges = new LinkEntity();
            rolePrivileges.LinkFromEntityName = "role";
            rolePrivileges.LinkFromAttributeName = "roleid";
            rolePrivileges.LinkToEntityName = "roleprivileges";
            rolePrivileges.LinkToAttributeName = "roleid";
            rolePrivileges.JoinOperator = JoinOperator.Inner;
            rolePrivileges.EntityAlias = "RP";
            rolePrivileges.Columns = new ColumnSet("privilegedepthmask");

            LinkEntity privilege = new LinkEntity();
            privilege.LinkFromEntityName = "roleprivileges";
            privilege.LinkFromAttributeName = "privilegeid";
            privilege.LinkToEntityName = "privilege";
            privilege.LinkToAttributeName = "privilegeid";
            privilege.JoinOperator = JoinOperator.Inner;
            privilege.EntityAlias = "P";
            privilege.Columns = new ColumnSet("name", "accessright");

            LinkEntity privilegeObjectTypeCodes = new LinkEntity();
            privilegeObjectTypeCodes.LinkFromEntityName = "roleprivileges";
            privilegeObjectTypeCodes.LinkFromAttributeName = "privilegeid";
            privilegeObjectTypeCodes.LinkToEntityName = "privilegeobjecttypecodes";
            privilegeObjectTypeCodes.LinkToAttributeName = "privilegeid";
            privilegeObjectTypeCodes.JoinOperator = JoinOperator.Inner;
            privilegeObjectTypeCodes.EntityAlias = "POTC";
            privilegeObjectTypeCodes.Columns = new ColumnSet("objecttypecode");

            ConditionExpression conditionExpression = new ConditionExpression();
            conditionExpression.AttributeName = "systemuserid";
            conditionExpression.Operator = ConditionOperator.Equal;
            conditionExpression.Values.Add(userId);

            userRoles.LinkCriteria = new FilterExpression();
            userRoles.LinkCriteria.Conditions.Add(conditionExpression);

            systemUseRole.LinkEntities.Add(userRoles);
            query.LinkEntities.Add(systemUseRole);

            rolePrivileges.LinkEntities.Add(privilege);
            rolePrivileges.LinkEntities.Add(privilegeObjectTypeCodes);
            query.LinkEntities.Add(rolePrivileges);


            EntityCollection retUserRoles = service.RetrieveMultiple(query);

            Console.WriteLine("Retrieved {0} records", retUserRoles.Entities.Count);
            foreach (Entity rur in retUserRoles.Entities)
            {
                string UserName = String.Empty;
                string SecurityRoleName = String.Empty;
                string PriviligeName = String.Empty;
                string AccessLevel = String.Empty;
                string SecurityLevel = String.Empty;
                string EntityName = String.Empty;

                UserName = ((AliasedValue)(rur["SU.fullname"])).Value.ToString();
                SecurityRoleName = (string)rur["name"];
                EntityName = ((AliasedValue)(rur["POTC.objecttypecode"])).Value.ToString();
                PriviligeName = ((AliasedValue)(rur["P.name"])).Value.ToString();



                switch (((AliasedValue)(rur["P.accessright"])).Value.ToString())
                {
                    case "1":
                        AccessLevel = "READ";
                        break;

                    case "2":
                        AccessLevel = "WRITE";
                        break;

                    case "4":
                        AccessLevel = "APPEND";
                        break;

                    case "16":
                        AccessLevel = "APPENDTO";
                        break;

                    case "32":
                        AccessLevel = "CREATE";
                        break;

                    case "65536":
                        AccessLevel = "DELETE";
                        break;

                    case "262144":
                        AccessLevel = "SHARE";
                        break;

                    case "524288":
                        AccessLevel = "ASSIGN";
                        break;

                    default:
                        AccessLevel = "";
                        break;
                }



                switch (((AliasedValue)(rur["RP.privilegedepthmask"])).Value.ToString())
                {
                    case "1":
                        SecurityLevel = "User";
                        break;

                    case "2":
                        SecurityLevel = "Business Unit";
                        break;

                    case "4":
                        SecurityLevel = "Parent: Child Business Unit";
                        break;

                    case "8":
                        SecurityLevel = "Organisation";
                        break;

                    default:
                        SecurityLevel = "";
                        break;
                }


                Console.WriteLine("User name:" + ((AliasedValue)rur["SU.fullname"]).Value);
                Console.WriteLine("Security Role name:" + rur["name"]);
                Console.WriteLine("Privilige name:" + ((AliasedValue)rur["P.name"]).Value);
                Console.WriteLine("Access Right :" + ((AliasedValue)rur["P.accessright"]).Value);
                Console.WriteLine("Security Level:" + ((AliasedValue)rur["RP.privilegedepthmask"]).Value);

            }
        }
    }
}
