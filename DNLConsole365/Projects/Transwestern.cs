using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;

namespace DNLConsole365.Projects
{
    public class Transwestern
    {
        #region Text Constants - Entity Attributes names
        // Transwestern RBJ - Team ID
        //private const string transwesternTeamId = "D6A27CEC-F945-E211-AA1B-00155D01ED23";
        public const string transwesternTeamId = "4A31CCDA-3689-E311-AF3B-00155D01EE65";
        public const string nonameCompanyId = "3D60EAA2-971D-E611-88FF-00155D0A2724";

        // Connection roles
        public const string contributorRoleId = "3357CAF5-6D14-E111-A170-000C290B0E27";
        public const string contributedByRoleId = "2BB209D2-6D14-E111-A170-000C290B0E27";

        // Bullhorn
        public const string note_bullhorn_synchronized = "bullhorn";

        // Contact entity fields
        public const string contact_id = "contactid";
        public const string contact_firstname = "firstname";
        public const string contact_lastname = "lastname";
        public const string contact_middlename = "middlename";
        public const string contact_email = "emailaddress1";
        public const string contact_email2 = "emailaddress2";
        public const string contact_salutation = "salutation";
        public const string contact_jobtitle = "jobtitle";
        public const string contact_parentcustomerid = "parentcustomerid";
        public const string contact_mobilephone = "mobilephone";
        public const string contact_spousename = "spousesname";
        public const string contact_birthdate = "birthdate";
        public const string contact_anniversary = "anniversary";
        public const string contact_jacketsize = "dnl_jacketsize";
        public const string contact_country = "address1_country";

        //Company enitity fields
        public const string company_id = "accountid";
        public const string company_name = "name";
        public const string company_otherphone = "telephone2";
        public const string company_industrycode = "industrycode";
        public const string company_idustrycodetext = "industrycodetext";
        public const string company_parentaccountid = "parentaccountid";
        public const string company_primarycontactid = "primarycontactid";
        public const string company_websiteurl = "websiteurl";

        //Note entity fields
        public const string note_id = "annotationid";
        //public const string note_modifiedby = "modifiedby";
        public const string note_text = "notetext";
        public const string note_objectid = "objectid";
        public const string note_objectypecode = "objecttypecode";
        public const string note_subject = "subject";
        public const string note_mimetype = "mimetype";

        //Task entity fields
        public const string task_id = "activityid";
        public const string task_regardningobjectid = "regardingobjectid";
        public const string task_description = "description";
        //public const string task_modifiedby = "modifiedby";
        public const string task_subject = "subject";
        // public const string task_duration = "actualdurationminutes";
        public const string task_duedate = "scheduledend";
        public const string task_priotity = "prioritycode";
        public const string task_category = "category";

        // Common entity fields
        //public const string contributors = "awx_contributors";
        public const string statecode = "statecode";
        public const string statecodetext = "statecodetext";
        public const string modifiedon = "modifiedon";
        public const string modifiedby = "modifiedby";
        public const string address_name = "address1_name";
        public const string address_line1 = "address1_line1";
        public const string address_line2 = "address1_line2";
        public const string address_city = "address1_city";
        public const string address_stateprovince = "address1_stateorprovince";
        public const string address_postalcode = "address1_postalcode";
        public const string fax = "fax";
        public const string mainphone = "telephone1";
        public const string leaseexpiration = "dnl_leaseexpiration";
        public const string areamaximum = "dnl_areamaximum";
        public const string description = "description";
        public const string owner = "ownerid";

        // Connection and connection role
        public const string connection_name = "name";
        public const string connection_record1id = "record1id";
        public const string connection_record2id = "record2id";
        public const string connection_record1roleid = "record1roleid";
        public const string connection_record2roleid = "record2roleid";


        // Additional query fields (teammembership and systemuser)
        public const string team_id = "teamid";
        public const string team_name = "name";
        public const string systemuser_id = "systemuserid";
        public const string systemuser_fullname = "fullname";
        public const string systemuser_firstname = "firstname";
        public const string systemuser_lastname = "lastname";

        // Custom fields
        public const string contributors = "contributors";
        //public const string bullhorn_synchronized = "dnl_bhsynchronized";
        public const string bullhorn_synced = "dnl_syncedwithbullhorn";
        public const string bullhorn_id = "dnl_bhid";
        public const string bullhordmodifiedon = "dnl_bhmodifiedon";
        #endregion

        private static QueryExpression GetContactsAndCompanies(string type)
        {
            QueryExpression query = new QueryExpression();
            query.EntityName = type;
            query.ColumnSet = new ColumnSet();

            LinkEntity linkConnections = new LinkEntity();

            // Link contacts to connection entity
            if (type == "contact")
            {
                linkConnections.LinkFromAttributeName = contact_id;
            }
            else if (type == "account")
            {
                // Link companies to connection entity
                linkConnections.LinkFromAttributeName = company_id;
            }
            else if (type == "task")
            {
                linkConnections.LinkFromAttributeName = task_regardningobjectid;
            }
            else if (type == "annotation")
            {
                linkConnections.LinkFromAttributeName = note_objectid;
            }

            // Set link attributes
            linkConnections.LinkFromEntityName = query.EntityName;
            linkConnections.LinkToAttributeName = connection_record1id;
            linkConnections.LinkToEntityName = "connection";
            // linkConnections.Columns = new ColumnSet(connection_name, connection_record2id);
            //linkConnections.EntityAlias = "connection";

            // Link connections entities with teams
            LinkEntity linkTeams = new LinkEntity
            {
                LinkFromAttributeName = connection_record2id,
                LinkFromEntityName = "connection",
                LinkToAttributeName = systemuser_id,
                LinkToEntityName = "teammembership"
            };

            // Check teamid
            linkTeams.LinkCriteria.AddCondition(team_id, ConditionOperator.Equal, transwesternTeamId);
            linkConnections.LinkEntities.Add(linkTeams);
            query.LinkEntities.Add(linkConnections);

            // Check if from and to dates available - if not - use default behavior and select records for last seven days
            /*if ((from == 0) || (to == 0))
            {
                // Select all records if count is specified
                if (count == 0)
                {
                    query.Criteria.AddCondition(modifiedon, ConditionOperator.Last7Days);
                }
            }
            else
            {
                query.Criteria.AddCondition(modifiedon, ConditionOperator.Between, new object[] { DateTimeHelper.UnixTimeStampToDateTime(from),
                                                                                                  DateTimeHelper.UnixTimeStampToDateTime(to) });
            }*/

            // Additional check - select record if it has not been synchronized with bullhorn.            
            // query.Criteria.AddCondition(bullhorn_synced, ConditionOperator.NotEqual, true);

            /* query.Criteria.AddCondition(modifiedon, ConditionOperator.Between, new object[] { new DateTime(2016, 05, 28, 0, 0, 0, DateTimeKind.Utc),
                                                                                               new DateTime(2016, 05, 28, 23, 59, 59, DateTimeKind.Utc)});*/

            // Return all records
            // Note entity does not have state code
            if (type != "annotation")
            {
                // Select only active entities - statecode = 0
                query.Criteria.AddCondition(statecode, ConditionOperator.Equal, 0);
            }

            // Select only unique enities
            query.Distinct = true;

            return query;
        }


        public static EntityCollection GetEntities(IOrganizationService service, string logicalName)
        {
            var query = GetContactsAndCompanies(logicalName);

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

            return entityCollection;
        }


        public static List<Entity> GetActiveUsersByBusinessUnit(IOrganizationService service, Guid buId)
        {
            var query = new QueryExpression("systemuser");
            query.ColumnSet = new ColumnSet("firstname", "lastname", "fullname");
            query.Criteria.AddCondition("businessunitid", ConditionOperator.Equal, buId);
            
            return service.RetrieveMultiple(query).Entities.ToList();
        }

        public static List<Entity> GetEntitiesByOwner(IOrganizationService service, string logicalName, Guid ownerId)
        {
            var query = new QueryExpression(logicalName);
            query.ColumnSet = new ColumnSet(false);
            query.Criteria.AddCondition("ownerid", ConditionOperator.Equal, ownerId);

            return service.RetrieveMultiple(query).Entities.ToList();
        }
    }
}
