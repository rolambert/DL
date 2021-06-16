using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Json;

using System.ServiceModel;
using System.ServiceModel.Description;

using DNLConsole365.Service;
using DNLConsole365.Projects;
using DNLConsole365.Sharepoint;

using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Tooling;

using Bytescout.Spreadsheet;
using System.Net;


namespace DNLConsole365
{
    class Program
    {

        public struct EntityMeta
        {
            public string name;
            public string oldValue;
            public string newValue;

            public EntityMeta(string n, string ov, string nv)
            {
                name = n;
                oldValue = ov;
                newValue = nv;
            }
        }

        public class EntityCount
        {
            public string EntityName { get; set; }
            public string SchemaName { get; set; }
            public int TotalCount { get; set; }
            public int ActiveCount { get; set; }
        }


        #region northrop 

        public static void BulkDeleteEntities(DataCollection<Entity> entityCollection)
        {

        }

        public static RetrieveAllEntitiesResponse RetrieveEntities(IOrganizationService service)
        {
            var req = new RetrieveAllEntitiesRequest();
            req.EntityFilters = EntityFilters.Entity;
            req.RetrieveAsIfPublished = true;

            return (RetrieveAllEntitiesResponse)service.Execute(req);
        }

        public static DataCollection<Entity> GetEntitiesCount(IOrganizationService service, string logicalName)
        {
            Console.WriteLine(string.Format("ENTITY RETRIEVE - {0}", logicalName));
            var query = new QueryExpression(logicalName);

            query.ColumnSet = new ColumnSet(true);

            //query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
            
            query.PageInfo.PagingCookie = null;
            int pageNumber = 1;
            var tempResult = new EntityCollection();
            var entityCollection = new EntityCollection();

            do
            {

                Console.WriteLine(string.Format("{0}-{1}", logicalName, pageNumber));
                query.PageInfo.Count = 5000;
                query.PageInfo.PageNumber = pageNumber++;
                query.PageInfo.PagingCookie = tempResult.PagingCookie;

                tempResult = service.RetrieveMultiple(query);
                entityCollection.Entities.AddRange(tempResult.Entities);
            }
            while (tempResult.MoreRecords);

            /*try
            {

                do
                {
                    
                    Console.WriteLine(string.Format("{0}-{1}", logicalName, pageNumber));
                    query.PageInfo.Count = 5000;
                    query.PageInfo.PageNumber = pageNumber++;
                    query.PageInfo.PagingCookie = tempResult.PagingCookie;

                    var tempResult = service.RetrieveMultiple(query);
                    entityCollection.Entities.AddRange(tempResult.Entities);
                }
                while (tempResult.MoreRecords);
            }
            catch (Exception ex)
            {
                //query.ColumnSet = new 
            }*/

            return entityCollection.Entities;
        }

        public static void CreateEntityNamesDoc(IOrganizationService service, DataCollection<Entity> entities)
        {
            Spreadsheet document = new Spreadsheet();

            // add new worksheet
            Worksheet Sheet = document.Workbook.Worksheets.Add("NorthRopSolutions");

            Sheet.Cell("A1").Value = "Solution Name";
            Sheet.Columns[0].Width = 250;
            Sheet.Cell("B1").Value = "Description";
            Sheet.Columns[1].Width = 500;
            
            foreach (var ent in entities)
            {

            }

            var rowNumber = 1;
            foreach (var ec in entities)
            {
                Sheet.Cell(rowNumber, 0).Value = ec["friendlyname"].ToString();
                Sheet.Cell(rowNumber, 1).Value = ec["description"].ToString();                
                rowNumber++;
            }

            // delete output file if exists already
            if (File.Exists("NorthRopSolutions.xls"))
            {
                File.Delete("NorthRopSolutions.xls");
            }

            // Save document
            document.SaveAs("NorthRopSolutions.xls");

            // Close Spreadsheet
            document.Close();
        }

        public static void CreateEntityCountDoc(IOrganizationService service, RetrieveAllEntitiesResponse response)
        {
            var entityList = new List<EntityCount>();

            foreach (var item in response.EntityMetadata)
            {
                var entities = GetEntitiesCount(service, item.LogicalName);

                // If No state available - set -5
                var activeCount = 0;
                var totalCount = entities.Count;

                Console.WriteLine(string.Format("ENTITY TOTAL - {0} - {1}", item.SchemaName, totalCount));

                if (totalCount > 0)
                {
                    if (entities.FirstOrDefault().Contains("statecode"))
                    {
                        activeCount = entities.Where(e => ((OptionSetValue)e["statecode"]).Value == 0).ToList().Count;
                    }
                    else
                    {
                        activeCount = -1;
                    }
                }

                var itemDisplayName = item.DisplayName.LocalizedLabels.FirstOrDefault();

                entityList.Add(new EntityCount()
                {
                    EntityName = itemDisplayName != null ? itemDisplayName.Label : "NONE",
                    SchemaName = item.LogicalName,
                    TotalCount = totalCount,
                    ActiveCount = activeCount
                });
            }


            Spreadsheet document = new Spreadsheet();

            // add new worksheet
            Worksheet Sheet = document.Workbook.Worksheets.Add("NorthRopEntityCount");

            Sheet.Cell("A1").Value = "Entity Name";
            Sheet.Columns[0].Width = 250;
            Sheet.Cell("B1").Value = "Schema Name";
            Sheet.Columns[0].Width = 250;
            Sheet.Cell("C1").Value = "Total Count";
            Sheet.Columns[1].Width = 150;
            Sheet.Cell("D1").Value = "Active Count";
            Sheet.Columns[1].Width = 150;


            var rowNumber = 1;
            foreach (var ec in entityList)
            {
                Sheet.Cell(rowNumber, 0).Value = ec.EntityName;
                Sheet.Cell(rowNumber, 1).Value = ec.SchemaName;
                Sheet.Cell(rowNumber, 2).Value = ec.TotalCount;
                Sheet.Cell(rowNumber, 3).Value = ec.ActiveCount;
                rowNumber++;
            }

            // delete output file if exists already
            if (File.Exists("NorthRopEntitiesCount.xls"))
            {
                File.Delete("NorthRopEntitiesCount.xls");
            }

            // Save document
            document.SaveAs("NorthRopEntitiesCount.xls");

            // Close Spreadsheet
            document.Close();
        }
        #endregion

        #region HuntonGroup

        public static void FixOpportunityProducts(IOrganizationService service, Guid opportunityDetailId)
        {
            var query = new QueryExpression("opportunityproduct");
            query.ColumnSet = new ColumnSet(true);
            query.Criteria.AddCondition("opportunityproductid", ConditionOperator.Equal, opportunityDetailId);

            

            var oppDetail = service.RetrieveMultiple(query).Entities.FirstOrDefault();//service.Retrieve("opportunityproduct", opportunityDetailId, new ColumnSet("uomid", "productid"));
            var oppDetailUomId = ((EntityReference)oppDetail["uomid"]).Id;
            
            var product = service.Retrieve("product", ((EntityReference)oppDetail["productid"]).Id, new ColumnSet("defaultuomid"));
            var productUomId = ((EntityReference)product["defaultuomid"]).Id;
            
            if (oppDetailUomId != productUomId)
            {
                oppDetail.Id = Guid.Empty;
                oppDetail["uomid"] = new EntityReference("uom", productUomId);
                oppDetail.Attributes.Remove("opportunityproductid");

                service.Create(oppDetail);
            }


        }

        /*public static void Hunton_SetupUsersDivisions()
        {
            var users2016 = Hunton.GetActiveUsers(OrganizationService.Instance.GetService(Guid.Empty));

            var users365 = Hunton.GetActiveUsers(OrganizationService2.Instance.GetService());


            Console.WriteLine("HT - {0}", users2016.Where(e => e.Contains("new_division") && ((OptionSetValue)e["new_division"]).Value == 1).ToList().Count);
            Console.WriteLine("BAS - {0}", users2016.Where(e => e.Contains("new_division") && ((OptionSetValue)e["new_division"]).Value == 2).ToList().Count);
            Console.WriteLine("HTS - {0}", users2016.Where(e => e.Contains("new_division") && ((OptionSetValue)e["new_division"]).Value == 3).ToList().Count);
            Console.WriteLine("HSP - {0}", users2016.Where(e => e.Contains("new_division") && ((OptionSetValue)e["new_division"]).Value == 4).ToList().Count);

            foreach (var user in users2016.Where(e => e.Contains("new_division")).ToList())
            {
                var userMatch = users365.Where(c => c["firstname"].ToString().Trim().ToLower() == user["firstname"].ToString().Trim().ToLower() &&
                                                    c["lastname"].ToString().Trim().ToLower() == user["lastname"].ToString().Trim().ToLower()).FirstOrDefault();
                if (userMatch != null)
                {
                    Console.WriteLine("MATCH: " + userMatch["firstname"].ToString() + " " + userMatch["lastname"].ToString());

                    if (userMatch.Contains("new_division"))
                    {
                        Console.WriteLine("HAS: {0}[{1}]", userMatch.FormattedValues.Contains("new_division") ?
                                                            userMatch.FormattedValues["new_division"].ToString() : "NAN",
                                                            ((OptionSetValue)userMatch["new_division"]).Value);
                    }
                    else
                    {
                        userMatch["new_division"] = new OptionSetValue(((OptionSetValue)user["new_division"]).Value);
                        Console.Write("SET: {0}[{1}]", user.FormattedValues["new_division"].ToString(), ((OptionSetValue)user["new_division"]).Value);
                        OrganizationService2.Instance.GetService().Update(userMatch);
                        Console.WriteLine(".........Updated!");
                    }


                }
                else
                {
                    Console.WriteLine("NO MATCH: " + user["firstname"].ToString() + " " + user["lastname"].ToString());
                }
            }
        }*/



        public static GrantAccessResponse Hunton_GrantAccess(OrganizationCredentials creds, Guid userId, EntityReference entity, EntityReference owner, bool isOwner = false)
        {
            var userReference = new EntityReference("systemuser", userId);
            var grantAccessRequest = new GrantAccessRequest
            {
                PrincipalAccess = new PrincipalAccess
                {
                    AccessMask = isOwner ? AccessRights.ReadAccess | 
                                            AccessRights.ShareAccess | 
                                            AccessRights.CreateAccess | 
                                            AccessRights.DeleteAccess | 
                                            AccessRights.AssignAccess : AccessRights.ReadAccess | 
                                                                        AccessRights.ShareAccess,
                    Principal = userReference
                },
                Target = entity
            };

            return (GrantAccessResponse)OrganizationService.Instance.GetService(creds, owner.Id).Execute(grantAccessRequest);
        }

        public static AssignResponse Hunton_AssignViews(OrganizationCredentials creads, Guid userId, Guid viewId, Guid ownerId)
        {
            AssignRequest assignRequest = new AssignRequest
            {
                Assignee = new EntityReference
                {
                    LogicalName = "systemuser",
                    // Here we could assign the visualization to the newly created user
                    Id = userId
                },

                Target = new EntityReference
                {
                    LogicalName = "userview",
                    Id = viewId
                }
            };

            return (AssignResponse)OrganizationService.Instance.GetService(creads, ownerId).Execute(assignRequest);
        }


        #region DashBoards

        public static AssignResponse Hunton_AssignDashboard(OrganizationCredentials creads, Guid userId, Guid dashboardId, Guid ownerId)
        {
            AssignRequest assignRequest = new AssignRequest
            {
                Assignee = new EntityReference
                {
                    LogicalName = "systemuser",
                    // Here we could assign the visualization to the newly created user
                    Id = userId
                },

                Target = new EntityReference
                {
                    LogicalName = "userform",
                    Id = dashboardId
                }
            };

            return (AssignResponse)OrganizationService.Instance.GetService(creads, ownerId).Execute(assignRequest);
        }

        public static void Hunton_ReassignDashboards(OrganizationCredentials cred2016, OrganizationCredentials cred365)
        {
            // Get users from CRMs
            var users2016 = Hunton.GetActiveUsers(OrganizationService.Instance.GetService(cred2016, Guid.Empty));
            var users365 = Hunton.GetActiveUsersWithSourceId(OrganizationService.Instance.GetService(cred365, Guid.Empty));

            // Get Dynamics User 365 Dashboards for Dynamics
            var userDynamicsDashboards365 = Hunton.GetUserDashboards(OrganizationService.Instance.GetService(cred365, Guid.Empty));

            foreach (var user in users2016.OrderBy(e => e["fullname"]).ToList())
            {
                try
                {

                    // Get 2016 Dashboards for 2016 User
                    var userDashboards2016 = Hunton.GetUserDashboardsByOwnerId(OrganizationService.Instance.GetService(cred2016, user.Id), user.Id);

                    Console.WriteLine("2016::{0} - OWNER OF {1}", user["fullname"].ToString(), userDashboards2016.Count);
                    if (userDashboards2016.Count == 0) continue;


                    // Check for Dynamics 365 User
                    if (users365.Any(u => u["new_sourceid"].ToString() == user.Id.ToString()))
                    {
                        // Get Dynamics 365 User
                        var user365 = users365.Where(u => u["new_sourceid"].ToString() == user.Id.ToString()).FirstOrDefault();

                        // Get Dynamics 365 Dashboards
                        var userDashboards365 = Hunton.GetUserDashboardsByOwnerId(OrganizationService.Instance.GetService(cred365, user365.Id), user365.Id);

                        Console.WriteLine("365::{0} - OWNER OF {1}", user365["fullname"].ToString(), userDashboards365.Count);

                        if (userDashboards365.Count != userDashboards2016.Count)
                        {
                            foreach (var d365 in userDashboards365)
                            {
                                if (!userDashboards2016.Any(d => d.Id == d365.Id))
                                {
                                    var resp = Hunton_AssignDashboard(cred365, new Guid("EA149F3E-6020-E711-8108-5065F38AF901"), d365.Id, ((EntityReference)d365["ownerid"]).Id);
                                }
                            }

                        }
                        else
                        {
                            continue;
                        }



                        // Walk through 2016 dashboards
                        foreach (var dashboard in userDashboards2016)
                        {
                            // Check if dashboard shared for user in 365
                            var dashboard365 = userDashboards365.Where(d => d["name"].ToString() == dashboard["name"].ToString() && d.Id == dashboard.Id).FirstOrDefault();

                            if (dashboard365 == null)
                            {
                                // REASIGN
                                //Console.WriteLine("365::{0} - NO valid owner", dashboard["name"].ToString());

                                var dynamicsDashboard365 = userDynamicsDashboards365.Where(d => d["name"].ToString() == dashboard["name"].ToString()).FirstOrDefault();

                                if (dynamicsDashboard365 != null)
                                {
                                    var resp = Hunton_AssignDashboard(cred365, user365.Id, dynamicsDashboard365.Id, Guid.Empty);
                                    Console.WriteLine("365::{0} reassigned to {2} - {1}", dynamicsDashboard365["name"].ToString(), resp.ResponseName, user365["fullname"].ToString());
                                    //Console.WriteLine("365::{0} - REASIGN valid owner", dashboard["name"].ToString());
                                }
                                else
                                {
                                    // Console.WriteLine("ERROR::{0} - NO DASH AT ALL", dashboard["name"].ToString());
                                }

                            }
                            else
                            {
                                Console.WriteLine("365::{0} - ALREADY has valid owner", dashboard365["name"].ToString());
                            }

                        }
                    }
                    else
                    {
                        Console.WriteLine("NO USER - " + user["fullname"].ToString());
                    }
                    // Get Dynamics User 365 Dashboards for Dynamics




                }
                catch (Exception ex)
                {
                    WriteLog(string.Format("USER - {0}\r\nERROR: {1}", user["fullname"].ToString(), ex.Message), "Errors.txt");
                }
            }
        }

        public static void Hunton_ShareUsersDasboards(OrganizationCredentials cred2016, OrganizationCredentials cred365)
        {
            // Get users from CRMs
            var users2016 = Hunton.GetActiveUsers(OrganizationService.Instance.GetService(cred2016, Guid.Empty));                
            var users365 = Hunton.GetActiveUsersWithSourceId(OrganizationService.Instance.GetService(cred365, Guid.Empty));

            // Get Dynamics User 365 Dashboards for Dynamics
            var userDynamicsDashboards365 = Hunton.GetUserDashboards(OrganizationService.Instance.GetService(cred365, Guid.Empty));

            var counter = 0;
            var overall = users2016.Count;

            foreach (var user in users2016.OrderBy(u => u["fullname"].ToString()).ToList())
            {
                try
                {
                    counter++;
                    // Get 2016 Dashboards for 2016 User
                    var userDashboards2016 = Hunton.GetUserDashboards(OrganizationService.Instance.GetService(cred2016, user.Id));

                    Console.WriteLine();
                    Console.WriteLine("{0} of {1}", counter, overall);
                    Console.WriteLine("2016::{0} - {1}", user["fullname"].ToString(), userDashboards2016.Count);
                    if (userDashboards2016.Count == 0) continue;

                    // Check for Dynamics 365 User
                    if (users365.Any(u => u["new_sourceid"].ToString() == user.Id.ToString()))
                    {
                        // Get Dynamics 365 User
                        var user365 = users365.Where(u => u["new_sourceid"].ToString() == user.Id.ToString()).FirstOrDefault();

                        // Get 365 Dashboard for User 365
                        var userInitialDashboards365 = Hunton.GetUserDashboards(OrganizationService.Instance.GetService(cred365, user365.Id));
                        Console.WriteLine("365::{0} - {1}", user365["fullname"].ToString(), userInitialDashboards365.Count);

                        // Walk through 2016 dashboards
                        foreach (var dashboard in userDashboards2016)
                        {

                            // Check if dashboard shared for user in 365
                            var dashboard365 = userInitialDashboards365.Where(d => d["name"].ToString() == dashboard["name"].ToString() && d.Id == dashboard.Id).FirstOrDefault();

                            // If no
                            if (dashboard365 == null)
                            {
                                // Search 365 Dashboards by owner
                                var owner365 = users365.Where(u => u["new_sourceid"].ToString() == ((EntityReference)dashboard["ownerid"]).Id.ToString()).FirstOrDefault();
                                var userDashboards365 = new List<Entity>();
                                if (owner365 != null)
                                {
                                    userDashboards365 = Hunton.GetUserDashboards(OrganizationService.Instance.GetService(cred365, owner365.Id));
                                }

                                // Seach in Owners Dashboards
                                dashboard365 = userDashboards365.Where(d => d["name"].ToString() == dashboard["name"].ToString() && d.Id == dashboard.Id).FirstOrDefault();

                                // Search in Dynamics Dashboards
                                if (dashboard365 == null)
                                {
                                    dashboard365 = userDynamicsDashboards365.Where(d => d["name"].ToString() == dashboard["name"].ToString() && d.Id == dashboard.Id).FirstOrDefault();
                                    if (dashboard365 == null)
                                    {
                                        Console.WriteLine("NO DASHBOARD AT ALL!!! NEED TO CREATE!");

                                    }
                                    else
                                    {
                                        Console.WriteLine("{0} - Dynamics Owner - Share as Is", dashboard365["name"].ToString());
                                        var resp = Hunton_GrantAccess(cred365, user365.Id, new EntityReference(dashboard365.LogicalName, dashboard365.Id), (EntityReference)dashboard365["ownerid"]);
                                        //Console.WriteLine("REASIGN??? AND SHARE????");
                                    }
                                }
                                else
                                {
                                    // Share
                                    Console.WriteLine("{0} - Need to share by Owner", dashboard365["name"].ToString());
                                    //    Console.WriteLine("GRANT for - YES DASHBOARD");
                                    var resp = Hunton_GrantAccess(cred365, user365.Id, new EntityReference(dashboard365.LogicalName, dashboard365.Id), (EntityReference)dashboard365["ownerid"]);
                                }
                            }
                            else
                            {
                                Console.WriteLine("{0} - Already Shared or Owner", dashboard365["name"].ToString());
                                // Nothing to do!!!
                            }

                            /*var userForm = new Entity("userform");
                            userForm["description"] = dashboard["description"];
                            userForm["formxml"] = dashboard["formxml"];
                            userForm["istabletenabled"] = dashboard["istabletenabled"];
                            userForm["name"] = dashboard["name"];
                            userForm["objecttypecode"] = dashboard["objecttypecode"];
                            userForm["type"] = dashboard["type"];
                            userForm["userformid"] = dashboard["userformid"];

                            OrganizationService2.Instance.GetService().Create(userForm);*/

                            //WriteLog(dashboard["name"].ToString(), "CreatedDashboards.txt");
                            //                        Console.WriteLine("BAD!");


                            /*Console.WriteLine("{0} - {1}",
                                            dashboard["name"].ToString(),
                                            userDynamicsDashboards365.Any(d => d["name"].ToString() == dashboard["name"].ToString()));*/

                            //Console.WriteLine();
                        }
                    }
                    else
                    {
                        Console.WriteLine("NO USER - " + user["fullname"].ToString());
                    }

                }
                catch (Exception ex)
                {
                    WriteLog(string.Format("USER - {0}\r\nERROR: {1}", user["fullname"].ToString(), ex.Message), "Errors.txt");
                }
            }
        }
        #endregion

        public static void Hunton_CreateAndAssignViews(OrganizationCredentials cred2016, OrganizationCredentials cred365)
        {
            var users365 = Hunton.GetActiveUsersWithSourceId(OrganizationService.Instance.GetService(cred365, Guid.Empty));
            var users2016 = Hunton.GetActiveUsers(OrganizationService.Instance.GetService(cred2016, Guid.Empty));

            var counter = 1;
            var overall = users2016.Count;

            foreach (var user2016 in users2016.OrderBy(e => e["fullname"]).ToList())
            {
                try
                {
                    // Get Dynamics 2016 Views
                    var userViews2016 = Hunton.GetUserViewsByOwnerId(OrganizationService.Instance.GetService(cred2016, user2016.Id), user2016.Id);

                    Console.WriteLine();
                    Console.WriteLine("{0} of {1}", counter, overall);
                    Console.WriteLine("{0} - OWNER OF {1}", user2016["fullname"].ToString(), userViews2016.Count);
                    counter++;

                    if (userViews2016.Count == 0) continue;

                    var cc = 1;
                    var ovcc = userViews2016.Count;

                    // Check for Dynamics 365 User
                    var user365 = users365.Where(u => u["new_sourceid"].ToString() == user2016.Id.ToString()).FirstOrDefault();

                    foreach (var userView2016 in userViews2016)
                    {
                        userView2016.Attributes.Remove("owningbusinessunit");
                        userView2016.Attributes.Remove("owninguser");
                        userView2016.Attributes.Remove("createdby");
                        userView2016.Attributes.Remove("createdon");
                        userView2016.Attributes.Remove("modifiedon");
                        userView2016.Attributes.Remove("modifiedby");
                        if (user365 != null)
                        {
                            Console.WriteLine("365::{1}-{2}::[{0}][{3}] - CREATE", userView2016["name"].ToString(), cc, ovcc, user365["fullname"].ToString());
                            userView2016["ownerid"] = new EntityReference(user365.LogicalName, user365.Id);
                        }
                        else
                        {
                            // Create UserView for Dynamics
                            Console.WriteLine("365::{1}-{2}::[{0}][{3}] - CREATE DYNAMICS", userView2016["name"].ToString(), cc, ovcc, "Dynamics #");

                            userView2016.Attributes.Remove("ownerid");
                            userView2016.Attributes.Remove("owneridtype");                   
                        }

                        try
                        {
                            var newId = Hunton.CreateUserView(OrganizationService.Instance.GetService(cred365, user365 != null ? user365.Id : Guid.Empty), userView2016);
                            Console.WriteLine("365::{1}-{2}::[{0}] - ++++++++++++++", userView2016["name"].ToString(), cc, ovcc);
                        }
                        catch(Exception ex)
                        {
                            Console.WriteLine("365::{1}-{2}::[{0}] - !!!!!!!!!!!!!!", userView2016["name"].ToString(), cc, ovcc);
                            WriteLog(string.Format("{2}\r\nERROR: {1}",  ex.Message, userView2016["name"].ToString()), "Errors.txt");
                        }

                        cc++;
                    }                    
                }
                catch (Exception ex)
                {
                    WriteLog(string.Format("USER - {0}\r\nERROR: {1}", user2016["fullname"].ToString(), ex.Message), "Errors.txt");
                }
            }
        }
    

        public static void Hunton_DeleteIncorectViews(OrganizationCredentials cred2016, OrganizationCredentials cred365)
        {
            var users365 = Hunton.GetActiveUsersWithSourceId(OrganizationService.Instance.GetService(cred365, Guid.Empty));

            var userDynamicsDashboards365 = Hunton.GetUserViews(OrganizationService.Instance.GetService(cred365, Guid.Empty));

            var counter = 1;
            var overall = users365.Count;
            foreach (var user365 in users365.OrderBy(e => e["fullname"]).ToList())
            {
                try
                {
                    // Get Dynamics 365 Dashboards
                    var userViews365 = Hunton.GetUserViewsByOwnerId(OrganizationService.Instance.GetService(cred365, user365.Id), user365.Id);

                    Console.WriteLine();
                    Console.WriteLine("{0} of {1}", counter, overall);                    
                    Console.WriteLine("{0} - OWNER OF {1}", user365["fullname"].ToString(), userViews365.Count);
                    counter++;

                    var cc = 0;
                    var ovcc = userViews365.Count;
                    foreach (var userView365 in userViews365)
                    {
                        cc++;
                        try
                        {
                            Console.WriteLine("{1}-{2}::[{0}] - START DELETE", userView365["name"].ToString(), cc, ovcc);
                            OrganizationService.Instance.GetService(cred365, user365.Id).Delete(userView365.LogicalName, userView365.Id);
                            Console.WriteLine("{1}-{2}::[{0}] +++++++++++++++", userView365["name"].ToString(), cc, ovcc);
                        }
                        catch(Exception ex)
                        {
                            Console.WriteLine("{1}-{2}::[{0}] !!!!!!!!!!!!!", userView365["name"].ToString(), cc, ovcc);
                            WriteLog(string.Format("USER - {0}\r\n{2}\r\nERROR: {1}", user365["fullname"].ToString(), ex.Message, userView365["name"].ToString()), "Errors.txt");
                        }
                        
                    }

                }
                catch (Exception ex)
                {
                    WriteLog(string.Format("USER - {0}\r\nERROR: {1}", user365["fullname"].ToString(), ex.Message), "Errors.txt");
                }
            }
        }

        public static void Hunton_FixUserViews(OrganizationCredentials cred365)
        {
            var users365 = Hunton.GetActiveUsersWithSourceId(OrganizationService.Instance.GetService(cred365, Guid.Empty));

            var counter = 0;
            var overall = users365.Count;

            foreach (var user365 in users365.OrderBy(u => u["fullname"].ToString()).ToList()/*.Where(u => u["fullname"].ToString().Contains("Hendrex")).ToList()*/)
            {
                counter++;
                var userViews365 = Hunton.GetUserViews(OrganizationService.Instance.GetService(cred365, user365.Id));
                var userOwnerViews365 = userViews365.Where(v => ((EntityReference)v["ownerid"]).Id == user365.Id).ToList();
                                     //   .Where(uv => uv["name"].ToString() == "Caldwell Opportunities").ToList();

                Console.WriteLine();
                Console.WriteLine("{0} of {1}", counter, overall);
                Console.WriteLine("{0} - SHARED [{1}] OWNER [{2}]", user365["fullname"].ToString(),
                                                        userViews365.Count,
                                                        userOwnerViews365.Count);

                var vc = 0;
                var ovc = userOwnerViews365.Count;

                foreach(var userView365 in userOwnerViews365)
                {
                    vc++;
                    Console.WriteLine("{1}-{2}::{0}", userView365["name"].ToString(), vc, ovc);

                    XDocument xdoc = new XDocument();
                    xdoc = XDocument.Parse(userView365["fetchxml"].ToString());

                    var users = 0;
                    var logs = new List<string>();
                    logs.Add(string.Format("u365[{0}]::v[{1}]", user365["fullname"].ToString(), userView365["name"].ToString()));



                    //foreach (XElement filterElement in xdoc.Element("fetch").Elements("entity").Elements("filter").Elements("condition").Elements("value"))
                    foreach (XElement filterElement in xdoc.Element("fetch").Elements("entity").Elements("filter").Elements("condition"))
                    {
                        //var userAttribute = filterElement.Attribute("uitype");
                        if (filterElement.Attribute("uitype") != null && filterElement.Attribute("uitype").Value == "systemuser" && filterElement.Attribute("value") != null)
                        //if (filterElement.Attribute("uitype") != null && filterElement.Attribute("uitype").Value == "systemuser") 
                        {
                            //    var u2016id = new Guid(filterElement.Value);
                            var u2016id = new Guid(filterElement.Attribute("value").Value);

                            if (users365.Any(u => u.Id.ToString() == u2016id.ToString().ToLower()))
                            {
                                Console.WriteLine(string.Format("VALID::u2016[{0}]::id[{1}]", filterElement.Attribute("uiname"), u2016id));
                                continue;                              
                            }
                            else
                            {
                                Console.WriteLine(string.Format("!!!!NOT VALID::u2016[{0}]::id[{1}]", filterElement.Attribute("uiname"), u2016id));                                
                            }

                            var repUser365 = users365.Where(u => u["new_sourceid"].ToString() == u2016id.ToString().ToLower()).FirstOrDefault();

                            if (repUser365 != null)
                            {
                                users++;
                                logs.Add(string.Format("PRSENT::u2016[{0}]::id[{1}]", filterElement.Attribute("uiname"), u2016id));
                                //filterElement.SetValue("{" + repUser365.Id + "}");
                                filterElement.Attribute("value").SetValue("{" + repUser365.Id + "}");
                                filterElement.Attribute("uiname").SetValue(repUser365["fullname"].ToString());
                            }
                            else
                            {
                                logs.Add(string.Format("MISSING::u2016[{0}]::id[{1}]", filterElement.Attribute("uiname"), u2016id));
                            }
                            
                            //WriteLog(string.Format("u365[{0}]::v[{1}]", user365["fullname"].ToString(), userView365["name"].ToString()), "UserViewsFix.txt");
                            //WriteLog(string.Format("u2016[{0}]::id[{1}]", filterElement.Attribute("uiname"), filterElement.Value), "UserViewsFix.txt");
                            // Console.WriteLine();
                            // Console.WriteLine(filterElement.Attribute("uitype"));
                            //Console.WriteLine(filterElement.Attribute("uiname"));
                            //Console.WriteLine(filterElement.Value);
                        }
                        //Console.WriteLine(filterElement.Name);
                    }

                    Console.WriteLine(users > 0 ? string.Format("{0} - USERS !!!!!!", users) : "SKIP");
                    if(users > 0)
                    {
                        //var updatedDoc = xdoc.ToString();
                        userView365["fetchxml"] = xdoc.ToString();

                        try
                        {
                            Console.WriteLine("{1}-{2}::{0} - UPDATE USERS", userView365["name"].ToString(), vc, ovc);
                            OrganizationService.Instance.GetService(cred365, user365.Id).Update(userView365);
                            Console.WriteLine("{1}-{2}::{0} - UPDATE USERS +++", userView365["name"].ToString(), vc, ovc);
                        }
                        catch(Exception ex)
                        {
                            WriteLog(string.Format("{1}-{2}::{0} - UPDATE USERS FAIL!!!!!!!", userView365["name"].ToString(), vc, ovc), "UserViewsErrorLog.txt");
                            WriteLog(ex.Message, "UserViewsErrorLog.txt");
                            Console.WriteLine("{1}-{2}::{0} - UPDATE USERS FAIL!!!!!!!", userView365["name"].ToString(), vc, ovc);
                        }

                        foreach (var log in logs)
                        {
                            
                            WriteLog(log, "UserViewsLOG.txt");
                        }
                    }

                   // Console.WriteLine("TEST");

                    /*foreach (XElement item in xdoc)
                    {
                        Console.WriteLine("Department Name - " + item.Value);
                    }*/


                }


            }
        }

        public static void Hunton_ShareViews(OrganizationCredentials cred2016, OrganizationCredentials cred365)
        {
            var users365 = Hunton.GetActiveUsersWithSourceId(OrganizationService.Instance.GetService(cred365, Guid.Empty));
            var users2016 = Hunton.GetActiveUsers(OrganizationService.Instance.GetService(cred2016, Guid.Empty));
             //   .Where(u => u["fullname"].ToString().Contains("Slade")).ToList();
                //.Where(u => u["lastname"].ToString() == "Sutter" || u["lastname"].ToString() == "Mcmeans" || u["lastname"].ToString() == "Beck").ToList(); ;

            // Get Dynamics User 365 Dashboards for Dynamics
            var userDynamicsViews365 = Hunton.GetUserViews(OrganizationService.Instance.GetService(cred365, Guid.Empty));

            var counter = 0;
            var overall = users2016.Count;

            foreach (var user2016 in users2016.OrderBy(e => e["fullname"]).ToList())
            {
                try
                {
                    counter++;
                    // Get Dynamics 365 Dashboards
                    var userViews2016 = Hunton.GetUserViews(OrganizationService.Instance.GetService(cred2016, user2016.Id)).Where(v => !v["name"].ToString().Contains("Cobalt")).ToList();

                    Console.WriteLine();
                    Console.WriteLine("{0} of {1}", counter, overall);
                    Console.WriteLine("{0} - SHARED [{1}] OWNER [{2}]", user2016["fullname"].ToString(), 
                                                            userViews2016.Count, 
                                                            userViews2016.Where(v => ((EntityReference)v["ownerid"]).Id == user2016.Id).ToList().Count);
                    

                    if (userViews2016.Count == 0) continue;

                    
                    // Check for Dynamics 365 User
                    var user365 = users365.Where(u => u["new_sourceid"].ToString().ToLower() == user2016.Id.ToString().ToLower()).FirstOrDefault();

                    if (user365 != null)
                    {
                        var userInitialViews365 = Hunton.GetUserViews(OrganizationService.Instance.GetService(cred365, user365.Id));

                        var cc = 1;
                        var ovcc = userViews2016.Count;

                        foreach (var userView2016 in userViews2016)
                        {
                            // Check if view shared for user in 365
                            var view365 = userInitialViews365.Where(v => v["name"].ToString() == userView2016["name"].ToString() && v.Id == userView2016.Id).FirstOrDefault();

                            // If no
                            if (view365 == null)
                            {
                                // Search 365 Dashboards by owner
                                var owner365 = users365.Where(u => u["new_sourceid"].ToString().ToLower() == ((EntityReference)userView2016["ownerid"]).Id.ToString().ToLower()).FirstOrDefault();

                                var userViews365 = new List<Entity>();
                                if (owner365 != null)
                                {
                                    userViews365 = Hunton.GetUserViews(OrganizationService.Instance.GetService(cred365, owner365.Id));
                                }                                

                                // Seach in Owners Views
                                view365 = userViews365.Where(d => d["name"].ToString() == userView2016["name"].ToString() && d.Id == userView2016.Id).FirstOrDefault();

                                // Search in Dynamics Dashboards
                                if (view365 == null)
                                {
                                    view365 = userDynamicsViews365.Where(d => d["name"].ToString() == userView2016["name"].ToString() && d.Id == userView2016.Id).FirstOrDefault();
                                    if (view365 == null)
                                    {
                                        Console.WriteLine("365::{1}-{2}::V[{0}]::U[{3}]::O[{4}] !!!!!!!!", userView2016["name"].ToString(), cc, ovcc, user2016["fullname"].ToString(), ((EntityReference)userView2016["ownerid"]).Name);

                                        userView2016.Attributes.Remove("owningbusinessunit");
                                        userView2016.Attributes.Remove("owninguser");
                                        userView2016.Attributes.Remove("createdby");
                                        userView2016.Attributes.Remove("createdon");
                                        userView2016.Attributes.Remove("modifiedon");
                                        userView2016.Attributes.Remove("modifiedby");
                                        if (owner365 != null)
                                        {
                                            Console.WriteLine("365::{1}-{2}::[{0}][{3}] - CREATE", userView2016["name"].ToString(), cc, ovcc, owner365["fullname"].ToString());
                                            userView2016["ownerid"] = new EntityReference(owner365.LogicalName, owner365.Id);
                                        }
                                        else
                                        {
                                            // Create UserView for Dynamics
                                            Console.WriteLine("365::{1}-{2}::[{0}][{3}] - CREATE DYNAMICS", userView2016["name"].ToString(), cc, ovcc, "Dynamics #");

                                            userView2016.Attributes.Remove("ownerid");
                                            userView2016.Attributes.Remove("owneridtype");
                                        }

                                        try
                                        {
                                            var newId = Hunton.CreateUserView(OrganizationService.Instance.GetService(cred365, owner365 != null ? owner365.Id : Guid.Empty), userView2016);
                                            Console.WriteLine("365::{1}-{2}::[{0}] - ++++++++++++++", userView2016["name"].ToString(), cc, ovcc);
                                        }
                                        catch (Exception ex)
                                        {
                                            Console.WriteLine("365::{1}-{2}::[{0}] - !!!!!!!!!!!!!!", userView2016["name"].ToString(), cc, ovcc);
                                            WriteLog(string.Format("{1}\r\nERROR: {0}", ex.Message, userView2016["name"].ToString()), "Errors.txt");
                                        }

                                        
                                        //WriteLog(string.Format("V[{0}]::U[{1}]::O[{2}]", userView2016["name"].ToString(), user2016["fullname"].ToString(), ((EntityReference)userView2016["ownerid"]).Name), "MissingViews.txt");
                                    }
                                    else
                                    {
                                        Console.WriteLine("365::{1}-{2}::V[{0}]::U[{3}]::IO[{4}] DS", view365["name"].ToString(), cc, ovcc, user365["fullname"].ToString(), user2016["fullname"].ToString());
                                        //Console.WriteLine("365::{0} - Dynamics Owner - Share as Is", view365["name"].ToString());
                                        try
                                        {
                                            var resp = Hunton_GrantAccess(cred365, user365.Id, new EntityReference(view365.LogicalName, view365.Id), (EntityReference)view365["ownerid"]);
                                            Console.WriteLine("365::{1}-{2}::V[{0}]::U[{3}]::IO[{4}] DS +++", view365["name"].ToString(), cc, ovcc, user365["fullname"].ToString(), user2016["fullname"].ToString());
                                        }
                                        catch(Exception ex)
                                        {
                                            WriteLog(string.Format("365::{1}-{2}::V[{0}]::U[{3}]::IO[{4}] DS\r\n{5}", view365["name"].ToString(), cc, ovcc, user365["fullname"].ToString(), user2016["fullname"].ToString(), ex.Message), "Errors.txt");
                                        }                                        
                                    }
                                }
                                else
                                {
                                    // Share
                                    Console.WriteLine("365::{1}-{2}::V[{0}]::U[{3}]::O[{4}] S", view365["name"].ToString(), cc, ovcc, user365["fullname"].ToString(), ((EntityReference)view365["ownerid"]).Name);
                                    //Console.WriteLine("365::{0} - Share by Owner", view365["name"].ToString());
                                    //    Console.WriteLine("GRANT for - YES DASHBOARD");
                                    try
                                    {
                                        var resp = Hunton_GrantAccess(cred365, user365.Id, new EntityReference(view365.LogicalName, view365.Id), (EntityReference)view365["ownerid"]);
                                        Console.WriteLine("365::{1}-{2}::V[{0}]::U[{3}]::O[{4}] S +++", view365["name"].ToString(), cc, ovcc, user365["fullname"].ToString(), ((EntityReference)view365["ownerid"]).Name);
                                    }
                                    catch(Exception ex)
                                    {
                                        WriteLog(string.Format("365::{1}-{2}::V[{0}]::U[{3}]::O[{4}] S\r\n{5}", view365["name"].ToString(), cc, ovcc, user365["fullname"].ToString(), ((EntityReference)view365["ownerid"]).Name, ex.Message), "Errors.txt");
                                    }
                                }
                            }
                            else
                            {
                                if(user365.Id == ((EntityReference)view365["ownerid"]).Id)
                                {
                                    Console.WriteLine("365::{1}-{2}::V[{0}]::U[{3}]::O[{4}] OW", view365["name"].ToString(), cc, ovcc, user365["fullname"].ToString(), ((EntityReference)view365["ownerid"]).Name);
                                }
                                else
                                {
                                    Console.WriteLine("365::{1}-{2}::V[{0}]::U[{3}]::O[{4}] ALREADY", view365["name"].ToString(), cc, ovcc, user365["fullname"].ToString(), ((EntityReference)view365["ownerid"]).Name);
                                }                               

                                //Console.WriteLine("365::{1}-{2}::V[{0}]::U[{3}]::O[{4}] OW", view365["name"].ToString(), cc, ovcc, user365["fullname"].ToString(), ((EntityReference)view365["ownerid"]).Name);
                                //Console.WriteLine("365::{0} - Already Owner", view365["name"].ToString());
                                try
                                {
                                //    var resp = Hunton_GrantAccess(cred365, user365.Id, new EntityReference(view365.LogicalName, view365.Id), (EntityReference)view365["ownerid"], true);
                                  //  Console.WriteLine("365::{1}-{2}::V[{0}]::U[{3}]::O[{4}] OW +++", view365["name"].ToString(), cc, ovcc, user365["fullname"].ToString(), ((EntityReference)view365["ownerid"]).Name);
                                }
                                catch(Exception ex)
                                {
                                    WriteLog(string.Format("365::{1}-{2}::V[{0}]::U[{3}]::O[{4}] OW\r\n{5}", view365["name"].ToString(), cc, ovcc, user365["fullname"].ToString(), ((EntityReference)view365["ownerid"]).Name, ex.Message), "Errors.txt");
                                }

                            }

                            cc++;
                        }
                    }
                    else
                    {
                        Console.WriteLine("365::NO USER - " + user2016["fullname"].ToString());
                        // DO nothing - user not found
                    }  
                }
                catch (Exception ex)
                {
                    WriteLog(string.Format("USER - {0}\r\nERROR: {1}", user2016["fullname"].ToString(), ex.Message), "Errors.txt");
                }
            }
        }


 

        public static void CheckSystemDashboards(OrganizationCredentials cred2016, OrganizationCredentials cred365)
        {
            var dashboards2016 = Hunton.GetSystemDashboards(OrganizationService.Instance.GetService(cred2016, Guid.Empty));
            var dashboards365 = Hunton.GetSystemDashboards(OrganizationService.Instance.GetService(cred365, Guid.Empty));

            Console.WriteLine("2016::{0}", dashboards2016.Count);
            Console.WriteLine("365::{0}", dashboards365.Count);

            foreach(var dashboard2016 in dashboards2016)
            {
                if(!dashboards365.Any(d => d["name"].ToString() == dashboard2016["name"].ToString()))
                {
                    Console.WriteLine("365::{0} - DOES NOT EXIST!", dashboard2016["name"].ToString());
                }
                else
                {
              //      Console.WriteLine("365::{0} - PRESENT!", dashboard2016["name"].ToString());
                }
            }

        }

        public static void Hunton_CheckUserViews(OrganizationCredentials cred365, List<EntityMeta> entitiesMeta)
        {
            var users365 = Hunton.GetActiveUsersWithSourceId(OrganizationService.Instance.GetService(cred365, Guid.Empty));

            //var timUser = users365.Where(u => u["lastname"].ToString() == "Dwyer").FirstOrDefault();

            var c = 0;
            var overall = users365.Count;
            foreach (var user365 in users365.OrderBy(u => u["fullname"]).ToList())
            {
                c++;
                Console.WriteLine("{0} of {1}", c, overall);
                Console.WriteLine(user365["fullname"].ToString());                

                var userViews365 = Hunton.GetUserViewsByOwnerId(OrganizationService.Instance.GetService(cred365, user365.Id), user365.Id);
                Console.WriteLine("OWNER OF {0}", userViews365.Count);

                if (userViews365.Count == 0) continue;

                //var HTViews = userViews365.Where(v => v["returnedtypecode"].ToString().Contains("new_nsssalestracking")).ToList();

                var vc = 0;
                var vcall = userViews365.Count;

                foreach (var view365 in userViews365)
                {
                    Console.WriteLine("{0}-{1}::[{2}][{3}]", vc, vcall, view365["name"].ToString(), view365["returnedtypecode"].ToString());

                    if(entitiesMeta.Any(e => e.name == view365["returnedtypecode"].ToString()))
                    {
                        var meta = entitiesMeta.Where(e => e.name == view365["returnedtypecode"].ToString()).FirstOrDefault();

                        if (view365["layoutxml"].ToString().Contains(meta.oldValue))
                        {
                            view365["layoutxml"] = view365["layoutxml"].ToString().Replace(meta.oldValue, meta.newValue);
                            OrganizationService.Instance.GetService(cred365, user365.Id).Update(view365);
                            Console.WriteLine(".........UPDATED!");
                        }                      
                        else
                        {
                            Console.WriteLine("...........OK!");
                        }
                    }
                    else
                    {
                        Console.WriteLine("...........SKIPPED!");
                    }
                }
            }           
        }

        public static void Hunton_CheckUsers(OrganizationCredentials cred2016, OrganizationCredentials cred365)
        {
            var users365 = Hunton.GetActiveUsersWithSourceId(OrganizationService.Instance.GetService(cred365, Guid.Empty));
            var users2016 = Hunton.GetActiveUsers(OrganizationService.Instance.GetService(cred2016, Guid.Empty));

            var counter = 0;
            var overall = users2016.Count;

            foreach (var user2016 in users2016.OrderBy(e => e["fullname"]).ToList())
            {
                counter++;
                Console.WriteLine();
                Console.WriteLine("{0} of {1}", counter, overall);
                if(!users365.Any(u => u["new_sourceid"].ToString().ToLower() == user2016.Id.ToString().ToLower()))
                {
                    Console.WriteLine("[{0}][{1}]",user2016["fullname"].ToString(), user2016.Id);
                }
                    
            }
        }


        #endregion

        #region TRANSWESTERN

        public static int GetTotalEntitiesCountForUsers(IOrganizationService service, List<Entity> users, string logicalName)
        {
            var totalCount = 0;
            var number = 0;
            Console.Clear();
            
            foreach (var owner in users)
            {
                number++;
                var records = Transwestern.GetEntitiesByOwner(service, logicalName, owner.Id);
                totalCount += records.Count;
                Console.WriteLine("{0}[{1} of {2}]::USER::[{4}] - {3}", logicalName.ToUpper(), number, users.Count, records.Count, owner["fullname"].ToString());
            }

            return totalCount;
        }

        public static void GetTotalEntitiesCount(IOrganizationService service)
        {
            Console.WriteLine("CONTACTS");
            var contacts = Transwestern.GetEntities(service, "contact");

            Console.Clear();
            Console.WriteLine("COMPANIES");
            var accounts = Transwestern.GetEntities(service, "account");

            Console.Clear();
            Console.WriteLine("TASKS");
            var tasks = Transwestern.GetEntities(service, "task");

            Console.Clear();
            Console.WriteLine("NOTES");
            var notes = Transwestern.GetEntities(service, "annotation");


            Console.WriteLine("CONTACTS: " + contacts.Entities.Count);
            Console.WriteLine("ACCOUNTS: " + accounts.Entities.Count);
            Console.WriteLine("TASKS: " + tasks.Entities.Count);
            Console.WriteLine("NOTES: " + notes.Entities.Count);
        }

        #endregion

        public static void GetTimeForDate(OrganizationCredentials dynamics2016, DateTime date)
        {
            var timeEntries = DynamicsLabs.GetTimeEntries(OrganizationService.Instance.GetService(dynamics2016, Guid.Empty), date);

            var minutes = timeEntries.Where(t => t.Contains("new_billableminutes")).Sum(e => (decimal)e["new_billableminutes"]);

            Console.WriteLine(minutes / 60);
        }

        public static void WriteLog(string msg, string filename)
        {
            //Console.WriteLine(msg);
            using (StreamWriter sw = File.AppendText(filename))
            {
                sw.WriteLine(msg);
            }
        }



        private static bool IsAscendentSortedList(List<int> list)
        {
            if (list.SequenceEqual(list.OrderBy(x => x))) return true;
            return false;
        }

        private static int GetCountofWord(string inputString, string word)
        {
            int counter = 0;
            while (true)
            {
                int start = inputString.IndexOf(word);
                if (start >= 0)
                {
                    inputString = inputString.Substring(start + word.Length);
                    counter++;
                }
                else
                {
                    break;
                }
            }

            return counter;
        }

        public static string RemoveSpecialCharacters(string inputString, string regex)
        {
            StringBuilder sb = new StringBuilder();

            foreach (var _char in inputString)
            {
                if (!regex.Contains(_char))
                {
                    sb.Append(_char);
                }

            }
            return sb.ToString();
        }


        private static List<string> _ListOfAllFiles = new List<string>();

        private static List<string> ListOfAllFiles(string path)
        {
            ListOfAllFilesRecurs(path);
            return _ListOfAllFiles;
        }


        private static void ListOfAllFilesRecurs(string path)
        {

            foreach (var item in Directory.GetDirectories(path))
            {
                foreach (var item2 in Directory.GetFiles(item))
                {
                    _ListOfAllFiles.Add(item2);
                }
                ListOfAllFiles(item);
            }
        }


        private static int GetCountOfWord2Task(string word, string path)
        {
            int counter = 0;
            foreach (var item in _ListOfAllFiles)
            {
                var file = File.ReadLines(item);

                foreach (var item2 in file)
                {
                    counter += GetCountofWord(item2, word);
                }
            }
            return counter;
        }

        static void Main(string[] args)
        {
            //var huntonCred2016 = new OrganizationCredentials("https://crm.huntongroup.com", "HUNTON\\cteam", "Trane1");
            var huntonCred365 = new OrganizationCredentials("https://huntongroup.api.crm.dynamics.com", "dynamics@huntongroup.com", "!CrmAdmin123");
            var huntonSandCred365 = new OrganizationCredentials("https://huntonsandbox.api.crm.dynamics.com", "dynamics@huntongroup.com", "CrmAdmin123!");
            //var dynamics2016 = new OrganizationCredentials("http://crm.dynamicalabs.com:2016/DynamicaLabs", "dnl\\d.bespalov", "1234@Project");
            //var transwestern2016 = new OrganizationCredentials("https://xrm.transwestern.net", "transwestern\\ixs20", "TTrraannss@@123");
            //var transwesternUAT2016 = new OrganizationCredentials("https://xrmuat.transwestern.net", "transwestern\\ixs20", "IXS!7ervice");
            var transwesternOnline = new OrganizationCredentials("https://transwestern.api.crm.dynamics.com", "igor.sarov@transwestern.com", "IXS!8ervice");
            var transwesternTEST2016 = new OrganizationCredentials("https://houtxtcrmfe01.transwestern.net/XRMTest", "transwestern\\ixs20", "TTrraannss@@123");
            var dnlDev365 = new OrganizationCredentials("https://dnldev9.crm4.dynamics.com", "denis.bespalov@dynamicalabs.com", "Challenger#86");
            var dnlProd365 = new OrganizationCredentials("https://dynamicalabs.crm4.dynamics.com", "denis.bespalov@dynamicalabs.com", "Challenger#86");
            var dnlTag365 = new OrganizationCredentials("https://dnltaggingdev.crm4.dynamics.com/", "denis.bespalov@dynamicalabs.com", "Challenger#86");
            //var sharepoint365 = new OrganizationCredentials("https://nibus.crm.dynamics.com", "admin@nibus.onmicrosoft.com", "5zruc19F");
            //    var maxSandbox365 = new OrganizationCredentials("https://mvicar-dev.crm4.dynamics.com", "denis.bespalov@dynamicalabs.com", "Foqo3363");
            //  var pupok365 = new OrganizationCredentials("https://pupok.crm4.dynamics.com", "strike_integration@pupok.onmicrosoft.com", "Alfa1234040");
            // var olesia365 = new OrganizationCredentials("https://galyt.crm.dynamics.com", "sptester@galyt.onmicrosoft.com", "Olesya2017");
            //var electroSonics2016 = new OrganizationCredentials("https://electrosonic.crm.dynamics.com", "Igor.Sarov@electrosonic.com", "Dynamics2018");
            var northropProd365 = new OrganizationCredentials("https://nandj.crm.dynamics.com", "APIService@NorthropandJohnson.com", "1234@Yachtapi");
            var northropSand365 = new OrganizationCredentials("https://njsandbox.crm.dynamics.com", "forceworks.support@NorthropandJohnson.com", "$Fw2016%");
            //var centrixSand365 = new OrganizationCredentials("https://citsandbox.crm.dynamics.com", "forceworks.support@centricsit.com", "$Fw20182%");
            //var aprDev365 = new OrganizationCredentials("https://aprdev.crm.dynamics.com", "forceworks.support@aprenergy.com", "$Fw2018%");            
            //var birchEquip365 = new OrganizationCredentials("https://birch.api.crm.dynamics.com", "forceworks.support@birchequipment.com", "$Fw2016%");
            var petsDev365 = new OrganizationCredentials("https://cfphdev.crm.dynamics.com", "forceworks.support@CompassionFirstPets.com", "DnlFw@2019!");
            var techData365 = new OrganizationCredentials("https://techdata.crm.dynamics.com", "svc_CloudDynCRM@techdata.com", "Dynamics$");
            //var empPetrolium = new OrganizationCredentials("https://empiresandbox.crm.dynamics.com/", "forceworks.support@empirepetroleum.com", "$Fw2016%");
            var empPetroliumSand = new OrganizationCredentials("https://empiresandbox.crm.dynamics.com/", "forceworks.support@empirepetroleum.com", "$Fw2016%");
            var empPetroliumUAT = new OrganizationCredentials("https://empireuat.crm.dynamics.com/", "CRMAdmin@empirepetroleum.com", "967uTsdCcUzUNsRA");
            var empPertoliumProd = new OrganizationCredentials("https://empirepetroleumpartners.crm.dynamics.com/", "CRMAdmin@empirepetroleum.com", "967uTsdCcUzUNsRA");
            var empPertoliumProdCopy = new OrganizationCredentials("https://empireprodrestore.crm.dynamics.com/", "CRMAdmin@empirepetroleum.com", "967uTsdCcUzUNsRA");


            //var dnl1 = OrganizationService.Instance.GetService(dnlProd365, Guid.Empty);

            /*var tdClient = TechData.CreateCrmConnection("svc_CloudDynCRM@techdata.com", "Dynamics$");
            var tService = TechData.GetService("svc_CloudDynCRM@techdata.com", "Dynamics$", "https://techdata.crm.dynamics.com");
             var td = OrganizationService.Instance.GetService(techData365, Guid.Empty);
             var query = new QueryExpression("account");
             query.ColumnSet = new ColumnSet(false);
             var ttt = tService.RetrieveMultiple(query);*/

            /*var contacts = NorthJohnson.CrmRequest(
                     System.Net.Http.HttpMethod.Get,
                     "https://nj2020sandbox.api.crm.dynamics.com/api/data/v9.1/contacts")
                         .Result.Content.ReadAsStringAsync();*/

            //var body = NorthJohnson.CrmRequest(System.Net.Http.HttpMethod.Get, "contacts");
            //NorthJohnson.ConnectTo();
           // var njSandServide = NorthJohnson.GetAccessToken();
           // NorthJohnson.CompareEmails(njSandServide);
            var njService = OrganizationService.Instance.GetService(northropProd365, Guid.Empty);

            NorthJohnson.ProcessActivityParty(njService);
            NorthJohnson.DeleteEmails(njService);

            Console.ReadKey();

            //var empService = OrganizationService.Instance.GetService(empPetroliumSand, Guid.Empty);

                //< solutioncomponentid >{ 60C19858 - BED8 - EA11 - A813 - 000D3A31EE79}</ solutioncomponentid >
           
               //< objectid >{ F1C9B38F - F2D1 - 4778 - 94C1 - C786698AF9C5}</ objectid >
            //EmpirePetrolium.RemoveComponentFromSolution(empService, new Guid("f1c9b38f-f2d1-4778-94c1-c786698af9c5"), 2, "FWEmpirePetroleumUCIv1");             
            //var spService = new SPService("https://empirepetroleumpartners.sharepoint.com/sites/CRM/", "CRMAdmin@empirepetroleum.com", "967uTsdCcUzUNsRA");

            //var listItems = spService.GetLIFullName("Opportunity");
            //var listItems2 = spService.GetListItems("Opportunity");

            //var t = spService.CreateFolder("_1111111");

            /*
            EmpirePetrolium.GetAndDisableDocumentsLocations(empService);

            Console.ReadKey();

            // Get FOLDERS from Sharepoint            
            //var folders = spService.GetFolderItems("Opportunity");
            var resultJson = JsonValue.Parse(folders.JsonRestData);
            var foldersArray = (resultJson["d"] as JsonObject)["results"] as JsonArray;

            // Transform Opportunities
            var processedSharepointFolders = foldersArray.Select(e => new EmpirePetrolium.SharepointFolder{
                Name = e["Name"],
                RelUrl = e["ServerRelativeUrl"],
                ItemsCount = e["ItemCount"]
            }).ToList();

            Console.WriteLine("SHAREPOINT FOLDERS = " + processedSharepointFolders.Count);
            var startTime = DateTime.Now;

            // Get document locations from CRM 
            var activeOpportunities = EmpirePetrolium.GetCrmEntities(empService, "opportunity", new string[] {"name", "po_pc", "parentaccountid"});
            Console.WriteLine("Active OPPO = " + activeOpportunities.Count);
            var docLocations = EmpirePetrolium.GetDocumentsLocations(empService);
            Console.WriteLine("Doc Locations = " + docLocations.Count);

            var c = 0;
            var opLength = activeOpportunities.Count;

            var shFoldersToCreate = new List<string>();
            var sharepointFilesCount = new Dictionary<string, int>();

            //WriteLog(string.Format("Id;PC#;OldSharepointLocation;FilesCount"), "d:\\1_DisabledSharepointFolder.txt");
            WriteLog(string.Format("Id;PC#;OldSharepointLocation;FilesCount"), "d:\\2_OldSharepointFolder.txt");
            WriteLog(string.Format("Id;PC#;OldSharepointLocation"), "d:\\2_NeedCreateSharepointFolder.txt");
            WriteLog(string.Format("Id;PC#;CurrentSharepointLocation"), "d:\\2_NoSharepointCreateNeed.txt");
            WriteLog(string.Format("Id;PC#;DocLocId;CurrentLocationName;CurrentSharepointLocation"), "d:\\2_RenameDocLocation.txt");
            */
            //.Where(e => e.Contains("po_pc") && e["po_pc"].ToString() == "4665").ToList()
            /*foreach (var opp in activeOpportunities)
            {
                c++;
                Console.WriteLine(string.Format("{0} of {1}", c, opLength));

                if(!opp.Contains("po_pc"))
                {
                    Console.WriteLine("No PC#");
                    WriteLog(string.Format("{0}", opp.Id.ToString()), "d:\\NoPOPC.txt");
                    continue;
                }

                var documentLocations = docLocations.Where(d => ((EntityReference)d["regardingobjectid"]).Id == opp.Id).                    
                    Select(k => new {
                    docId = k.Id,
                    relUrl = k["relativeurl"].ToString(),
                    name = k["name"].ToString(),
                    parentId = k.Contains("parentsiteorlocation") ? ((EntityReference)k["parentsiteorlocation"]).Id : Guid.Empty,
                    popc = opp.Contains("po_pc") ? opp["po_pc"].ToString() : "NONE"
                }).ToList();            

                if(documentLocations.Any(d => d.parentId == Guid.Empty))
                {
                    Console.WriteLine("PARENT EMPTY!");
                }

                // STAGE 1
                // Disable locations with parent - non-standard opportunity folder
                if (documentLocations.Any(d => d.parentId.ToString().ToLower() != new Guid("686F69E6-527F-E511-80E3-3863BB349E38").ToString().ToLower()))
                {
                    var notOpportunityParentLocations = documentLocations.Where(d => d.parentId.ToString().ToLower() != new Guid("686F69E6-527F-E511-80E3-3863BB349E38").ToString().ToLower()).ToList();

                    foreach(var dl in notOpportunityParentLocations)
                    {

                        WriteLog(string.Format("{0}--[{1}]---[{2}]---[{3}]", opp.Id.ToString(), opp["po_pc"].ToString(), dl.parentId, dl.name), "d:\\NonStandardLocation.txt");
                        WriteLog(string.Format("{0}--[{1}----{2}]", dl.popc, dl.name, dl.relUrl), "d:\\DisableLocationLocations.txt");

                        if (processedSharepointFolders.Any(e => e.Name == dl.relUrl))
                        { // Write to log old folder - check files ????? Do nothing if no folder

                            // Check if already have information about that folder
                            if (!sharepointFilesCount.ContainsKey(dl.relUrl))
                            {
                                var filesCount = 0;

                                var result2 = spService.GetItemsFromFolder("Opportunity/" + dl.relUrl);
                                var resultJson2 = JsonValue.Parse(result2.JsonRestData);
                                var filesFromFolderArray = ((resultJson2["d"] as JsonObject)["Files"] as JsonObject)["results"] as JsonArray;
                                var foldersFromFolderArray = ((resultJson2["d"] as JsonObject)["Folders"] as JsonObject)["results"] as JsonArray;

                                filesCount += filesFromFolderArray.Count;

                                foreach (var fld in foldersFromFolderArray)
                                {

                                    var resultFiles1 = spService.GetFilesFromFolder("Opportunity/" + dl.relUrl + "/" + fld["Name"]);
                                    var resultFilesJson1 = JsonValue.Parse(resultFiles1.JsonRestData);
                                    var filesArray1 = (resultFilesJson1["d"] as JsonObject)["results"] as JsonArray;

                                    filesCount += filesArray1.Count;                                    
                                }

                                sharepointFilesCount.Add(dl.relUrl, filesCount);
                            }

                            WriteLog(string.Format("{0};{1};{2};{3}", opp.Id, opp["po_pc"].ToString(), dl.relUrl, sharepointFilesCount[dl.relUrl]), "d:\\2_OldSharepointFolder.txt");
                        }

                        try
                        {
                            // Disable unwanted
                               EmpirePetrolium.DisableLocation(empService, dl.docId); 
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }                        

                        documentLocations.Remove(dl);
                    }
                }*/

            // STAGE 2 
            // Create additional location and sharepoint folder
            /*if (documentLocations.Count == 0)
            {
                try
                {
                    EmpirePetrolium.CreateDocumentLocation(empService, opp.Id, opp["po_pc"].ToString());
                }
                catch(Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }

                if (!processedSharepointFolders.Any(e => e.Name == opp["po_pc"].ToString()))
                { // No Folder - Create sharepoint folder
                  //WriteLog(string.Format("{0}--[{1}]--[{2}]", opp.Id, opp["po_pc"].ToString(),documentLocations[0].relUrl), "d:\\1_NeedCreateSharepointFolder.txt");
                    WriteLog(string.Format("{0};{1};{2}", opp.Id, opp["po_pc"].ToString(), opp["po_pc"].ToString()), "d:\\2_NeedCreateSharepointFolder.txt");

                    if (!shFoldersToCreate.Contains(opp["po_pc"].ToString()))
                        shFoldersToCreate.Add(opp["po_pc"].ToString());
                }

                WriteLog(string.Format("{0}--[{1}]", opp.Id.ToString(), opp["po_pc"].ToString()), "d:\\NoDocLocations.txt");
            }

            // STAGE 3
            // If more than 1 folder - disable it as save number of files
            if (documentLocations.Count > 1)
            {
                // Search for unwanted location
                var defaultDocLocations = documentLocations.Where(l => l.relUrl != l.name).ToList();
                var poPcDocLocations = documentLocations.Where(l => l.relUrl != l.popc).ToList();

                var removeLocationsList = poPcDocLocations.Count == 1 ? poPcDocLocations : defaultDocLocations;


                // Get files count for unwanted location
                foreach (var dl in removeLocationsList)
                {
                    if (processedSharepointFolders.Any(e => e.Name == dl.relUrl))
                    { // Write to log old folder - check files ????? Do nothing if no folder

                        // Check if already have information about that folder
                        if (!sharepointFilesCount.ContainsKey(dl.relUrl))
                        {
                            var filesCount = 0;

                            var result2 = spService.GetItemsFromFolder("Opportunity/" + dl.relUrl);
                            var resultJson2 = JsonValue.Parse(result2.JsonRestData);
                            var filesFromFolderArray = ((resultJson2["d"] as JsonObject)["Files"] as JsonObject)["results"] as JsonArray;
                            var foldersFromFolderArray = ((resultJson2["d"] as JsonObject)["Folders"] as JsonObject)["results"] as JsonArray;

                            filesCount += filesFromFolderArray.Count;

                            foreach (var fld in foldersFromFolderArray)
                            {

                                var resultFiles1 = spService.GetFilesFromFolder("Opportunity/" + dl.relUrl + "/" + fld["Name"]);
                                var resultFilesJson1 = JsonValue.Parse(resultFiles1.JsonRestData);
                                var filesArray1 = (resultFilesJson1["d"] as JsonObject)["results"] as JsonArray;

                                filesCount += filesArray1.Count;                                    
                            }

                            sharepointFilesCount.Add(dl.relUrl, filesCount);
                        }

                        WriteLog(string.Format("{0};{1};{2};{3}", opp.Id, opp["po_pc"].ToString(), dl.relUrl, sharepointFilesCount[dl.relUrl]), "d:\\2_OldSharepointFolder.txt");
                    }


                    WriteLog(string.Format("{0}--[{1}----{2}]", dl.popc, dl.name, dl.relUrl), "d:\\DisableLocationLocations.txt");

                    try
                    {
                        // Disable unwanted
                        EmpirePetrolium.DisableLocation(empService, dl.docId);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }

                    // Remove from list
                    documentLocations.Remove(dl);
                }                    
            }                

            // STAGE 4
            // IF only one root folder - check name and sharepoint folder
            if(documentLocations.Count == 1)
            {
                if(documentLocations[0].relUrl == documentLocations[0].popc)
                {
                    // Check only sharepoint folder and name                        
                    if(!processedSharepointFolders.Any(e => e.Name == documentLocations[0].popc))
                    { // No Folder - Create sharepoint folder
                        //WriteLog(string.Format("{0}--[{1}]--[{2}]", opp.Id, opp["po_pc"].ToString(),documentLocations[0].relUrl), "d:\\1_NeedCreateSharepointFolder.txt");
                        WriteLog(string.Format("{0};{1};{2}", opp.Id, opp["po_pc"].ToString(), documentLocations[0].relUrl), "d:\\2_NeedCreateSharepointFolder.txt");

                        if (!shFoldersToCreate.Contains(documentLocations[0].popc))
                            shFoldersToCreate.Add(documentLocations[0].popc);
                    }
                    // Folder exist - nothing to do

                    // Check if namne fit - fo rename
                    if (documentLocations[0].name != documentLocations[0].popc)
                    {
                        WriteLog(string.Format("{0}--[{1}]--({2})--({3})", opp.Id, opp["po_pc"].ToString(), documentLocations[0].name, documentLocations[0].relUrl), "d:\\1_NameAndPCNotMatch.txt");
                        WriteLog(string.Format("{0};{1};{2};{3};{4}", opp.Id, opp["po_pc"].ToString(), documentLocations[0].docId, documentLocations[0].name, documentLocations[0].relUrl), "d:\\2_RenameDocLocation.txt");

                        try
                        {
                            EmpirePetrolium.UpdateNameAndPath(empService, documentLocations[0].docId, documentLocations[0].popc); 
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }

                    }
                }
                else
                {
                    if (processedSharepointFolders.Any(e => e.Name == documentLocations[0].relUrl))
                    { // Write to log old folder - check files ????? Do nothing if no folder

                        // Check if already have information about that folder
                        if (!sharepointFilesCount.ContainsKey(documentLocations[0].relUrl))
                        {
                            var filesCount = 0;

                            var result2 = spService.GetItemsFromFolder("Opportunity/" + documentLocations[0].relUrl);
                            var resultJson2 = JsonValue.Parse(result2.JsonRestData);
                            var filesFromFolderArray = ((resultJson2["d"] as JsonObject)["Files"] as JsonObject)["results"] as JsonArray;
                            var foldersFromFolderArray = ((resultJson2["d"] as JsonObject)["Folders"] as JsonObject)["results"] as JsonArray;

                            filesCount += filesFromFolderArray.Count;

                            foreach (var fld in foldersFromFolderArray)
                            {

                                var resultFiles1 = spService.GetFilesFromFolder("Opportunity/" + documentLocations[0].relUrl + "/" + fld["Name"]);
                                var resultFilesJson1 = JsonValue.Parse(resultFiles1.JsonRestData);
                                var filesArray1 = (resultFilesJson1["d"] as JsonObject)["results"] as JsonArray;

                                filesCount += filesArray1.Count;

                                //Console.WriteLine(f["Name"].ToString());
                            }

                            sharepointFilesCount.Add(documentLocations[0].relUrl, filesCount);
                        }

                        WriteLog(string.Format("{0};{1};{2};{3}", opp.Id, opp["po_pc"].ToString(), documentLocations[0].relUrl, sharepointFilesCount[documentLocations[0].relUrl]), "d:\\2_OldSharepointFolder.txt");
                    }

                    // Check new folder for PO_PC
                    if (!processedSharepointFolders.Any(e => e.Name == documentLocations[0].popc))
                    {// Need create new Sharepoint folder
                        WriteLog(string.Format("{0};{1};{2}", opp.Id, opp["po_pc"].ToString(), documentLocations[0].relUrl), "d:\\2_NeedCreateSharepointFolder.txt");

                        if (!shFoldersToCreate.Contains(documentLocations[0].popc))
                            shFoldersToCreate.Add(documentLocations[0].popc);
                    }
                    else
                    { // Don not create
                        WriteLog(string.Format("{0};{1};{2}", opp.Id, opp["po_pc"].ToString(), documentLocations[0].relUrl), "d:\\2_NoSharepointCreateNeed.txt");
                        //Console.WriteLine(documentLocations[0].relUrl + " - OK");
                    }

                    // Rename document location to PO_PC
                    try
                    {
                        EmpirePetrolium.UpdateNameAndPath(empService, documentLocations[0].docId, documentLocations[0].popc); 
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    WriteLog(string.Format("{0};{1};{2};{3};{4}", opp.Id, opp["po_pc"].ToString(), documentLocations[0].docId, documentLocations[0].name, documentLocations[0].relUrl), "d:\\2_RenameDocLocation.txt");
                }
            }


            //Console.WriteLine("DOCSCOUNT = " + documentLocations.Count);
        }*/


            // Create sharepoint folders
            /*Console.WriteLine("Create Unique FOLDERS = " + shFoldersToCreate.Count);
            Console.ReadKey();
            var cf = 0;
            var foldersCount = shFoldersToCreate.Count;
            foreach (var shFolder in shFoldersToCreate)
            {
                cf++;
                Console.WriteLine(string.Format("{0} of {1}", cf, foldersCount));
                WriteLog(shFolder, "d:\\CreateSharepoint.txt");
                try
                {
                    spService.CreateFolderForOpportunity(shFolder);
                }
                catch(Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }

            Console.WriteLine("Start Time " + startTime.ToLocalTime());
            Console.WriteLine("END Time " + DateTime.Now.ToLocalTime());*/

            Console.ReadKey();

//            WriteLog(string.Format("{0};{1};{2};{3};{4}", pr["name"].ToString()
            //WriteLog("Name;Mode;PrimartyEntity;Scope", "d:\\opportunities.csv");
            



            /* var td = OrganizationService.Instance.GetService(techData365, Guid.Empty);
             var query = new QueryExpression("account");
             query.ColumnSet = new ColumnSet("false");
             var ttt = td.RetrieveMultiple(query);*/


//var tagDnl = OrganizationService.Instance.GetService(dnlTag365, Guid.Empty);
//var solutions = ContinuosDeployment.GetSolutions(tagDnl);
//ContinuosDeployment.ExportSolution(tagDnl, "DynamicaLabsLicensing");
//ContinuosDeployment.ExportSolution(tagDnl, "DynamicaLabsTagging");
//ContinuosDeployment.ExportSolution(tagDnl, "Default");



//var twOnline = OrganizationService.Instance.GetService(transwesternOnline, Guid.Empty);

//var twUsers = Transwestern.GetEntities(twOnline,"systemusers");

//var eugen = new Guid("BEF70A41-8FD4-E711-80EA-3863BB344C30");
//var alex = new Guid("6AD3297A-9A89-E711-80E4-3863BB357C40");
//var alexey = new Guid("A49479AB-CA5C-E811-A847-000D3A2A924B");
//var nina = new Guid("76D3297A-9A89-E711-80E4-3863BB357C40");
//var lina = new Guid("15582BB1-738D-E711-80E3-3863BB35AFD0");
//var sdk = new Guid("174C7B1C-55FF-E811-A85F-000D3A2A979C");

/*var dnlProd = OrganizationService.Instance.GetService(dnlProd365, Guid.Empty);

var iterations = DynamicsLabs.GetTeamIterations(dnlProd).Where(e => ((DateTime)e["dnl_enddate"]).DayOfWeek == DayOfWeek.Friday).ToList();

foreach(var it in iterations)
{
    var currentDate = (DateTime)it["dnl_enddate"];
    it["dnl_enddate"] = currentDate.AddDays(2);
    dnlProd.Update(it);
}*/


Console.ReadKey();

            //var ps = dnlProd.Retrieve("dnl_projectsprint", new Guid("4b59b12d-7785-ea11-a811-000d3a4b2c9e"), new ColumnSet("dnl_approvedworkitemslist",
            //"dnl_approvedmustworkitemslist",                                                                                           "dnl_approvedshouldworkitemslist"))
            

            //var all = ps["dnl_approvedworkitemslist"].ToString().Replace("[", "").Replace("]", "").Replace("\"", "").Split(',');
            //var must = ps["dnl_approvedmustworkitemslist"].ToString().Replace("[", "").Replace("]", "").Replace("\"", "").Split(',');
            //ring[] should = new string[0];//ps["dnl_approvedshouldworkitemslist"].ToString().Replace("[", "").Replace("]", "").Replace("\"", "").Split(',');

            //          var wiList = DynamicsLabs.GetActiveWorkItems(dnlProd, all);



            //        var closedMustCount = wiList.Count(e => must.Contains(e.Id.ToString()));
            //        var closedShouldCount = wiList.Count(e => should.Contains(e.Id.ToString()));
            //var allShouldCount = wiList.Count(e => all.Contains(e.Id.ToString()));

            //      var totalPercent = (wiList.Count() / all.Length) * 100.00m;

            //    var wi = new Entity("dnl_workitem", new Guid("87988306-c88a-ea11-a811-000d3a38a85b"));
            //  wi["dnl_devestimate"] = (wiList.Count() / all.Length) * 100.00m;

            //dnlProd.Update(wi);

            //var mustPercent = (wiList.Count(e => must.Contains(e.Id.ToString())) / must.Length) * 100;
            //var shouldPercent = (wiList.Count(e => should.Contains(e.Id.ToString())) / should.Length) * 100;

            //var mustAndShouldPercent = ((wiList.Count(e => must.Contains(e.Id.ToString())) + wiList.Count(e => should.Contains(e.Id.ToString()))) * 100) / (must.Length + should.Length);
            //var g = new Guid(ps1[0]);

            Console.ReadKey();

            /*var projects = DynamicsLabs.GetProjects(dnlProd);

            var num = 1;
            foreach(var pr in projects)
            {
                if (num >= 10)
                {
                    pr["dnl_projectnumber"] = "PN-000" + num;
                }
                else
                {
                    pr["dnl_projectnumber"] = "PN-0000" + num;
                }
                dnlProd.Update(pr);
                num++;
            }*/


            //var emailsList = DynamicsLabs.GetNewEmails(dnlProd);
            //var tasksList = DynamicsLabs.GetTodayTasks(dnlProd);
            //try
            //{
            //    var plList = DynamicsLabs.GetActivePaymentLines(dnlProd);
            //    Console.WriteLine("PL - " + plList.Count);
            //}
            //catch { Console.WriteLine("CANT READ PL!"); }

            //try
            //{
            //    var whList = DynamicsLabs.GetWorkHours(dnlProd);
            //    Console.WriteLine("WH - " + whList.Count);                  
            //}
            //catch { Console.WriteLine("CANT READ WH!"); }

            //Console.WriteLine("EMAILS - " + emailsList.Count);
            //Console.WriteLine("TASKS - " + tasksList.Count);



            Console.ReadKey();

            //StripeServices.CreateCustomer("dnlstripe@gmail.com");


            //            var petsSandService = OrganizationService.Instance.GetService(petsDev365, Guid.Empty);

            //          var hosp = petsSandService.Retrieve("account", new Guid("7CF3FB4B-17C7-E811-A847-000D3A33B5CC"), new ColumnSet(false));


            //            FixOpportunityProducts(OrganizationService.Instance.GetService(huntonCred365, Guid.Empty), new Guid("A67DCE0F-52C4-E811-A961-000D3A1D7B43"));

            //DynamicsLabs.ProcessDeliverables(OrganizationService.Instance.GetService(dnlProd365, Guid.Empty));

            //   int x = 4; int b = 2;

            // x -= b /= x * b;



            /*      var deanaAlbano = new Guid("32EE30EB-82E7-E811-A982-000D3A34E641");
                  var alisaBowman = new Guid("898EBF2F-2418-E811-A953-000D3A34EDEB");
                  var jeniJerovski = new Guid("0620D1BE-2318-E811-A957-000D3A34EF1D");

                  var jonPaulson = new Guid("08C331A4-6CA1-E811-A960-000D3A36478D"); //contact

                  var birchService = OrganizationService.Instance.GetService(birchEquip365, Guid.Empty);
                  var petsDev = OrganizationService.Instance.GetService(petsDev365, Guid.Empty);

                  var host = petsDev.Retrieve("account", new Guid("C2038A92-7DB8-E811-A848-000D3A342217"), new ColumnSet("fw_tier"));

                  var allAppointments = BirchEquipments.GetCrmEntities(birchService, "appointment");

                  string fetchApps = @"  
                         <fetch version='1.0' output-format='xml-platform' mapping='logical'>
                         <entity name='appointment'>
                              <attribute name='subject' />
                              <attribute name='statecode' />
                              <attribute name='scheduledstart' />
                              <attribute name='scheduledend' />
                              <attribute name='createdby' />
                              <attribute name='regardingobjectid' />
                              <attribute name='instancetypecode' />
                              <attribute name='activityid' />
                               <attribute name='createdon' />`                 
                          <filter type='and'>                         
                               <condition attribute='createdon' operator='last-seven-days' />
                               <condition attribute='scheduledstart' operator='not-null' />
                               <condition attribute='scheduledend' operator='not-null' />
                              <condition attribute='instancetypecode' operator='ne' value='2' />
                          </filter>
                          <link-entity name='activityparty' from='activityid' to='activityid'>
                            <filter type='and'>
                              <condition attribute='partyid' operator='eq-userid' />
                              <condition attribute='participationtypemask' operator='ne' value='9' />
                            </filter>
                          </link-entity>
                        </entity>
                      </fetch>";

                  var fetchAppointments = BirchEquipments.GetCrmEntiesUsingFetch(birchService, fetchApps);

                  var distinctApps = fetchAppointments.GroupBy(a => a.Id).Select(g => g.First()).ToList();

                  foreach(var aps in fetchAppointments)
                  {
                      if(!allAppointments.Any(a => a.Id == aps.Id))
                      {
                          Console.WriteLine("NO!");
                      }
                  }*/



            //   var allContacts = BirchEquipments.GetCrmEntities(birchService, "contact", new string[] { "ownerid", "fullname" } );

            // var notAliceContacts = allContacts.Where(a => ((EntityReference)a["ownerid"]).Id != jeniJerovski).ToList();

            //var sharedApps = BirchEquipments.GetAccessRightsForContacts(birchService, new Guid("446B9345-13F0-E711-A94D-000D3A34E641"));

            //var inherited = sharedApps.Where(a => a.Contains("RIGHTS.inheritedaccessrightsmask") && (int)((AliasedValue)a["RIGHTS.inheritedaccessrightsmask"]).Value > 0).ToList();

            //var notAliceInherited = inherited.Where(a => ((EntityReference)a["ownerid"]).Id != alisaBowman).ToList();

            //var sharedToUsers = BirchEquipments.GetPrincipalObjectAccess(birchService);


            //var groupedBy = sharedToUsers.GroupBy(a => ((AliasedValue)a["ct.fullname"]).Value);
            //Console.WriteLine(groupedBy.Count());

            //foreach(var inh in notAliceInherited)
            {
                //  var poa = new Entity("principalobjectaccess", (Guid)((AliasedValue)inh["RIGHTS.principalobjectaccessid"]).Value);

                //                poa["inheritedaccessrightsmask"] = 0;

                //   birchService.Delete(poa.LogicalName, poa.Id);
            }

            //foreach (var cont in notAliceContacts)
            //{

            //}
            //var sharedApps = BirchEquipments.GetAccessRightsForAppoitments(birchService, notOwnedAppointments.FirstOrDefault().Id, deanaAlbano);
            Console.ReadKey();

            /*var allAppointments = BirchEquipments.GetCrmEntities(birchService, "appointment");
            

            var jonPaulsonApps = allAppointments.Where(a => ((EntityReference)a["regardingobjectid"]).Id == jonPaulson).ToList();

            Console.WriteLine("Found Appointments - {0}\n Jon Paul Appointments - {1}", allAppointments.Count, jonPaulsonApps.Count);

            var notOwnedAppointments = jonPaulsonApps.Where(a => ((EntityReference)a["ownerid"]).Id != deanaAlbano).ToList();

            Console.WriteLine("Found Appointments - {0}\nNot Owned Appointments - {1}", allAppointments.Count, notOwnedAppointments.Count);*/


            /*RetrievePrincipalAccessRequest principalAccessRequest = new RetrievePrincipalAccessRequest
            {
                Principal = new EntityReference("systemuser", deanaAlbano),
                Target = notOwnedAppointments.FirstOrDefault().ToEntityReference()
            };
            RetrievePrincipalAccessResponse principalAccessResponse = (RetrievePrincipalAccessResponse)birchService.Execute(principalAccessRequest);

            Console.ReadKey();*/

            //  var sharedApps = BirchEquipments.GetAccessRightsForAppoitments(birchService, notOwnedAppointments.FirstOrDefault().Id, deanaAlbano);

            //Guid gUserId = ((WhoAmIResponse)crmService.Execute(new WhoAmIRequest())).UserId;



            Console.ReadKey();

            //principalAccessResponse.AccessRights; // Is Access user to account

            /*try
            {
                var centService = OrganizationService.Instance.GetService(centrixSand365, Guid.Empty);

                var oppo = new Entity("opportunity", new Guid("8D5953C7-69B6-E811-A96D-000D3A109280"));
                oppo["stageid"] = new Guid("26127c92-bce1-47e2-8175-136479b64509"); // Develop
                centService.Update(oppo);
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }*/





            //FixOpportunityProducts(OrganizationService.Instance.GetService(huntonCred365, Guid.Empty), new Guid("4EA89DA3-BC20-E811-8121-C4346BAC3FA4"));



                                 
            //CreateEntityCountDoc(OrganizationService.Instance.GetService(northropProd365, Guid.Empty), entResp);

            /*            var query = new QueryExpression("workflow");
                        query.ColumnSet = new ColumnSet("ownerid", "owninguser", "mode", "name", "scope", "primaryentity", "runas");
                        query.Criteria.AddCondition("ownerid", ConditionOperator.Equal, new Guid("23d47ffe-af8e-e511-80e3-3863bb354d90"));

                        var proccesses = OrganizationService.Instance.GetService(electroSonics2016, Guid.Empty).RetrieveMultiple(query).Entities.ToList();

                        WriteLog("Name;Mode;PrimartyEntity;Scope", "d:\\opportunities.csv");
                        foreach (var pr in proccesses)
                        {
                            WriteLog(string.Format("{0};{1};{2};{3};{4}", pr["name"].ToString(), 
                                pr.FormattedValues.Contains("primaryentity") ? pr.FormattedValues["primaryentity"].ToString() : "N/A", 
                                pr.FormattedValues["mode"].ToString(), 
                                pr.FormattedValues["scope"].ToString(),
                                pr.FormattedValues.Contains("runas") ? pr.FormattedValues["runas"].ToString() : "N/A"), "d:\\opportunities.csv");
                        }

                        var realTime = proccesses.Where(e => ((OptionSetValue)e["mode"]).Value == 1).ToList();
                        var nonRealTime = proccesses.Where(e => ((OptionSetValue)e["mode"]).Value != 1).ToList();*/
            //proccesses.Where()



            //Console.WriteLine(proccesses.Count);

            /*var query = new QueryExpression("savedquery");
            query.ColumnSet = new ColumnSet(true);// "name", "ownerid");           
            query.Criteria.AddCondition("querytype", ConditionOperator.Equal, 0);
            query.Criteria.AddCondition("returnedtypecode", ConditionOperator.Equal, "account");

            var systemViews = OrganizationService.Instance.GetService(nartemSandbox365, Guid.Empty).RetrieveMultiple(query).Entities.ToList();*/

            /* var query = new QueryExpression("contact");
             query.ColumnSet = new ColumnSet("fullname");
             //query.ad

             LinkEntity link = query.AddLink("incident", "contactid", "customerid", JoinOperator.Inner);
             link.Columns.AddColumn("new_lastvisitdate");
             link.EntityAlias = "caseX";

             link.LinkCriteria = new FilterExpression();
             link.LinkCriteria.AddCondition("incident", "new_lastvisitdate", ConditionOperator.LastMonth);
             query.Distinct = true;            */


            /*LinkEntity link2 = query.AddLink("incident", "contactid", "customerid", JoinOperator.Inner);
            link2.Columns.AddColumn("new_lastvisitdate");
            link2.EntityAlias = "case2X";

            query.Criteria = new FilterExpression();
            query.Criteria.AddCondition("case2X", "new_lastvisitdate", ConditionOperator.ThisMonth);*/
            //query.Criteria.AddCondition("caseX", "new_lastvisitdate", ConditionOperator.OnOrBefore, DateTime.Now);

            //var contacts = OrganizationService.Instance.GetService(nartemSandbox365, Guid.Empty).RetrieveMultiple(query).Entities.ToList();

            //var systemViews = OrganizationService.Instance.GetService(nartemSandbox365, Guid.Empty).RetrieveMultiple(query).Entities.ToList();

            //query.Criteria = new FilterExpression();
            //qx.Criteria.AddCondition("tsk", "activityid", ConditionOperator.Null);



            //  var res = RightMoveTest.GetWebResources(OrganizationService.Instance.GetService(nartemSandbox365, Guid.Empty));


            // GET APR OPPORTUNITY

            // 112
            //var opp1 = OrganizationService.Instance.GetService(aprDev365, Guid.Empty).Retrieve("opportunity", new Guid("015C640E-9C32-E811-80EF-3863BB2ED198"), new ColumnSet(true));
            // 96
            /*var opp2 = OrganizationService.Instance.GetService(aprDev365, Guid.Empty).Retrieve("opportunity", new Guid("13353077-9C32-E811-80EE-3863BB2E0548"), new ColumnSet(true));

            foreach(var att in opp2.Attributes.Keys)
            {
                /*if(att == "originatingleadid")
                {
                    Console.WriteLine(((EntityReference)opp1["originatingleadid"]).Id);
                }*/

              /*  if (opp1.Contains(att) && !att.Contains("apr_"))
                {
                    WriteLog($"[CONTAINS] {att}", "opportunitycompare.txt");
                }*/
                /*else
                {
                    WriteLog($"[NOT] {att}", "opportunitycompare.txt");
                }*/
            //}



            // CREATE PSA TASK
            /*var newTask = new Entity("msdyn_projecttask");

            newTask["msdyn_subject"] = "DEFAULT TASK";
            newTask["msdyn_project"] = new EntityReference("msdyn_project", new Guid("766F1A37-919D-E711-8138-E0071B6502A1"));
            newTask["msdyn_scheduledstart"] = DateTime.UtcNow;
            newTask["msdyn_scheduledend"] = DateTime.UtcNow;
            newTask["msdyn_wbsid"] = "6";

            OrganizationService.Instance.GetService(dnlDev365, Guid.Empty).Create(newTask);*/


            /*var savedQuery = OrganizationService.Instance.GetService(pupok365, Guid.Empty).Retrieve("savedquery", new Guid("4337019D-6BE1-E711-A828-000D3A27889D"), new ColumnSet(true));
            var savedQuery2 = OrganizationService.Instance.GetService(pupok365, Guid.Empty).Retrieve("savedquery", new Guid("4C97E9AA-D0E0-E711-A827-000D3A2B2B9F"), new ColumnSet(true));

            //savedQuery["fetchxml"] = savedQuery2["fetchxml"];
            savedQuery["layoutxml"] = savedQuery2["layoutxml"];
            OrganizationService.Instance.GetService(pupok365, Guid.Empty).Update(savedQuery);*/
            //TwilloServices.CheckSharepoint();
            //TwilloServices.SendSMSNew();
            //TwilloServices.SendSmsMessage();
            /*var currDate = DateTime.UtcNow;
            var testDate = currDate.AddDays(-1);
            var testSt = StripeServices.DateTimeToUnixTimestamp(testDate);
            var reqString = "https://api.stripe.com/v1/charges";
            reqString += string.Format("?limit=100&created[gte]={0}", StripeServices.DateTimeToUnixTimestamp(currDate.AddDays(-1)));
            var ret = StripeServices.ExecuteStripeGetRequest(reqString);
            //JsonValue resJson = JsonValue.Parse(ret);
            //var t = resJson["data"] as JsonObject;
            var stripePayments = resJson["data"] as JsonArray;

            foreach(var sPayment in stripePayments)
            {
              //  Console.WriteLine("Payment ID {0} [{1}][{2}]", sPayment["id"], sPayment["customer"], sPayment["amount"]);
            }
            //var crmPayments = StripeServices.GetCrmPayments(OrganizationService.Instance.GetService(maxSandbox365, Guid.Empty), testDate);
                   */

            //   var documentsList = SharepointServices.GetDocuments(OrganizationService.Instance.GetService(sharepoint365, Guid.Empty));
            //var documentsLocations = SharepointServices.GetDocumentsLocations(OrganizationService.Instance.GetService(olesia365, Guid.Empty));
            //var documentsSites = SharepointServices.GetSites(OrganizationService.Instance.GetService(olesia365, Guid.Empty));
          //  var documentsLocations = SharepointServices.GetDocumentsLocations(OrganizationService.Instance.GetService(sharepoint365, Guid.Empty));
//            var documentsSites = SharepointServices.GetSites(OrganizationService.Instance.GetService(sharepoint365, Guid.Empty));


            /*var defaultSites = documentsSites.Where(e => (bool)e["isdefault"] == true).ToList();

            var sharepointUrl = defaultSites.First()["absoluteurl"].ToString();*/

          //foreach (var d in documentsList)
            {
         //     Console.WriteLine("1");
            }

           var testAccountId = "F4079586-E0EB-E711-A94B-000D3A36478D"; /* OLESYA */
           // var testAccountId = "0EF93A0F-FED8-E711-A94A-000D3A308373"; /*NIBUS*/

            /*var accountLocations = documentsLocations.Where(e => e.Contains("regardingobjectid") &&                                                               
                                                                ((EntityReference)e["regardingobjectid"]).Id == Guid.Parse(testAccountId) &&
                                                                e.Contains("relativeurl") && !string.IsNullOrEmpty(e["relativeurl"].ToString())).First();*/


            /*if(accountLocations.Contains("parentsiteorlocation"))
             {
                 var parentLocation = (EntityReference)accountLocations["parentsiteorlocation"];

                 var parentLocation2 = OrganizationService.Instance.GetService(olesia365, Guid.Empty).Retrieve(parentLocation.LogicalName, parentLocation.Id, new ColumnSet(true));

                 Console.WriteLine("{0}\n{1}", parentLocation2["name"], parentLocation2["relativeurl"]);
                 Console.WriteLine();

                 var parentLocation3 = (EntityReference)parentLocation2["parentsiteorlocation"];

                 var parentLocation4 = OrganizationService.Instance.GetService(olesia365, Guid.Empty).Retrieve(parentLocation3.LogicalName, parentLocation3.Id, new ColumnSet(true));

                 Console.WriteLine("{0}\n{1}", parentLocation4["name"], parentLocation4["absoluteurl"]);
            }*/

            // You might want to work on configuring these 3 settings - URL, User Name, Password !!
            //ISharePointService spService = new SPService("https://galyt.sharepoint.com", "olesya@galyt.onmicrosoft.com", "z?8KEAY`Sk@p*7Y6");
           //ISharePointService spService = new SPService(sharepointUrl, "sptester@galyt.onmicrosoft.com", "Olesya2017");
          //  ISharePointService spService = new SPService(sharepointUrl, "admin@nibus.onmicrosoft.com", "5zruc19F");

            // These are the are part of Sharepoint List Item. Think of better way to handle
            // the field changes.. !!
            //Dictionary<string, string> fields = new Dictionary<string, string>();
            //fields.Add("Title", "SOME TEST TITLE");

            // This is the list name what we see in the Sharepoint. 
            // You might want to keep it as configurable value !!
            //var result = spService.CreateListItem("Leads from CRM", fields);
            //var result = spService.GetFilesFromFolder("account/" + accountLocations["relativeurl"].ToString());

            
            //JsonValue resultJson = JsonValue.Parse(result.JsonRestData);

            /*var valuesArray = (resultJson["d"] as JsonObject)["results"] as JsonArray;

            foreach (var file in valuesArray)
            {
                Console.WriteLine("FILE NAME = {0}", file["Name"]);
                Console.WriteLine("URL - {0}", file["ServerRelativeUrl"]);
                Console.WriteLine("SIZE - {0}", file["Length"]);
                Console.WriteLine("CREATED ON - {0}", file["TimeCreated"]);
                var dateValue = DateTime.Parse(file["TimeCreated"]);
                Console.WriteLine("ID - {0}", file["UniqueId"]);
                var fileId = (string)file["UniqueId"];
                
                var tmpList = SharepointServices.GetEntitesByField(OrganizationService.Instance.GetService(sharepoint365, Guid.Empty), 
                                            "dnl_previewentity", "dnl_sourceid", (string)file["UniqueId"]);
                if (tmpList.Count() < 1)
                {
                    Console.WriteLine("NOTHING!");
                }
                else
                {
                    Console.WriteLine("HAVE {0}", tmpList.Count());
                }
            }*/
            //var barray1 = array[0];
            //var barray2 = array[1];


/*            var result2 = spService.GetListItems("account");


            //Success.Set(context, result.Success);
            if (result.Success)
            {
                Console.WriteLine("LIST ITEM ID = " + result.ListItemId);                
            }*/

            //var elena = OrganizationService.Instance.GetService(transwesternTEST2016, Guid.Empty).Retrieve("systemuser", new Guid("EEE29B25-F4B2-E611-80C7-000D3A9181D6"), new ColumnSet(true));

            // Create New empty Bid contractor
            //var newBidContractor = new Entity("new_biddingcontractor");
            // Set reference to cloned opportunity 
            // newBidContractor["new_opportunitynameid"] = new EntityReference("opportunity", myEntity.Id);
            // Set source bid contractor id
            //newBidContractor.Attributes.Add("new_cloneid", "{8d690fd0-b193-483d-b01e-353e556c379a}");

            //service.Create(newBidContractor);

            //var bid = OrganizationService.Instance.GetService(huntonSandCred365, Guid.Empty).Create(newBidContractor);

            //Console.WriteLine("TEST!");

            //GetTotalEntitiesCount(OrganizationService.Instance.GetService(transwestern2016, Guid.Empty));

            /*Console.WriteLine("Contacts: " + coll.Entities.Count);

            var users = Transwestern.GetActiveUsersByBusinessUnit(OrganizationService.Instance.GetService(transwestern2016, Guid.Empty), new Guid("E19BB03D-CA83-E511-AD7B-00155D0A2F17"));
            //var ownerIdsArray = users.Select(e => e.Id).ToArray();            

            var accountsCount = GetTotalEntitiesCountForUsers(OrganizationService.Instance.GetService(transwestern2016, Guid.Empty), users, "account");
            var contactsCount = GetTotalEntitiesCountForUsers(OrganizationService.Instance.GetService(transwestern2016, Guid.Empty), users, "contact"); 
            var tasksCount = GetTotalEntitiesCountForUsers(OrganizationService.Instance.GetService(transwestern2016, Guid.Empty), users, "task"); 
            var notesCount = GetTotalEntitiesCountForUsers(OrganizationService.Instance.GetService(transwestern2016, Guid.Empty), users, "annotation");*/


            /*Console.Clear();
            Console.WriteLine("USERS: {0}", users.Count);
            Console.WriteLine("ACCOUNTS: {0}", accountsCount);
            Console.WriteLine("CONTACTS: {0}", contactsCount);
            Console.WriteLine("TASKS: {0}", tasksCount);
            Console.WriteLine("NOTES: {0}", notesCount);*/

            //var op2016 = OrganizationService.Instance.GetService(huntonCred2016, Guid.Empty).Retrieve("opportunity", new Guid("dc7709ae-063b-e711-80e3-005056ab6787"), new ColumnSet("modifiedon"));

            //var op365 = new Entity("opportunity", new Guid("dc7709ae-063b-e711-80e3-005056ab6787"));
            //op365["new_overridenmodifiedon"] = op2016["modifiedon"];

            //OrganizationService.Instance.GetService(huntonCred365, Guid.Empty).Update(op365);

            // GetTimeForDate(dynamics2016, new DateTime(2017, 7, 3));

            /*var response = Hunton.RetrieveEntities(OrganizationService.Instance.GetService(huntonCred365, Guid.Empty));
            // Get Custom
            var customEntityList365 = response.EntityMetadata.Where(i => i.IsCustomEntity == true && i.LogicalName.Contains("new_")).OrderBy(e => e.LogicalName).ToList();
            
            response = Hunton.RetrieveEntities(OrganizationService.Instance.GetService(huntonCred2016, Guid.Empty));
            // Get Custom
            var customEntityList2016 = response.EntityMetadata.Where(i => i.IsCustomEntity == true && i.LogicalName.Contains("new_")).OrderBy(e => e.LogicalName).ToList();

            var customEntityMeta = new List<EntityMeta>();

            foreach (var custEntity in customEntityList365)
            {
                var val2016 = customEntityList2016.Where(e => e.LogicalName == custEntity.LogicalName).FirstOrDefault();

                customEntityMeta.Add(new EntityMeta(custEntity.LogicalName,
                                                    val2016 != null ? val2016.ObjectTypeCode.Value.ToString() : string.Empty,
                                                    custEntity.ObjectTypeCode.Value.ToString()));               
            }*/



            /*foreach (var custEntity in customEntityMeta)
            {
                Console.WriteLine("{0}::[{1}]::[{2}]", custEntity.name, custEntity.oldValue, custEntity.newValue);
            }*/


            //Hunton_CheckUserViews(huntonCred365, customEntityMeta);


            //Hunton_FixUserViews(huntonCred365);            

            //  var user = OrganizationService.Instance.GetService(huntonCred365, Guid.Empty).Retrieve("systemuser", new Guid("c4934afe-abb5-dd11-b30c-001372628491"), new ColumnSet("fullname", "isdisabled"));
            //Hunton_CheckUsers(huntonCred2016, huntonCred365);
            //Hunton_ShareUsersDasboards(huntonCred2016, huntonCred365);
            //Hunton_DeleteIncorectViews(huntonCred2016, huntonCred365);
            //  Hunton_CreateAndAssignViews(huntonCred2016, huntonCred365);
            //Hunton_FixUserViews(huntonCred365);
            //Hunton_ShareViews(huntonCred2016, huntonCred365);

            //CheckSystemDashboards(huntonCred2016, huntonCred365);

            //Hunton_ReassignDashboards(huntonCred2016, huntonCred365);
            //Hunton_ShareUsersDasboards(huntonCred2016, huntonCred365);
            Console.ReadKey();


            /*var contacts = LegalServices.GetRecords(OrganizationService.Instance.GetService(), "fw_matter", new string[] { "fw_filenumber" });

            var uid = 1060;

            foreach(var contact in contacts.Where(e => !e.Contains("fw_filenumber")).ToList())
            {
                contact["fw_filenumber"] = "C-" + uid;
                uid++;
                Console.WriteLine(contact["fw_filenumber"].ToString());
                OrganizationService.Instance.GetService().Update(contact);
            }*/

            /*  var timeEntries = LegalServices.GetAssociatedRecords(OrganizationService.Instance.GetService(),                                                              
                                                               "fw_timeentry",
                                                               new Guid("84B8CEEF-E019-E711-8110-C4346BAC8AF8"),
                                                               false);

              var invoiceEntries = timeEntries.Where(e => (Guid)((AliasedValue)e["invoice.invoiceid"]).Value == new Guid("84B8CEEF-E019-E711-8110-C4346BAC8AF8")).ToList();

              foreach(var ie in invoiceEntries)
              {
                  // Disable and Remove                
                  timeEntries.Remove(ie);

                  var excludeTimeEntries = timeEntries.Where(t => t.Id == ie.Id).ToList();

                  foreach(var excludeEntriy in excludeTimeEntries)
                  {
                      // Diss
                      var invoiceId = (Guid)((AliasedValue)excludeEntriy["invoice.invoiceid"]).Value;
                  }
              }

              var timeEntries2 = LegalServices.GetAssociatedRecords(OrganizationService.Instance.GetService(),
                                                   "fw_timeentry",
                                                   new Guid("84B8CEEF-E019-E711-8110-C4346BAC8AF8"));

              var activeTimeEntries = timeEntries2.Where(e => e.Contains("invoice.invoiceid") &&
                                                              e.Contains("invoice.statecode") &&
                                                              ((OptionSetValue)((AliasedValue)e["invoice.statecode"]).Value).Value == 0).ToList();

              foreach(var te in timeEntries)
              {
                  var excludeTimeEntries = activeTimeEntries.Where(t => t.Id == te.Id).ToList();
                  // Remove Reference to Expense in Invoice
                  foreach(var teex in excludeTimeEntries)
                  {
                      Console.WriteLine(string.Format("REMOVE - {0} [{1}]", teex.Id, teex.FormattedValues["invoice.statecode"].ToString()));
                  }

                  // Make Expense Inactive
              }



              Console.WriteLine("CONNECTED - " + timeEntries2.Where(e => e.Contains("invoice.invoiceid") &&
                                                                      e.Contains("invoice.statecode") && 
                                                                      ((OptionSetValue)((AliasedValue)e["invoice.statecode"]).Value).Value == 0).ToList().Count);







              foreach (var te in timeEntries2)
              {
                  Console.WriteLine(string.Format("{0} [{1}][{2}]", te.Id, ((AliasedValue)te["invoice.invoiceid"]).Value.ToString(), te.FormattedValues["invoice.statecode"].ToString()));
              }

              var expenseEntries = LegalServices.GetAssociatedRecords(OrganizationService.Instance.GetService(),
                                                                         "fw_expenseentry",
                                                                         new Guid("84B8CEEF-E019-E711-8110-C4346BAC8AF8"),
                                                                         false);
              var expenseEntries2 = LegalServices.GetAssociatedRecords(OrganizationService.Instance.GetService(),
                                                             "fw_expenseentry",
                                                             new Guid("84B8CEEF-E019-E711-8110-C4346BAC8AF8"));
              Console.WriteLine("=====");
              foreach (var te in expenseEntries)
              {
                  Console.WriteLine(te.Id);
              }

              Console.WriteLine("T:{0} E:{1}", timeEntries.Count,
                                                  expenseEntries.Count);*/


            //            var userId = new Guid("8F1109E1-BCE4-E611-8106-C4346BAC3AC4");
            //          var testId = new Guid("8D908ED4-BCE4-E611-8106-C4346BAC3AC4");


            //        Console.WriteLine("ROLE = " + LegalServices.CheckUserHasRole(OrganizationService.Instance.GetService(), testId, "Legal Service Admin"));



            //Console.WriteLine(givenRoles.Entities.Count);
            //var id = new Guid("22381355-2920-E711-8110-C4346BAC8AF8");
            //var id = new Guid("1E381355-2920-E711-8110-C4346BAC8AF8");
            /*   var tr = LegalServices.RetrieveRecord(OrganizationService.Instance.GetService(),
                                                       "fw_transaction",
                                                       id,
                                                       new string[] { "fw_name", "statecode", "statuscode"  });

               LegalServices.UpdateEntityStatus(OrganizationService.Instance.GetService(), tr, 0, (int)LegalServices.TransactionsStatus.Reconciled);

               var timeEntries = LegalServices.GetAssociatedRecords(OrganizationService.Instance.GetService(),
                                                                            new EntityReference("fw_matter", new Guid("6FC660BA-F6F9-E611-8107-C4346BAC0A3C")),
                                                                            "fw_timeentry");





               foreach (var te in timeEntries)
               {
                   Console.WriteLine(te.Id);
               }

                var expenseEntries = LegalServices.GetAssociatedRecords(OrganizationService.Instance.GetService(),
                                                                          new EntityReference("fw_matter", new Guid("6FC660BA-F6F9-E611-8107-C4346BAC0A3C")),
                                                                            "fw_expenseentry");
               Console.WriteLine("=====");
               foreach (var te in expenseEntries)
               {
                   Console.WriteLine(te.Id);
               }

               Console.WriteLine("T:{0} E:{1}",  timeEntries.Count,
                                                   expenseEntries.Count);*/

            Console.ReadKey();


            //       var priceList = LegalServices.GetDefaultPriceList(OrganizationService.Instance.GetService());

            /*     var attrs = new Dictionary<string, object>();
                 attrs.Add("name", "Generated Associated Test Invoice");
                 attrs.Add("pricelevelid", new EntityReference(priceList.LogicalName, priceList.Id));
                 attrs.Add("fw_from", "From Matter TEST");
                 attrs.Add("fw_bankaccountid", new EntityReference("fw_bankaccount", new Guid("142670C0-8CFE-E611-8108-C4346BAC0A3C")));
                 attrs.Add("customerid", new EntityReference("account", new Guid("64F3D3F9-B9FD-E611-8108-C4346BAC1938")));
                 attrs.Add("fw_matter", new EntityReference("fw_matter", new Guid("BBEAE952-E213-E711-810B-C4346BAC0A3C")));

                 var invoiceId = LegalServices.CreateRecord(OrganizationService.Instance.GetService(), "invoice", attrs);*/

            // Associate Time Entries
            /* LegalServices.AssociateRecords(OrganizationService.Instance.GetService(),
                                             new EntityReference("invoice", invoiceId),
                                             "fw_invoice_fw_timeentry",
                                             timeEntries);
             // Associate Expenses
             LegalServices.AssociateRecords(OrganizationService.Instance.GetService(),
                                             new EntityReference("invoice", invoiceId),
                                             "fw_invoice_fw_expenseentry",
                                             expenseEntries);*/



            /*var invoices = LegalServices.GetAssociatedRecords(OrganizationService.Instance.GetService(),
                                            new EntityReference("fw_matter", new Guid("BBEAE952-E213-E711-810B-C4346BAC0A3C")));*/

            Console.ReadKey();

            /*var matterQuery = new QueryExpression("fw_matter");
            matterQuery.ColumnSet = new ColumnSet(true);
            //leadsQuery.Criteria.AddCondition("fw_sourcesystemnumber", ConditionOperator.NotNull);

            var matterCollection = OrganizationService.Instance.GetService().RetrieveMultiple(matterQuery);

            foreach(var matter in matterCollection.Entities)
            {
                var timeEntries = LegalServices.GetAssociatedRecords(OrganizationService.Instance.GetService(), 
                                                                        new EntityReference(matter.LogicalName, matter.Id), 
                                                                        "fw_timeentry");

                var expenseEntries = LegalServices.GetAssociatedRecords(OrganizationService.Instance.GetService(),
                                                                          new EntityReference(matter.LogicalName, matter.Id),
                                                                            "fw_expenseentry");

                Console.WriteLine("{0} - T:{1} E:{2}", matter.Contains("fw_name") ? matter["fw_name"].ToString() : "EMPTY", 
                                                        timeEntries.Entities.Count, 
                                                        expenseEntries.Entities.Count);
                Console.ReadKey();
                Console.Clear();
            }*/

            //Console.WriteLine("TEST - " + matterCollection.Entities.Count);

            Console.ReadKey();
        }

        
    }
}
