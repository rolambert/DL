using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace DNLConsole365.Projects
{
    public class DynamicsLabs
    {
        // Reassign applications
        public static void ProcessDeliverables(IOrganizationService service)
        {
            var count = 0;
            var todayDeliverables = GetDeliverablesCreatedToday(service).Where(d => d.Contains("dnl_targetprocessid")).ToList();
            var tpids = string.Empty;
            foreach (var td in todayDeliverables)
            {
                count++;
                Console.WriteLine("Processing {0} of {1} - TPID [{2}]", count, todayDeliverables.Count(), td["dnl_targetprocessid"]);
                var oldDeliverable = GetCrmEntityByAttributeValue(service, "dnl_deliverable", 
                    "dnl_targetprocessid", 
                    (DateTime)td["createdon"],
                    td["dnl_targetprocessid"],
                    new string[] { "dnl_deliverableid", "dnl_name" , "dnl_syncvsoupdate" }, false);

                if(oldDeliverable == null)
                {
                    Console.WriteLine("No duplicate. Skip {0}...", td["dnl_targetprocessid"]);
                }
                else
                {
                    var oldDelApplications = GetReleaseApplicationsForDeliverable(service, oldDeliverable.Id);

                    if(oldDelApplications.Count == 0)
                    {
                        Console.WriteLine("No Applications. Mark for Deletion {0}...", td["dnl_targetprocessid"]);

                      //  if(oldDeliverable.Contains("dnl_syncvsoupdate") && (bool)oldDeliverable["dnl_syncvsoupdate"] == true)
                        //{
                            //Console.WriteLine("WRONG!");
                        //}
                        //else
                       // {
                        var tasks = GetTasksForDeliverable(service, oldDeliverable.Id);
                        var workItems = GetWorkItemsForDeliverable(service, oldDeliverable.Id);

                        if(tasks.Count() > 0)
                        {
                            Console.WriteLine("HAS TASKS {1} ! > {0}", td["dnl_targetprocessid"], tasks.Count());
                            tpids += "[T]" + td["dnl_targetprocessid"].ToString() + "*";
                        }
                        if(workItems.Count() > 0)
                        {
                            tpids += "[W]" + td["dnl_targetprocessid"].ToString() + "*";
                            Console.WriteLine("HAS WORKITEMS {1}! > {0}", td["dnl_targetprocessid"], workItems.Count());
                        }

                        if(workItems.Count() == 0 && tasks.Count() == 0)
                        {
                            Console.WriteLine("DELETE!");
                            oldDeliverable["dnl_deliverablestatus"] = new OptionSetValue(6);
                            oldDeliverable["dnl_updatefromtargetprocess"] = true;
                            service.Update(oldDeliverable);
                        }                        
                    }
                    else
                    {
                        Console.WriteLine("Process APPLICATIONS!");
                        Console.WriteLine("PROCESS FOR {0}", td["dnl_targetprocessid"]);
                        foreach (var app in oldDelApplications)
                        {
                            app["dnl_deliverableid"] = new EntityReference("dnl_deliverable", td.Id);
                            service.Update(app);
                        }
                    }
                }
            }
            Console.WriteLine(tpids);
            Console.WriteLine("-------!!!COMPLETED!!!--------");
            Console.ReadKey();
        }

        public static List<Entity> GetProjects(IOrganizationService service)
        {
            var query = new QueryExpression("dnl_project");
            query.ColumnSet = new ColumnSet("dnl_targetprocessid", "dnl_projectnumber");
            query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);

            return service.RetrieveMultiple(query).Entities.OrderBy(x => x["dnl_targetprocessid"]).ToList();
        }

        public static List<Entity> GetNewEmails(IOrganizationService service)
        {
            var query = new QueryExpression("email");
            query.ColumnSet = new ColumnSet(true);
            query.Criteria.AddCondition("createdon", ConditionOperator.LastXHours, 1);

            return service.RetrieveMultiple(query).Entities.ToList();
        }

        public static List<Entity> GetTodayTasks(IOrganizationService service)
        {
            var query = new QueryExpression("task");
            query.ColumnSet = new ColumnSet(true);
            query.Criteria.AddCondition("createdon", ConditionOperator.Today);

            return service.RetrieveMultiple(query).Entities.ToList();
        }

        public static List<Entity> GetActivePaymentLines(IOrganizationService service)
        {
            var query = new QueryExpression("dnl_paymentline")
            {
                ColumnSet = new ColumnSet(true),
                Criteria = new FilterExpression
                {
                    Conditions =
                            {
                                new ConditionExpression("statecode", ConditionOperator.Equal, 0),
                                new ConditionExpression("dnl_paymentstatus", ConditionOperator.Equal, 1)
                            }
                }
            };
             
            return service.RetrieveMultiple(query).Entities.ToList();
        }

        public static List<Entity> GetWorkHours(IOrganizationService service)
        {
            var query = new QueryExpression("dnl_workhours")
            {
                ColumnSet = new ColumnSet("dnl_paidamount", "dnl_paidhours", "dnl_totalhours", "dnl_paymentline"),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("statecode", ConditionOperator.Equal, 0)
                    }
                }
            };

            return service.RetrieveMultiple(query).Entities.ToList();
        }

        public static List<Entity> GetTasksForDeliverable(IOrganizationService service, Guid deliverableId)
        {
            var query = new QueryExpression("task");

            query.ColumnSet = new ColumnSet(false);
            query.Criteria.AddCondition("regardingobjectid", ConditionOperator.Equal, deliverableId);

            return service.RetrieveMultiple(query).Entities.ToList();
        }

        public static List<Entity> GetWorkItemsForDeliverable(IOrganizationService service, Guid deliverableId)
        {
            var query = new QueryExpression("dnl_workitem");

            query.ColumnSet = new ColumnSet(false);
            query.Criteria.AddCondition("dnl_deliverable", ConditionOperator.Equal, deliverableId);

            return service.RetrieveMultiple(query).Entities.ToList();
        }

        public static List<Entity> GetReleaseApplicationsForDeliverable(IOrganizationService service, Guid deliverableId)
        {
            var query = new QueryExpression("dnl_releaseapplication");

            query.ColumnSet = new ColumnSet("dnl_applicationid", "dnl_releaseapplicationid", "dnl_workitem", "statecode");
            query.Criteria.AddCondition("dnl_deliverableid", ConditionOperator.Equal, deliverableId);

            return service.RetrieveMultiple(query).Entities.ToList();
        }

        public static List<Entity> GetDeliverablesCreatedToday(IOrganizationService service)
        {
            var query = new QueryExpression("dnl_deliverable");
            query.ColumnSet = new ColumnSet("dnl_targetprocessid", "dnl_deliverableid", "dnl_name", "createdon");
            query.Criteria.AddCondition("createdon", ConditionOperator.Yesterday);

            return service.RetrieveMultiple(query).Entities.ToList();
        }

        public static Entity GetCrmEntityByAttributeValue(IOrganizationService service, string entityName, string targetAttributeName, DateTime dateParam, object targetAttributeValue = null, string[] cols = null, bool isActive = true)
        {
            var query = new QueryExpression(entityName);

            query.ColumnSet = cols == null ? new ColumnSet(false) : new ColumnSet(cols);

            query.Criteria.AddCondition("createdon", ConditionOperator.LessThan, dateParam);

            if (targetAttributeValue != null)
            {
                query.Criteria.AddCondition(targetAttributeName, ConditionOperator.Equal, targetAttributeValue);
            }
            if (isActive)
            {
                query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); // By Default select only active records
            }
            return service.RetrieveMultiple(query).Entities.FirstOrDefault();
        }

        public static List<Entity> GetTimeEntries(IOrganizationService service, DateTime date)
        {
            var query = new QueryExpression("new_timeentry");
            query.ColumnSet = new ColumnSet(true);
            query.Criteria.AddCondition("ownerid", ConditionOperator.Equal, "4F7452C6-1A87-E511-9412-00155D253707");
            query.Criteria.AddCondition("actualstart", ConditionOperator.OnOrAfter, date);

            return service.RetrieveMultiple(query).Entities.ToList();
        }

        
        public static List<Entity> GetActiveWorkItems(IOrganizationService service, string[] ids)
        {
            var query = new QueryExpression("dnl_workitem");
            query.ColumnSet = new ColumnSet(false);

            query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
            query.Criteria.AddCondition("dnl_status", ConditionOperator.Equal, 13); // Live & Tested
            query.Criteria.AddCondition("dnl_workitemid", ConditionOperator.In, ids);

            return service.RetrieveMultiple(query).Entities.ToList();
        }

        public static List<Entity> GetTeamIterations(IOrganizationService service)
        {
            var query = new QueryExpression("dnl_teamiteration");
            query.ColumnSet = new ColumnSet("dnl_enddate", "dnl_startdate","dnl_name");

            query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
            query.Criteria.AddCondition("dnl_startdate", ConditionOperator.OnOrAfter, new DateTime(2020, 5, 1));
            
            return service.RetrieveMultiple(query).Entities.ToList();
        }

        public static void Test()
        {
            var request = new AssociateEntitiesRequest
            {
                Moniker2 = null,
                Moniker1 = null,
                RelationshipName = "11",
                Parameters = new ParameterCollection()
                {
                    
                }
            };
        }
    }
}
