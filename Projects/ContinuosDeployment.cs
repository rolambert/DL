using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DNLConsole365.Projects
{
    public class ContinuosDeployment
    {
        public static List<Entity> GetSolutions(IOrganizationService service)
        {
            var query = new QueryExpression("solution");
            query.ColumnSet = new ColumnSet("uniquename");
            query.Criteria.AddCondition("isvisible", ConditionOperator.Equal, "true");

            return service.RetrieveMultiple(query).Entities.ToList();
        }

        public static void ExportSolution(IOrganizationService service, string solutionName)
        {
            var dateTimeNow = DateTime.Now;
            Console.WriteLine(string.Format("{0}:{1}:{2}", dateTimeNow.Hour, dateTimeNow.Minute, dateTimeNow.Second));
            var exportSolutionRequest = new ExportSolutionRequest();
            exportSolutionRequest.Managed = false;
            exportSolutionRequest.SolutionName = solutionName;

            try
            {
                var exportSolutionResponse = (ExportSolutionResponse)service.Execute(exportSolutionRequest);

                byte[] exportXml = exportSolutionResponse.ExportSolutionFile;
                string filename = solutionName + ".zip";
                File.WriteAllBytes("d:\\test" + filename, exportXml);

                dateTimeNow = DateTime.Now;
                Console.WriteLine(string.Format("{0}:{1}:{2}", dateTimeNow.Hour, dateTimeNow.Minute, dateTimeNow.Second));
                Console.WriteLine("Solution exported to {0}.", "d:\\test" + filename);
            }
            catch(Exception ex)
            {
                Console.WriteLine("ERROR - " + ex.Message);
            }
        }
    }
}
