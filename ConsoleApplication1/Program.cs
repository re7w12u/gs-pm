using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Workflow;
using Microsoft.SharePoint.Client.WorkflowServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication1
{
    class Program
    {
        static void Main(string[] args)
        {

            StartWFOnPremise();
            //StartWFOnLine();
        }

        private static void StartWFOnLine()
        {
            //Site Details
            string siteUrl = "https://jbes.sharepoint.com/sites/gosport";
            string userName = "julien@jbes.onmicrosoft.com";
            string password = "xAp09sDt!*$";

            //Name of the SharePoint2010 List Workflow
            string workflowName = "Approbation 2010";

            //Name of the List to which the Workflow is Associated
            string targetListName = "Liste des demandes";

            //Guid of the List to which the Workflow is Associated
            Guid targetListGUID = new Guid("cbe00484-055e-4dca-ade1-863ef7b202ba");

            //Guid of the ListItem on which to start the Workflow
            //Guid targetItemGUID = new Guid("");

            using (ClientContext clientContext = new ClientContext(siteUrl))
            {
                SecureString securePassword = new SecureString();

                foreach (char c in password.ToCharArray()) securePassword.AppendChar(c);

                clientContext.Credentials = new SharePointOnlineCredentials(userName, securePassword);

                Web web = clientContext.Web;

                List list = web.Lists.GetByTitle(targetListName);
                ListItem item = list.GetItemById(1);

                clientContext.Load(item);
                clientContext.ExecuteQuery();

                string guid = item["GUID"].ToString();

                //Workflow Services Manager which will handle all the workflow interaction.
                WorkflowServicesManager wfServicesManager = new WorkflowServicesManager(clientContext, web);

                //Will return all Workflow Associations which are running on the SharePoint 2010 Engine
                WorkflowAssociationCollection wfAssociations = list.WorkflowAssociations;

                //Get the required Workflow Association
                WorkflowAssociation wfAssociation = wfAssociations.GetByName(workflowName);

                clientContext.Load(wfAssociation);

                clientContext.ExecuteQuery();

                //Get the instance of the Interop Service which will be used to create an instance of the Workflow
                InteropService workflowInteropService = wfServicesManager.GetWorkflowInteropService();

                var initiationData = new Dictionary<string, object>();

                //Start the Workflow
                ClientResult<Guid> resultGuid = workflowInteropService.StartWorkflow(wfAssociation.Name, new Guid(), targetListGUID, Guid.Parse(guid), initiationData);

                clientContext.ExecuteQuery();
            }
        }

        private static void StartWFOnPremise()
        {
            //Site Details
            string siteUrl = "http://rdits-sp13-dev2/sites/gs/";
            string userName = "julien.bessiere@non.Schneider-electric.com";
            string password = "password";

            //Name of the SharePoint2010 List Workflow
            string workflowName = "Approbation 2010";

            //Name of the List to which the Workflow is Associated
            string targetListName = "cadeaux";

            //Guid of the List to which the Workflow is Associated
            Guid targetListGUID = new Guid("cb1ef47f-6fbf-4119-97c5-057006db6a13");

            //Guid of the ListItem on which to start the Workflow
            //Guid targetItemGUID = new Guid("B10F6CF0-86F2-4F6D-B982-BBF9FB38897E");
            Guid targetItemGUID = new Guid("31D3B801-5C0C-4758-8AE0-3BF23BD6E430");

            using (ClientContext clientContext = new ClientContext(siteUrl))
            {
                //SecureString securePassword = new SecureString();

                //foreach (char c in password.ToCharArray()) securePassword.AppendChar(c);

                //clientContext.Credentials = new SharePointOnlineCredentials(userName, securePassword);

                Web web = clientContext.Web;

                List list = web.Lists.GetByTitle(targetListName);
                ListItem item = list.GetItemById(19);

                clientContext.Load(item);
                clientContext.ExecuteQuery();

                string guid = item["GUID"].ToString();

                //Workflow Services Manager which will handle all the workflow interaction.
                WorkflowServicesManager wfServicesManager = new WorkflowServicesManager(clientContext, web);

                //Will return all Workflow Associations which are running on the SharePoint 2010 Engine
                WorkflowAssociationCollection wfAssociations = web.Lists.GetByTitle(targetListName).WorkflowAssociations;

                //Get the required Workflow Association
                WorkflowAssociation wfAssociation = wfAssociations.GetByName(workflowName);

                clientContext.Load(wfAssociation);

                clientContext.ExecuteQuery();

                //Get the instance of the Interop Service which will be used to create an instance of the Workflow
                InteropService workflowInteropService = wfServicesManager.GetWorkflowInteropService();

                var initiationData = new Dictionary<string, object>();
                initiationData.Add("DisplayName", "Julien Bessiere");
                initiationData.Add("AccountId", "i:0#.f|membership|julien@jbes.onmicrosoft.com");
                initiationData.Add("AccountType", "User");

                //Start the Workflow
                ClientResult<Guid> resultGuid = workflowInteropService.StartWorkflow(wfAssociation.Name, new Guid(), targetListGUID, Guid.Parse(guid), initiationData);

                try
                {
                    clientContext.ExecuteQuery();
                }
                catch (Exception ex)
                {
                    throw;
                }


            }
        }
    }
}
