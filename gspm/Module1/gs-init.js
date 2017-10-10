

// ensure library loading 
SP.SOD.executeOrDelayUntilScriptLoaded(function () {
    SP.SOD.registerSod('sp.workflowservices.js', SP.Utilities.Utility.getLayoutsPageUrl("sp.workflowservices.js"));
    SP.SOD.executeFunc('sp.workflowservices.js', "SP.WorkflowServices.WorkflowServicesManager", null);
}, "SP.js");



// Go Sport Present Management System
var GSPMS = GSPMS || {};

GSPMS.web = null;
GSPMS.context = null;
GSPMS.itemId = null;
GSPMS.listId = null;
GSPMS.list = null;
GSPMS.item = null;
GSPMS.workflow = null;
GSPMS.workflowName = "Approval WF 2010";
GSPMS.wfDefinitionId = "{d6fed34a-2214-47c0-a0b6-8779e8eb2379}";

GSPMS.setItemAsReadOnly = function () {

};


GSPMS.startWorkflow2010 = function () {
    GSPMS.web = GSPMS.context.get_web();
    list = GSPMS.web.get_lists().getById(GSPMS.listId);
    item = list.getItemById(GSPMS.itemId);
    GSPMS.context.load(item);

    GSPMS.context.executeQueryAsync(function () {
        GSPMS.getManagerInfo(item);               
    },
    function (sender, args) {
        console.error("ERROR 1: " + args.get_message());
    });    
};

GSPMS.getManagerInfo = function (item) {
    // get selected manager info
    var managerName = item.get_item("Manager").get_lookupValue();
    var managerId = item.get_item("Manager").get_lookupId();
    var manager = GSPMS.web.getUserById(managerId);

    GSPMS.context.load(manager);
    GSPMS.context.executeQueryAsync(function () {
        var login = manager.get_loginName();
        var email = manager.get_email();
        // get parameters
        var xml = GSPMS.getAssocData(managerName, managerId, login);
        // actual WF trigger
        GSPMS.triggerWF(item, xml);

    }, function (sender, args) {
        console.error("ERROR 3: " + args.get_message());
    });
}

GSPMS.triggerWF = function (item, xml) {

    //Workflow Services Manager
    var wfServicesManager = new SP.WorkflowServices.WorkflowServicesManager(GSPMS.context, GSPMS.web);

    //Workflow Interop Service used to interact with SharePoint 2010 Engine Workflows
    var interopService = wfServicesManager.getWorkflowInteropService()
    itemUniqueId = item.get_item("UniqueId").toString();
    itemGuid = item.get_item("GUID").toString();
    //Start the Site Workflow by Passing the name of the Workflow and the initiation Parameters.
    interopService.startWorkflow(GSPMS.workflowName, null, GSPMS.listId, itemGuid, xml);

    GSPMS.context.executeQueryAsync(function () {
        console.log("workflow started");
        SP.UI.Notify.addNotification('Your element has been submitted to your manager.', false);
    }, function (sender, args) {
        console.error("ERROR 2: " + args.get_message());
    });
}

GSPMS.startWorkflow = function () {
    GSPMS.web = GSPMS.context.get_web();
    GSPMS.list = GSPMS.web.get_lists().getById(GSPMS.listId);    
    GSPMS.item = GSPMS.list.getItemById(GSPMS.item.id);
    GSPMS.workflows = GSPMS.list.get_workflowAssociations();
    GSPMS.servicesManager = SP.WorkflowServices.WorkflowServicesManager.newObject(GSPMS.context, GSPMS.web);
    GSPMS.subs = GSPMS.servicesManager.getWorkflowSubscriptionService().enumerateSubscriptionsByDefinition(GSPMS.wfDefinitionId);
    GSPMS.workflowDefinitions = GSPMS.servicesManager.getWorkflowDeploymentService().enumerateDefinitions(false);

    GSPMS.context.load(GSPMS.workflowDefinitions);            
    GSPMS.context.load(GSPMS.list);
    GSPMS.context.load(GSPMS.item);
    GSPMS.context.load(GSPMS.workflows);
    GSPMS.context.load(GSPMS.servicesManager);
    GSPMS.context.load(GSPMS.subs);

    GSPMS.context.executeQueryAsync(GSPMS.onQuerySucceeded, GSPMS.onQueryFailed);
};

GSPMS.onQuerySucceeded = function () {
    // enumerateDefinition returns ClientCollection object
    var definitionsEnum = GSPMS.workflowDefinitions.getEnumerator();

    var empty = true;

    //console.log('Site ' + GSPMS.web.get_url() + ':');

    // Going through the definitions
    while (definitionsEnum.moveNext()) {

        var def = definitionsEnum.get_current();

        // Displaying information about this definition - DisplayName and Id
        console.log(def.get_displayName() + " (id: " + def.get_id() + ")");

        empty = false;

    }






    //    var enumerator = GSPMS.workflows.getEnumerator();
    var enumerator = GSPMS.subs.getEnumerator();
    while (enumerator.moveNext()) {
        var sub = enumerator.get_current();


        console.log('Web: ' + GSPMS.web.get_url() + ', Subscription: ' + sub.get_name() + ', id: ' + sub.get_id());

        var initiationParams = {};
        GSPMS.servicesManager.getWorkflowInstanceService().startWorkflowOnListItem(sub, GSPMS.item.id, initiationParams);

        context.executeQueryAsync(function (sender, args) {
            console.log('Workflow started.');
        }, GSPMS.onQueryFailed);




        //if (workflow.get_name() == GSPMS.workflowName) {
        //    var url = 'http://' + window.location.hostname + GSPMS.item.get_item("FileRef");
        //    var templateId = '{' + workflow.get_id().toString() + '}';
        //    var workflowParameters = "<root />";
        //    //if (params && params.parameters) {
        //    //    var p;
        //    //    if (params.parameters.length == undefined) p = [params.parameters];
        //    //    p = params.parameters.slice(0);
        //    //    workflowParameters = "<Data>";
        //    //    for (var i = 0; i < p.length; i++)
        //    //        workflowParameters += "<" + p[i].Name + ">" + p[i].Value + "</" + p[i].Name + ">";
        //    //    workflowParameters += "</Data>";
        //    //}
        //    // trigger the workflow
        //    jQuery().SPServices({
        //        operation: "StartWorkflow",
        //        async: true,
        //        item: url,
        //        templateId: templateId,
        //        workflowParameters: workflowParameters,
        //        completefunc: GSPMS.onWorkflowStarted
        //    });
        //    break;
        //}
    }
};

GSPMS.onWorkflowStarted = function () {
    console.log("workflow started");
}

GSPMS.onQueryFailed = function () {
   console.error("Error with Start workflow");
};

GSPMS.getSelectedItem = function () {
    GSPMS.listId = SP.ListOperation.Selection.getSelectedList();
    var items = SP.ListOperation.Selection.getSelectedItems();
    if (items.length == 0) {
        alert("Please select an element in the list");
        return false;
    }
    else if (items.length > 1) {
        alert("You can only submit on item a time.");
        return false;
    }
    else {
        GSPMS.itemId = items[0].id;
        return true;
    }
}

GSPMS.getAssocData = function (name, id, login)
{
    var value = '<dfs:myFields xmlns:xsd="http://www.w3.org/2001/XMLSchema" \
                               xmlns:dms="http://schemas.microsoft.com/office/2009/documentManagement/types" \
                               xmlns:dfs="http://schemas.microsoft.com/office/infopath/2003/dataFormSolution" \
                               xmlns:q="http://schemas.microsoft.com/office/infopath/2009/WSSList/queryFields" \
                               xmlns:d="http://schemas.microsoft.com/office/infopath/2009/WSSList/dataFields" \
                               xmlns:ma="http://schemas.microsoft.com/office/2009/metadata/properties/metaAttributes" \
                               xmlns:pc="http://schemas.microsoft.com/office/infopath/2007/PartnerControls" \
                               xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"> \
                       <dfs:queryFields></dfs:queryFields> \
                       <dfs:dataFields>\
                      <d:SharePointListItem_RW>\
                     <d:Approvers>\
                        <Assignment>\
                           <Assignee>\
                              <pc:Person>\
                                 <pc:DisplayName>'+name+'</pc:DisplayName>\
                                 <pc:AccountId>'+login+'</pc:AccountId>\
                                 <pc:AccountType>User</pc:AccountType>\
                              </pc:Person>\
                           </Assignee>\
                           <d:Stage xsi:nil="true" />\
                           <d:AssignmentType>Serial</d:AssignmentType>\
                        </Assignment>\
                     </d:Approvers>\
                     <d:ExpandGroups>true</d:ExpandGroups>\
                     <d:NotificationMessage>test</d:NotificationMessage>\
                     <d:DueDateforAllTasks xsi:nil="true" />\
                     <d:DurationforSerialTasks>5</d:DurationforSerialTasks>\
                     <d:DurationUnits>Day</d:DurationUnits>\
                     <d:CC />\
                     <d:CancelonRejection>false</d:CancelonRejection>\
                     <d:CancelonChange>true</d:CancelonChange>\
                     <d:EnableContentApproval>false</d:EnableContentApproval>\
                      </d:SharePointListItem_RW>\
                       </dfs:dataFields>\
                    </dfs:myFields> ';
  
    return value;

}


// run this method to start all process
GSPMS.SendToManager = function () {

    if (SP.ClientContext != undefined)
        GSPMS.context = SP.ClientContext.get_current();

    if (GSPMS.getSelectedItem()) {
        GSPMS.setItemAsReadOnly();
        GSPMS.startWorkflow2010();
    }
};