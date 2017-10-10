alert("laoded");

// Go Sport Present Management System
var GSPMS = GSPMS || {};

GSPMS.web = null;
GSPMS.context = null;
GSPMS.item = null;
GSPMS.listId = null;
GSPMS.list = null;
GSPMS.item = null;
GSPMS.workflow = null;
GSPMS.workflowName = "Approval WF 2010";
GSPMS.wfDefinitionId = "d6fed34a-2214-47c0-a0b6-8779e8eb2379";

GSPMS.setItemAsReadOnly = function () {

};

GSPMS.startWorkflow = function () {
    GSPMS.web = GSPMS.context.get_web();
    GSPMS.list = GSPMS.web.get_lists().getById(GSPMS.listId);    
    GSPMS.item = GSPMS.list.getItemById(GSPMS.item.id);
    GSPMS.workflows = GSPMS.list.get_workflowAssociations();
    GSPMS.servicesManager = SP.WorkflowServices.WorkflowServicesManager.newObject(GSPMS.context, GSPMS.web);
    GSPMS.subs = GSPMS.servicesManager.getWorkflowSubscriptionService().enumerateSubscriptionsByDefinition(GSPMS.wfDefinitionId);
    
        
    GSPMS.context.load(GSPMS.list);
    GSPMS.context.load(GSPMS.item);
    GSPMS.context.load(GSPMS.workflows);
    GSPMS.context.load(GSPMS.servicesManager);
    GSPMS.context.load(GSPMS.subs);

    GSPMS.context.executeQueryAsync(GSPMS.onQuerySucceeded, GSPMS.onQueryFailed);
};

GSPMS.onQuerySucceeded = function () {
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
        GSPMS.item = items[0];
        return true;
    }
}

// run this method to start all process
GSPMS.SendToManager = function () {

    if (SP.ClientContext != undefined)
        GSPMS.context = SP.ClientContext.get_current();

    if (GSPMS.getSelectedItem()) {
        GSPMS.setItemAsReadOnly();
        GSPMS.startWorkflow();
    }
};