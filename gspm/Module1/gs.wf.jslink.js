(function () {
    if (typeof window.SPClientTemplates === 'undefined')
        return;
    var siteCtx = {};

    siteCtx.Templates = {};

    siteCtx.Templates.Fields = {
        'GS_WF': {
            'View': function () {
                var r = new GSWorkflowRenderer(ctx);
                return r.render();
            }
        }
    };

    //register the template to render custom field
    window.SPClientTemplates.TemplateManager.RegisterTemplateOverrides(siteCtx);
})();



function GSWorkflowRenderer(ctx) {

    this.ctx = ctx;
    this.itemId = ctx.CurrentItem.ID;
    this.id = "GSLink" + this.itemId;
    this.debug = true;

    this.getStatus = function () {

        var wfInternalName = this.ctx.ListSchema.Field.find(function (i) { return i.DisplayName == "Approbation 2010"; }).Name;

        var url = _spPageContextInfo.siteServerRelativeUrl + "/_api/web/lists('" + ctx.listName + "')/items(" + ctx.CurrentItem.ID + ")?$select=" + wfInternalName;

        var self = this;

        $.ajax({
            url: url,
            type: "GET",
            headers: { "ACCEPT": "application/json;odata=verbose" },
            success: function (data) {
                var status = data.d[wfInternalName]

                //NotStarted = 0
                //FailedOnStart = 1
                //InProgress = 2
                //ErrorOccurred = 3
                //StoppedByUser = 4
                //Completed = 5
                //FailedOnStartRetrying = 6
                //ErrorOccurredRetrying = 7
                //ViewQueryOverflow = 8
                //Canceled = 15
                //Approved = 16
                //Rejected = 17

                if (!this.debug && status != null && [2, 5, 16, 17].indexOf(status) > -1) {
                    this.disable();
                } else {
                    this.enable();
                }
            }.bind(this),
            error: function () {
                alert("Failed to get customer");
            }
        });
    }

    this.disable = function () {
        $("#" + this.id).html("<div>WF already running</div>");
    }

    this.enable = function () {
        var img = document.createElement("img");
        img.src = '/_layouts/15/images/discoveryUpdateStats_16x16.png';

        var a = document.createElement("a");
        a.appendChild(img);
        a.title = "submit to manager";
        a.href = "#";

        a.setAttribute('onclick', 'var x = new GSWorkflow(' + this.itemId + '); x.SendToManager();');

        var div = document.createElement("div");
        div.id = this.id;
        div.appendChild(a);

        $("#" + this.id).html(div.outerHTML);
    }

    this.render = function () {
        this.getStatus();
        var siteUrl = _spPageContextInfo.siteServerRelativeUrl;
        return "<div id=" + this.id + "><img src='" + siteUrl + "/siteassets/gs/ajax-loader-fb.gif' /></div>";
    }
}


function GSWorkflow(item_id) {

    this.itemId = item_id;

    this.web = null;
    this.context = null;
    this.listId = null;
    this.list = null;
    this.item = null;
    this.workflow = null;
    this.currentUser = null;
    this.workflowName = "Approbation 2010";
    this.wfDefinitionId = "{98D90551-EA55-46A3-A6D0-743C30C008DA}";
    //wfDefinitionId = "{67786373-1EA1-452B-8495-2EB736BB0703}";

    this.getItem = function () {
        this.listId = SP.ListOperation.Selection.getSelectedList();

        this.web = this.context.get_web();
        this.list = this.web.get_lists().getById(this.listId);
        this.item = this.list.getItemById(this.itemId);
        this.currentUser = this.web.get_currentUser();

        this.context.load(this.currentUser);
        this.context.load(this.item);

        this.context.executeQueryAsync(
            this.getManagerInfo.bind(this),
            function (sender, args) { console.error("ERROR 1: " + args.get_message()); }
        );
    };

    this.getId = function () {
        return "GSLink" + this.itemId;
    }

    this.getManagerInfo = function () {
        // get selected manager info
        var managerName = this.item.get_item("Manager").get_lookupValue();
        var managerId = this.item.get_item("Manager").get_lookupId();
        var manager = this.web.getUserById(managerId);

        this.context.load(manager);
        let self = this;
        this.context.executeQueryAsync(
            function () {
                var login = manager.get_loginName();
                var email = manager.get_email();
                var xml = self.getAssocData(managerName, managerId, login);
                this.triggerWF(xml);
            }.bind(this),
            function (sender, args) {
                console.error("ERROR 3: " + args.get_message());
            });
    }

    this.triggerWF = function (item, xml) {

        //Workflow Services Manager
        var wfServicesManager = new SP.WorkflowServices.WorkflowServicesManager(this.context, this.web);

        //Workflow Interop Service used to interact with SharePoint 2010 Engine Workflows
        var interopService = wfServicesManager.getWorkflowInteropService()
        itemUniqueId = this.item.get_item("UniqueId").toString();
        itemGuid = this.item.get_item("GUID").toString();
        //Start the Site Workflow by Passing the name of the Workflow and the initiation Parameters.
        interopService.startWorkflow(this.workflowName, null, this.listId, itemGuid, xml);

        this.context.executeQueryAsync(
            function () {
                SP.UI.Notify.addNotification('Your element has been submitted to your manager.', false);
                this.setItemAsReadOnly();
            }.bind(this),
            function (sender, args) {
                console.error("ERROR 2: " + args.get_message());
            });
    }

    this.getAssocData = function (name, id, login) {
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
                                 <pc:DisplayName>'+ name + '</pc:DisplayName>\
                                 <pc:AccountId>'+ login + '</pc:AccountId>\
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

    this.onQueryFailed = function () {
        console.error("Error with Start workflow");
    };

    this.setItemAsReadOnly = function () {

        console.log("setItemAsReadOnly");

        //GSPMS.item.breakRoleInheritance(true);
        //GSPMS.item.get_roleAssignments().getByPrincipal(GSPMS.currentUser).deleteObject();

        ////var collRoleDefinitionBinding = SP.RoleDefinitionBindingCollection.newObject(GSPMS.context)
        ////collRoleDefinitionBinding.add(GSPMS.web.get_roleDefinitions().getByType(SP.RoleType.reader));

        ////GSPMS.item.get_roleAssignments().add(GSPMS.currentUser, collRoleDefinitionBinding);
        ////GSPMS.context.load(GSPMS.currentUser);
        ////GSPMS.context.load(GSPMS.item);

        //GSPMS.context.executeQueryAsync(
        //    function () {
        //        console.log("item set as read only");
        //        SP.UI.Notify.addNotification('item set as read only', false);
        //}, function (sender, args) {
        //    console.error("ERROR 4 : " + args.get_message());
        //});
    }

    // run this method to start all process
    this.SendToManager = function () {

        if (SP.ClientContext != undefined)
            this.context = SP.ClientContext.get_current();

        this.getItem();

    };


}

