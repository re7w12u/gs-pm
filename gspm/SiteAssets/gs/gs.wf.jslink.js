(function () {
    if (typeof window.SPClientTemplates === 'undefined')
        return;

    var siteCtx = {};
    siteCtx.Templates = {};
    siteCtx.OnPreRender = loadLibraries;
    siteCtx.Templates.Fields = {
        'GS_WF': {
            'View': function () {
                var r = new GSWorkflowRenderer(ctx);
                return r.render();
            },
            //'NewForm': function () {
            //    var n = new GSWorkFlowNewFormRenderer(ctx)
            //    return n.init();
            //}
        }
    };

    //register the template to render custom field
    window.SPClientTemplates.TemplateManager.RegisterTemplateOverrides(siteCtx);
})();

function loadLibraries(ctx) {
    // ensure WF library loading 
    SP.SOD.executeOrDelayUntilScriptLoaded(function () {
        SP.SOD.registerSod('sp.workflowservices.js', SP.Utilities.Utility.getLayoutsPageUrl("sp.workflowservices.js"));
        SP.SOD.registerSod('jquery.js', ctx.HttpRoot + "/siteassets/gs/jquery-3.2.1.min.js");
        SP.SOD.registerSod('jquery.spservices.js', ctx.HttpRoot + "/siteassets/gs/jquery.SPServices.js");
        SP.SOD.registerSod('jquery.classywiggle.js', ctx.HttpRoot + "/siteassets/gs/jquery.classywiggle.js");

        SP.SOD.registerSodDep('jquery.spservices.js', 'jquery.js');
        SP.SOD.registerSodDep('jquery.classywiggle.js', 'jquery.js');

        SP.SOD.executeFunc('sp.workflowservices.js', "SP.WorkflowServices.WorkflowServicesManager", null);
        SP.SOD.executeFunc('jquery.js', null, null);
        SP.SOD.executeFunc('jquery.spservices.js', null, null);
        SP.SOD.executeFunc('jquery.classywiggle.js', null, null);
    }, "SP.js");
}

function GSWorkflowRenderer(ctx) {

    this.ctx = ctx;
    this.itemId = ctx.CurrentItem.ID;
    this.id = "GSLink" + this.itemId;
    this.debug = false;

    this.getStatus = function () {

        var wfInternalName = this.ctx.ListSchema.Field.find(function (i) { return i.DisplayName == "Approbation 2010"; }).Name;
        //var wfInternalName = this.ctx.ListSchema.Field.find(function (i) { return i.DisplayName == "Approval 2010"; }).Name;

        var url = _spPageContextInfo.siteServerRelativeUrl + "/_api/web/lists('" + ctx.listName + "')/items(" + ctx.CurrentItem.ID + ")?$select=" + wfInternalName;

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

                if (this.debug || status == undefined || status == 0) this.enable();
                else this.disable(status);

            }.bind(this),
            error: function () {
                console.error("Failed to get workflow status");
            }
        });
    }

    this.disable = function (status) {
        var siteUrl = _spPageContextInfo.siteServerRelativeUrl;
        var html = "";
        if (status == 2)// In Progress checknames.png
            html = "<img src='" + siteUrl + "/siteassets/gs/wf_in_progress_16x16.jpg' title='En attente de validation auprès de votre manager.' />";
        else if (status == 5 || status == 16) // completed or approved
            html = "<img src='" + siteUrl + "/_layouts/15/images/componentactive.png' title='la demande a été approuvée.' />";
        else if (status == 4 || status == 17 || status == 15) // canceled or rejected 
            html = "<img src='" + siteUrl + "/_layouts/15/images/componentdegraded.png' title='La demande a été rejetée ou annulée.' />";
        else
            html = "<img src='" + siteUrl + "/_layouts/15/images/removeitem.gif' title='La demande a été rejetée ou annulée.' />";

        $("#" + this.id).html(html);
    }

    this.enable = function () {
        var img = document.createElement("img");
        img.id = "GSIMG" + this.itemId
        img.src = '/_layouts/15/images/discoveryUpdateStats_16x16.png';
        img.className = 'wiggle';

        var a = document.createElement("a");
        a.appendChild(img);
        a.title = "Envoyer la demande à votre manager.";
        a.href = "#";

        a.setAttribute('onclick', 'var x = new GSWorkflow(' + this.itemId + '); x.SendToManager();');

        var div = document.createElement("div");
        div.id = this.id;
        div.appendChild(a);

        this.shake(img.id);

        $("#" + this.id).html(div.outerHTML);
    }.bind(this);

    this.shake = function (id) {
        
        function wiggleForOneSecond(el) {
            el.ClassyWiggle();
            setTimeout(function () { el.ClassyWiggle('stop') }, 1000)
        }

        setInterval(function () { wiggleForOneSecond($('#'+ id)) }, 5000);



        //var classname = 'shake';
        //var icon = $("#" + id);
        //if (icon.hasClass(classname)) {
        //    icon.removeClass(classname);
        //    setTimeout(self.shake.bind(this, id), 3000);
        //}
        //else {
        //    icon.addClass(classname);
        //    setTimeout(self.shake.bind(this, id), 1000);
        //}
    }.bind(this);

    this.render = function () {
        this.getStatus();
        var siteUrl = _spPageContextInfo.siteServerRelativeUrl;
        return "<div id=" + this.id + " style='text-align: center;'><img src='" + siteUrl + "/siteassets/gs/ajax-loader-fb.gif' /></div>";
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

    // config rdits-sp13-dev2 - JULIEN
    this.workflowName = "Approbation 2010";
    this.wfDefinitionId = "{98D90551-EA55-46A3-A6D0-743C30C008DA}";

    //// config rdits-sp13-dev3 - CLE
    //this.workflowName = "Approval 2010";
    //this.wfDefinitionId = "{E47E17E2-B00D-4D61-BED9-065B3DDC1849}";


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
                var xml = self.getAssocData(managerName, login);
                this.spservices(managerName, login);
                //this.triggerWF(xml);
            }.bind(this),
            function (sender, args) {
                console.error("ERROR 3: " + args.get_message());
            });
    }

    this.triggerWF = function (xml) {

        //Workflow Services Manager
        var wfServicesManager = new SP.WorkflowServices.WorkflowServicesManager(this.context, this.web);

        //Workflow Interop Service used to interact with SharePoint 2010 Engine Workflows
        var interopService = wfServicesManager.getWorkflowInteropService()
        itemGuid = this.item.get_item("GUID").toString();
        //Start the Site Workflow by Passing the name of the Workflow and the initiation Parameters.
        interopService.startWorkflow(this.workflowName, null, this.listId, itemGuid, xml);

        this.context.executeQueryAsync(
            function () {
                SP.UI.Notify.addNotification('Your element has been submitted to your manager.<br>Your page will be automatically refreshed...', false);
                setTimeout(function () {
                    this.setItemAsReadOnly();
                    location.reload(true);
                }.bind(this), 1000);

            }.bind(this),
            function (sender, args) {
                console.error("ERROR 2: " + args.get_message());
            });
    }

    this.getAssocData = function (name, login) {

        var assocData = '<dfs:myFields xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:dms="http://schemas.microsoft.com/office/2009/documentManagement/types" xmlns:dfs="http://schemas.microsoft.com/office/infopath/2003/dataFormSolution" xmlns:q="http://schemas.microsoft.com/office/infopath/2009/WSSList/queryFields" xmlns:d="http://schemas.microsoft.com/office/infopath/2009/WSSList/dataFields" xmlns:ma="http://schemas.microsoft.com/office/2009/metadata/properties/metaAttributes" xmlns:pc="http://schemas.microsoft.com/office/infopath/2007/PartnerControls" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
                        '<dfs:queryFields></dfs:queryFields>' +
                        '<dfs:dataFields>' +
                        '<d:SharePointListItem_RW>' +
                        '<d:Approvers>' +
                        '<d:Assignment>' +
                        '<d:Assignee>' +
                        '<pc:Person><pc:DisplayName></pc:DisplayName><pc:AccountId>' + login + '</pc:AccountId><pc:AccountType>User</pc:AccountType></pc:Person>' +
                        '</d:Assignee>' +
                        '<d:Stage xsi:nil="true" />' +
                        '<d:AssignmentType>Serial</d:AssignmentType>' +
                        '</d:Assignment>' +
                        '</d:Approvers>' +
                        '<d:ExpandGroups>true</d:ExpandGroups>' +
                        '<d:NotificationMessage>Please approve <a href="#1">test</a></d:NotificationMessage>' +
                        '<d:DueDateforAllTasks xsi:nil="true" /><d:DurationforSerialTasks xsi:nil="true" />' +
                        '<d:DurationUnits>Day</d:DurationUnits>' +
                        '<d:CC />' +
                        '<d:CancelonRejection>true</d:CancelonRejection>' +
                        '<d:CancelonChange>false</d:CancelonChange>' +
                        '<d:EnableContentApproval>false</d:EnableContentApproval>' +
                        '</d:SharePointListItem_RW>' +
                        '</dfs:dataFields>' +
                        '</dfs:myFields>';

        return assocData;

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

    ///https://gist.github.com/madhur/1584225
    this.spservices = function (approverName, loginName) {

        if (loginName != null) {

            var assocData = this.getAssocData(approverName, loginName);
            var fileRef = "http://" + location.host + this.item.get_item("FileRef");

            $().SPServices({
                operation: "StartWorkflow",
                item: fileRef,
                templateId: this.wfDefinitionId,
                workflowParameters: assocData,
                completefunc: this.onWFStarted.bind(this)
            });



        };
    }

    this.onWFStarted = function () {
        SP.UI.Notify.addNotification('Your element has been submitted to your manager.<br>Your page will be automatically refreshed...', false);
        setTimeout(function () {
            //this.setItemAsReadOnly();
            location.reload(true);
        }.bind(this), 1000);

    };
}

function GSWorkFlowNewFormRenderer(ctx) {

    this.init = function () {
        console.log("new form");

        return "<div>Hello</div>";
    };

}