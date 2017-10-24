var GSG = GSG || {};
GSG.renderer = [];

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
                GSG.renderer.push(r);
                return r.render();
            },
            'NewForm': function () {
                var n = new GSWorkFlowFormRenderer(ctx)
                return n.init();
            },
            'EditForm': function () {
                var n = new GSWorkFlowFormRenderer(ctx)
                return n.init();
            }
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
        SP.SOD.executeFunc('jquery.classywiggle.js', null, EnsureGSGRender);
    }, "SP.js");
}

function EnsureGSGRender() {
    for (var i = 0; i < GSG.renderer.length; i++) {
        GSG.renderer[i].getStatus();
    }
}

function GSWorkflowRenderer(ctx) {

    this.ctx = ctx;
    this.itemId = ctx.CurrentItem.ID;
    this.id = "GSLink" + this.itemId;
    this.debug = false;

    this.getStatus = function () {

        var wfInternalName = this.ctx.ListSchema.Field.find(function (i) { return i.DisplayName == "Approbation 2010"; }).Name;
        //var wfInternalName = this.ctx.ListSchema.Field.find(function (i) { return i.DisplayName == "Approval 2010"; }).Name;

        var url = _spPageContextInfo.siteServerRelativeUrl + "/_api/web/lists('" + ctx.listName + "')/items(" + this.itemId + ")?$select=" + wfInternalName;

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

                // for debug purpose only - display direct link to set item as read only - skip workflow trigger
                //var ro = $("<a href='#' id='RO" + this.itemId + "'>READONLY</a>");
                //ro.click(function () {
                //    var gs = new GSWorkflow(this.itemId);
                //    gs.setItemAsReadOnly();
                //}.bind(this));
                //$("#" + this.id).html(ro);

            }.bind(this),
            error: function () {
                console.error("Failed to get workflow status");
            }
        });

    }.bind(this);

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
        setInterval(function () { wiggleForOneSecond($('#' + id)) }, 5000);

    }.bind(this);

    this.render = function () {
        //this.getStatus();
        var siteUrl = _spPageContextInfo.siteServerRelativeUrl;
        return "<div id=" + this.id + " style='text-align: center;'><img src='" + siteUrl + "/siteassets/gs/ajax-loader-fb.gif' /></div>";
    }
}

function GSWorkflow(item_id) {

    /***** PARAMETERS *****/
    // config rdits-sp13-dev2 - JULIEN
    //this.workflowName = "Approbation 2010";
    //this.wfDefinitionId = "{98D90551-EA55-46A3-A6D0-743C30C008DA}";
    //this.managerInternalField = "Manager";

    //jbes online
    this.workflowName = "Approbation 2010";
    this.wfDefinitionId = "{9279E1FF-1D32-4423-85B7-C7F21998A701}";
    this.managerInternalField = "Nom_x0020_du_x0020_manager";


    /****** END OF PARAMETERS - DO NOT EDIT BELOW UNLESS YOU KNOW MORE OR LESS WHAT YOU ARE DOING *****/

    this.itemId = item_id;
    this.web = null;
    this.context = null;
    this.listId = null;
    this.list = null;
    this.item = null;
    this.workflow = null;
    this.currentUser = null;
    this.manager = null;
    this.dlg = null;

    this.getItem = function () {
        var d = $.Deferred();

        if (SP.ClientContext != undefined)
            this.context = SP.ClientContext.get_current();

        this.listId = SP.ListOperation.Selection.getSelectedList();

        this.web = this.context.get_web();
        this.list = this.web.get_lists().getById(this.listId);
        this.item = this.list.getItemById(this.itemId);
        this.currentUser = this.web.get_currentUser();

        this.context.load(this.web);
        this.context.load(this.currentUser);
        this.context.load(this.item);

        this.context.executeQueryAsync(
            function () { d.resolve() }.bind(this), //this.getManagerInfo.bind(this),
            function (sender, args) { console.error("ERROR 1: " + args.get_message()); }
        );

        return d.promise();
    };

    this.getId = function () {
        return "GSLink" + this.itemId;
    };

    this.getManagerInfo = function () {
        var d = $.Deferred();

        // get selected manager info
        var managerId = this.item.get_item(this.managerInternalField).get_lookupId();
        this.manager = this.web.getUserById(managerId);

        this.context.load(this.manager);
        this.context.executeQueryAsync(
            function () { d.resolve(); }.bind(this),
            function (sender, args) { console.error("ERROR 3: " + args.get_message()); });
        return d.promise();
    };

    this.getAssocData = function (name, login) {

        var assocData = '<dfs:myFields xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:dms="http://schemas.microsoft.com/office/2009/documentManagement/types" xmlns:dfs="http://schemas.microsoft.com/office/infopath/2003/dataFormSolution" xmlns:q="http://schemas.microsoft.com/office/infopath/2009/WSSList/queryFields" xmlns:d="http://schemas.microsoft.com/office/infopath/2009/WSSList/dataFields" xmlns:ma="http://schemas.microsoft.com/office/2009/metadata/properties/metaAttributes" xmlns:pc="http://schemas.microsoft.com/office/infopath/2007/PartnerControls" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
                        '<dfs:queryFields></dfs:queryFields>' +
                        '<dfs:dataFields>' +
                        '<d:SharePointListItem_RW>' +
                        '<d:Approvers>' +
                        '<d:Assignment>' +
                        '<d:Assignee>' +
                        '<pc:Person><pc:DisplayName>' + name + '</pc:DisplayName><pc:AccountId>' + login + '</pc:AccountId><pc:AccountType>User</pc:AccountType></pc:Person>' +
                        '</d:Assignee>' +
                        '<d:Stage xsi:nil="true" />' +
                        '<d:AssignmentType>Serial</d:AssignmentType>' +
                        '</d:Assignment>' +
                        '</d:Approvers>' +
                        '<d:ExpandGroups>true</d:ExpandGroups>' +
                        '<d:NotificationMessage>Please approve</d:NotificationMessage>' +
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
        var d = $.Deferred();

        var exec = function () {
            console.log("changing item permissions...")
            // break inheritance
            this.item.breakRoleInheritance(true);

            var roleAss = this.item.get_roleAssignments();
            var roleDef = this.web.get_roleDefinitions();

            this.context.load(roleAss, 'Include(Member)');
            this.context.load(roleDef);
            this.context.executeQueryAsync(Function.createDelegate(this, function () {

                // remove all but owners
                var roleAssEnum = roleAss.getEnumerator();
                while (roleAssEnum.moveNext()) {
                    var currentAss = roleAssEnum.get_current();
                    // filter groups only
                    var member = currentAss.get_member();
                    if (member.get_principalType() == SP.Utilities.PrincipalType.sharePointGroup) {
                        // keep owners anyhow
                        if (member.get_loginName().indexOf('Owners') == -1) {
                            roleAss.getByPrincipalId(member.get_id()).deleteObject();
                        }
                    }
                }


                // add current user as read only                
                var readRole = roleDef.getByName('Read');
                var collRoleDefinitionBinding = SP.RoleDefinitionBindingCollection.newObject(this.context);
                collRoleDefinitionBinding.add(readRole);

                this.item.get_roleAssignments().add(this.currentUser, collRoleDefinitionBinding);
                this.item.get_roleAssignments().add(this.manager, collRoleDefinitionBinding);

                this.context.executeQueryAsync(
                    Function.createDelegate(this, function () {
                        d.resolve();
                        //SP.UI.Notify.addNotification('Your item has been set to read only.<br>Your page will be automatically refreshed...', false);
                    }),
                    Function.createDelegate(this, function (s, a) {
                        console.error(a.get_message());
                    }));

            }));

        };

        this.getItem()
            .then(this.getManagerInfo.bind(this))
            .then(exec.bind(this));

        return d.promise();
    }

    // run this method to start all process
    this.SendToManager = function () {
        this.dlg = SP.UI.ModalDialog.showWaitScreenWithNoClose("Please wait...", "starting approval process...", null, null);
        this.getItem()
            .then(this.getManagerInfo.bind(this))
            .then(this.triggerWf.bind(this));
        //.then(this.triggerAPIWF.bind(this));
    };

    this.triggerWf = function () {
        var login = this.manager.get_loginName();
        var email = this.manager.get_email();
        var name = this.manager.get_title();
        this.spservices(name, login);
    };

    ///https://gist.github.com/madhur/1584225
    this.spservices = function (approverName, loginName) {
        console.log("triggering wf using $().SPService. Wait for response...");

        if (loginName != null) {

            var assocData = this.getAssocData(approverName, loginName);
            var fileRef = location.protocol + "//" + location.host + this.item.get_item("FileRef");

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
        console.log("workflow request completed. Proceeding...")
        //SP.UI.Notify.addNotification('Your element has been submitted to your manager.', false);
        this.setItemAsReadOnly().then(function () {
            this.dlg.close();
            setTimeout(function () { location.reload(true); }, 1);
        }.bind(this));
    };


    // Attempt to start workflow using client API... not to avail !!
    this.triggerAPIWF = function () {
        console.log("triggering wf using jsom api");

        var login = this.manager.get_loginName();
        var email = this.manager.get_email();
        var name = this.manager.get_title();
        var xml = this.getAssocData(name, login);

        itemGuid = this.item.get_item("GUID").toString();

        //Workflow Services Manager
        var wfServicesManager = new SP.WorkflowServices.WorkflowServicesManager(this.context, this.web);

        var subscription = wfServicesManager.getWorkflowSubscriptionService().getSubscription(this.wfDefinitionId);

        this.context.load(subscription);

        this.context.executeQueryAsync(
            function (sender, args) {
                console.log("Subscription load success. Attempting to start workflow.");
                var inputParameters = {};

                wfServicesManager.getWorkflowInstanceService().startWorkflowOnListItem(subscription, itemGuid, xml);

                this.context.executeQueryAsync(
                    function (sender, args) { console.log("Successfully starting workflow."); },
                    function (sender, args) {
                        console.log("Failed to start workflow.");
                        console.log("Error: " + args.get_message() + "\n" + args.get_stackTrace());
                    }
                );
            }.bind(this),
        function (sender, args) {
            console.log("Failed to load subscription.");
            console.log("Error: " + args.get_message() + "\n" + args.get_stackTrace());
        }
    );

        ////Workflow Interop Service used to interact with SharePoint 2010 Engine Workflows
        //var interopService = wfServicesManager.getWorkflowInteropService()
        //itemGuid = this.item.get_item("GUID").toString();
        ////Start the Site Workflow by Passing the name of the Workflow and the initiation Parameters.
        //var wfName = "Workflow1 - Workflow Start";
        //interopService.startWorkflow(wfName, null, this.listId, itemGuid, {});

        //this.context.executeQueryAsync(
        //    this.onWFStarted.bind(this),
        //    function (sender, args) {
        //        console.error("ERROR 2: " + args.get_message());
        //    });
    }
}

function GSWorkFlowFormRenderer(ctx) {

    this.init = function () {
        var url = _spPageContextInfo.siteServerRelativeUrl + '/_layouts/15/images/discoveryUpdateStats_16x16.png';
        return "<div>Une fois votre demande sauvegardée, cliquez sur l'icône <img src='" + url + "'> qui se trouve dans la liste afin de soumettre votre demande à votre manager.</div>";
    };

}