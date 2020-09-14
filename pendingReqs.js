var pendingReqs = {
    settings: {
        bSendEmails: false,
        fromEmailAddress: "daniel.schauer@foo.bar"
    },
    formDigest: "",
    endpoints: {
        getSiteGroups: _spPageContextInfo.siteAbsoluteUrl +"/_api/web/siteGroups",
        formDigest: _spPageContextInfo.siteAbsoluteUrl +"/_api/contextInfo",
        userProfileById: function(Id){  /* https://docs.microsoft.com/en-us/previous-versions/office/developer/sharepoint-rest-reference/dn499819(v=office.15)#getuserbyid-method */ return _spPageContextInfo.siteAbsoluteUrl +"/_api/web/GetUserById("+ Id +")"; },
        groupUsersByGroupId: function(groupId){  /* https://docs.microsoft.com/en-us/previous-versions/office/developer/sharepoint-rest-reference/dn499819(v=office.15)#getuserbyid-method */ return _spPageContextInfo.siteAbsoluteUrl +"/_api/web/siteGroups("+ groupId +")/users"; },
        deleteAccessRequestById: function(itemId){ return _spPageContextInfo.webAbsoluteUrl +"/_api/web/lists/GetByTitle('Access Requests')/Items("+ itemId +")"; },
        sendEmail: _spPageContextInfo.webServerRelativeUrl +"/_api/SP.Utilities.Utility.SendEmail"
    },
    appendUIControlsAboveView: function(){
        var selectGroup = document.createElement("SELECT");
        selectGroup.id = "selGroup";
        var labelSelectGroup = document.createElement("LABEL");
        labelSelectGroup.setAttribute("for",selectGroup.id);
        labelSelectGroup.id = "LABEL_"+ selectGroup.id;
        labelSelectGroup.innerHTML = "Select Site Group";
        document.querySelector("#contentBox").insertAdjacentElement("afterbegin", labelSelectGroup);
        document.querySelector("#"+ labelSelectGroup.id).insertAdjacentElement("afterend", selectGroup);
        var sendEmail = document.createElement("INPUT");
        sendEmail.setAttribute("type","checkbox");
        sendEmail.setAttribute("checked","checked");
        sendEmail.id = "chkSendEmail";
        var labelSendEmails = document.createElement("LABEL");
        labelSendEmails.setAttribute("for",sendEmail.id);
        labelSendEmails.innerHTML = "Send Email";
        labelSendEmails.id = "LABEL_"+ sendEmail.id;
        document.querySelector("#selGroup").insertAdjacentElement("afterend", labelSendEmails);
        document.querySelector("#"+labelSendEmails.id).insertAdjacentElement("afterend", sendEmail);
        var processRequests = document.createElement("INPUT");
        processRequests.setAttribute("type","button");
        processRequests.id = "btnProcessSelectedRequests";
        processRequests.value = "Process Selected Requests";
        //processRequests.addEventListener("click",pendingReqs.processSelectedRequests);
        document.querySelector("#"+ sendEmail.id).insertAdjacentElement("afterend", processRequests);
        pendingReqs.appendSiteGroupOptions();
    },
    appendSiteGroupOptions: function(){
        var selectGroup = document.getElementById("selGroup");
        var optionGroup = document.createElement("OPTION");
        optionGroup.value = 0;
        optionGroup.innerText = "Select a Group";
        selectGroup.insertAdjacentElement("afterbegin", optionGroup);
        for ( var iGroup = 0; iGroup < pendingReqs.siteGroups.length; iGroup++ ){
            var optionGroupSelect = document.createElement("OPTION");
            optionGroupSelect.value = pendingReqs.siteGroups[iGroup].Id;
            optionGroupSelect.innerText = pendingReqs.siteGroups[iGroup].Title;
            selectGroup.insertAdjacentElement("afterbegin", optionGroupSelect);
        }
    },
    siteGroups: [],
    getSiteGroups: function(url, fxCallback){
        var xhr = new XMLHttpRequest();
        if ( typeof(url) === "undefined" ) url = pendingReqs.endpoints.getSiteGroups;
        xhr.open("GET", url, true);
        xhr.setRequestHeader("accept","application/json;odata=verbose");
        xhr.setRequestHeader("content-type","application/json;odata=verbose");
        xhr.setRequestHeader("X-RequestDigest",pendingReqs.formDigest);
        xhr.onreadystatechange = function(){
            if ( xhr.readyState === 4 ){
                var data = JSON.parse(xhr.response);
                pendingReqs.siteGroups = pendingReqs.siteGroups.concat(data.d.results);
                if ( typeof(data.d.__next) !== "undefined" ){
                    pendingReqs.getSiteGroups(data.d.__next);
                }
                else {
                    if ( typeof(fxCallback) === "function" ){
                        fxCallback();
                    }
                }
            }
        };
        xhr.send();
    },
    getUserProfileById: function(userId, fxCallback){
        var xhr = new XMLHttpRequest();
        xhr.open("GET", pendingReqs.endpoints.userProfileById(userId), false);
        xhr.setRequestHeader("accept","application/json;odata=verbose");
        xhr.setRequestHeader("content-type","application/json;odata=verbose");
        xhr.setRequestHeader("X-RequestDigest",pendingReqs.formDigest);
        xhr.onreadystatechange = function(){
            if ( xhr.readyState === 4 ){
                var data = JSON.parse(xhr.response);
                if ( typeof(fxCallback) === "function" ){
                    fxCallback(data.d);
                }
            }
        };
        xhr.send();
    },
    addUserToGroup: function(userAccount, groupId, rowId, userEmail, fxCallback){
        var xhr = new XMLHttpRequest();
        xhr.open("POST", pendingReqs.endpoints.groupUsersByGroupId(groupId), false);
        xhr.setRequestHeader("accept","application/json;odata=verbose");
        xhr.setRequestHeader("content-type","application/json;odata=verbose");
        //xhr.setRequestHeader("X-RequestDigest",pendingReqs.formDigest);
        var sendData = {
            __metadata: {type:"SP.User"},
            'LoginName':userAccount //'i:0#.w|domain\\user'
        };
        xhr.onreadystatechange = function(){
            if ( xhr.readyState === 4 ){
                if ( xhr.status >= 200 && xhr.status < 300 ){
                    var data = JSON.parse(xhr.response);
                    if ( typeof(fxCallback) === "function" ){
                        fxCallback(data.d);
                    }
                }
            }
        };
        xhr.send(JSON.stringify(sendData));
        
    },
    sendEmailNotification: function(from, to, groupName, fxCallback){
        var xhr = new XMLHttpRequest();
        xhr.open("POST", pendingReqs.endpoints.sendEmail, true);
        xhr.setRequestHeader("accept","application/json;odata=verbose");
        xhr.setRequestHeader("content-type","application/json;odata=verbose");
        xhr.setRequestHeader("X-RequestDigest",pendingReqs.formDigest);
        var sendData = {
            'properties': {
                '__metadata': { 'type': 'SP.Utilities.EmailProperties' },
                'From': from,
                'To': { 'results': [to] },
                'Body': "You've been granted access to |"+ _spPageContextInfo.webTitle +"| via the group |"+ groupName +"|",
                'Subject': "Permissions Granted"
            }
        };
        xhr.onreadystatechange = function(){
            if ( xhr.readyState === 4 ){
                if ( xhr.status >= 200 && xhr.status < 300 ){
                    if ( typeof(fxCallback) === "function" ){
                        fxCallback();
                    }
                }
            }
        };
        xhr.send(JSON.stringify(sendData));
        
    },
    deleteAccessRequest: function(itemIdNum, fxCallback){
        var xhr = new XMLHttpRequest();
        xhr.open("POST", pendingReqs.endpoints.deleteAccessRequestById(itemIdNum), true);
        //xhr.setRequestHeader("accept","application/json;odata=verbose");
        //xhr.setRequestHeader("content-type","application/json;odata=verbose");
        xhr.setRequestHeader("IF-MATCH","*");
        xhr.setRequestHeader("X-HTTP-METHOD","DELETE");
        xhr.setRequestHeader("X-RequestDigest",pendingReqs.formDigest);
        xhr.onreadystatechange = function(){
            if ( xhr.readyState === 4 ){
                if ( xhr.status >= 200 && xhr.status < 400 ){
                    //var data = JSON.parse(xhr.response);
                    // deleting the pending request doesn't remove it from the page, so hide that table row
                    document.querySelector("TR[id$=',"+ itemIdNum +",0']").style.display = "none";
                    if ( typeof(fxCallback) === "function" ){
                        //fxCallback(data.d);
                        fxCallback(itemIdNum);
                    }
                }
            }
        };
        xhr.send();
        
    },
    processSelectedRequests: function(){
        var bFoundSelections = false;
        for ( var ctx in g_ctxDict ) {
            var context = g_ctxDict[ctx];
            //https://docs.microsoft.com/en-us/previous-versions/office/developer/sharepoint-2010/ff409526(v=office.14)
            var selectedItems = SP.ListOperation.Selection.getSelectedItems();
            for ( var i = 0; i < selectedItems.length; i++ ){
                var selection = selectedItems[i];
                var dataRow = null;
                if ( context.ListData.Rows.length > 0 ) {
                    bFoundSelections = true;
                    context.ListData.Rows.forEach(function(row){
                        if ( row.ID === selection.id ){
                            dataRow = row;
                            var itemId = dataRow.ID;
                            var userId = dataRow.RequestedForUserId;
                            var userLogin = "";
                            var userEmail = "";
                            var addToGroupId = document.querySelector("#selGroup").value;
                            pendingReqs.getUserProfileById(userId, function(profile){
                                userLogin = profile.LoginName;
                                userEmail = profile.Email;
                                pendingReqs.addUserToGroup(userLogin, addToGroupId, itemId, userEmail, function(){
                                    if ( pendingReqs.settings.bSendEmails === true ) {
                                        pendingReqs.sendEmailNotification(pendingReqs.settings.fromEmailAddress, userEmail, document.querySelector("#selGroup OPTION[value='"+ addToGroupId +"']").innerText);
                                    }
                                    pendingReqs.deleteAccessRequest(itemId, function(itemIdNumber){
                                        SP.UI.Notify.addNotification("<font color='red'>Unable to delete request |"+ itemIdNumber +"|</font>",false);
                                    });
                                });
                            });
                        }
                    });
                    break;
                }
            }
            if ( bFoundSelections === true ) {
                break;
            }
        }
    },
    onLoad: setTimeout(function(){
        pendingReqs.getSiteGroups(undefined, pendingReqs.appendUIControlsAboveView);
    }, 123)
};
