/*http://expertsoverlunch.com/sandbox/Lists/Tasks/NewForm.aspx*/
/*http://expertsoverlunch.com/sandbox/Lists/PeopleAndGroupPickers/NewForm.aspx*/
var pickerListeners = pickerListeners || [];
var picker = picker || {
    instances: {},
    formFields: {},
    pageWebParts: [],
    bDebug: false,
    editMode: false,
    log: function(message, bIgnoreDebugReq) {
        if ( typeof(bIgnoreDebugReq) === "undefined" ) { var bIgnoreDebugReq = false; }
        if ( picker.bDebug === true || bIgnoreDebugReq === true ) {
            try{console.log(message);}catch(err){}
        }
    },
    isPageInEditMode: function() {
        /*https://sharepoint.stackexchange.com/questions/149096/a-way-to-identify-when-page-is-in-edit-mode-for-javascript-purposes*/
        var result = (window.MSOWebPartPageFormName != undefined) && ((document.forms[window.MSOWebPartPageFormName] && document.forms[window.MSOWebPartPageFormName].MSOLayout_InDesignMode && ("1" == document.forms[window.MSOWebPartPageFormName].MSOLayout_InDesignMode.value)) || (document.forms[window.MSOWebPartPageFormName] && document.forms[window.MSOWebPartPageFormName]._wikiPageMode && ("Edit" == document.forms[window.MSOWebPartPageFormName]._wikiPageMode.value)));
        picker.editMode = result || false;
        return result || false;
    },
    ajax: {
        lastCall: {},
        create: function(restURL, object, fxCallback, fxFailed) {
            picker.ajax.lastCall = { xhr: null, readyState: null, data: null, status: null, url: restURL, error: null };
            var xhr = new XMLHttpRequest();
            xhr.open('POST', restURL, true);
            xhr.setRequestHeader("X-RequestDigest", document.getElementById("__REQUESTDIGEST").value);
            xhr.setRequestHeader("IF-MATCH", "*");
            xhr.setRequestHeader("X-HTTP-Method", "POST");
            xhr.setRequestHeader("accept", "application/json;odata=verbose");
            xhr.setRequestHeader("content-type", "application/json;odata=verbose");
            xhr.onreadystatechange = function() {
                picker.ajax.lastCall.readyState = xhr.readyState;
                picker.ajax.lastCall.status = xhr.status;
                if (xhr.readyState === 4) {
                    if (xhr.status !== 201) {
                        picker.ajax.lastCall.data = xhr.response;
                        if (typeof(fxFailed) === "function") {
                            fxFailed(xhr, xhr.response, xhr.status);
                        }
                    } else {
                        /*var resp = JSON.parse(xhr.response);*/
                        picker.ajax.lastCall.data = xhr.response;
                        if (typeof(fxCallback) === "function") {
                            fxCallback(xhr, xhr.response);
                        }
                    }
                }
            };
            xhr.send(object);
        },
        update: function(restURL, object, fxCallback, fxFailed) {
            picker.ajax.lastCall = { xhr: null, readyState: null, data: null, status: null, url: restURL, error: null };
            var xhr = new XMLHttpRequest();
            xhr.open('POST', restURL, true);
            xhr.setRequestHeader("X-RequestDigest", document.getElementById("__REQUESTDIGEST").value);
            xhr.setRequestHeader("IF-MATCH", "*");
            xhr.setRequestHeader("X-HTTP-Method", "MERGE");
            xhr.setRequestHeader("accept", "application/json;odata=verbose");
            xhr.setRequestHeader("content-type", "application/json;odata=verbose");
            xhr.onreadystatechange = function() {
                picker.ajax.lastCall.readyState = xhr.readyState;
                picker.ajax.lastCall.status = xhr.status;
                if (xhr.readyState === 4) {
                    if (xhr.status !== 200 && xhr.status !== 204) {
                        picker.ajax.lastCall.data = xhr.response;
                        if (typeof(fxFailed) === "function") {
                            fxFailed(xhr, xhr.response, xhr.status);
                        }
                    } else {
                        var resp = xhr.response;
                        picker.ajax.lastCall.data = resp;
                        if (typeof(fxCallback) === "function") {
                            fxCallback(xhr, resp);
                        }
                    }
                }
            };
            xhr.send(object);
        },
        delete: function(restURL, object, fxCallback) {
            picker.ajax.lastCall = { xhr: null, readyState: null, data: null, status: null, url: restURL, error: null };
            var xhr = new XMLHttpRequest();
            xhr.open('POST', restURL, true);
            xhr.setRequestHeader("X-RequestDigest", document.getElementById("__REQUESTDIGEST").value);
            xhr.setRequestHeader("IF-MATCH", "*");
            xhr.setRequestHeader("X-HTTP-Method", "DELETE");
            xhr.setRequestHeader("accept", "application/json;odata=verbose");
            xhr.setRequestHeader("content-type", "application/json;odata=verbose");
            xhr.onreadystatechange = function() {
                picker.ajax.lastCall.readyState = xhr.readyState;
                picker.ajax.lastCall.status = xhr.status;
                if (xhr.readyState === 4) {
                    if (xhr.status !== 200) {
                        picker.ajax.lastCall.data = xhr.response;
                        if (typeof(fxFailed) === "function") {
                            fxFailed(xhr, xhr.response, xhr.status);
                        }
                    } else {
                        var resp = xhr.response;
                        picker.ajax.lastCall.data = resp;
                        if (typeof(fxCallback) === "function") {
                            fxCallback(xhr, resp);
                        }
                    }
                }
            };
            xhr.send(object);
        },
        read: function(restURL, fxCallback, fxLastPage, fxFailed, sDataType) {
            if ( typeof(sDataType) === "undefined" ){
                var sDataType = "json";
            }
            picker.ajax.lastCall = { xhr: null, readyState: null, data: null, status: null, url: restURL, error: null };
            var xhr = new XMLHttpRequest();
            xhr.open('GET', restURL, true);
            var now = new Date();
            /* 4 hours later */
            var later = new Date(now.valueOf()+(1000*60*60*4));
            xhr.setRequestHeader("Expires", later);
            xhr.setRequestHeader("Last-Modified", now);
            xhr.setRequestHeader("Cache-Control", "Public");
            xhr.setRequestHeader("X-RequestDigest", document.getElementById("__REQUESTDIGEST").value);
            switch (sDataType) {
                case "json":
                    xhr.setRequestHeader("accept", "application/json;odata=verbose");
                    xhr.setRequestHeader("content-type", "application/json;odata=verbose");
                    break;
                case "xml":
                    xhr.setRequestHeader("content-type", "text/xml");
                    break;
                case "html":
                    xhr.setRequestHeader("content-type", "text/html");
                    break;
                case "text":
                    xhr.setRequestHeader("content-type", "text/plain");
                    break;
                case "css":
                    xhr.setRequestHeader("content-type", "text/css");
                    break;
                default:
                    picker.log("picker.ajax.read... unknown sDataType value |"+ sDataType +"|",true);
            }

            xhr.onreadystatechange = function() {
                picker.ajax.lastCall.readyState = xhr.readyState;
                picker.ajax.lastCall.status = xhr.status;
                if (xhr.readyState === 4) {
                    if (xhr.status !== 200) {
                        picker.ajax.lastCall.data = xhr.response;
                        if (typeof(fxFailed) === "function") {
                            fxFailed(xhr, xhr.response, xhr.status);
                        }
                    } else {
                        switch (sDataType){
                            case "json":
                                var resp = JSON.parse(xhr.response);
                                break;
                            case "html":
                                var resp = document.createElement("div");
                                resp.innerHTML = xhr.response;                            
                                break;
                            case "xml":
                                var resp = xhr.responseXML;
                                break;
                            default: 
                                var resp = xhr.response;
                        }
                        picker.ajax.lastCall.data = resp;
                        if (typeof(fxCallback) === "function") {
                            fxCallback(xhr, resp);
                        }
                        if (typeof(resp.d.__next) !== "undefined") {
                            picker.ajax.read(resp.d.__next, fxCallback, fxLastPage);
                        } else if (typeof(fxLastPage) === "function") {
                            fxLastPage(xhr, resp);
                        }
                    }
                }
            };
            picker.ajax.lastCall.xhr = xhr;
            xhr.send();
        },
        createSettingsList: function(afterCreatingSettingsListFx){
            var sListDisplayName = "dsMagic-pickers";
            var listDef = null;
            if ( typeof(afterCreatingSettingsListFx) === "undefined" ) {
                var afterCreatingSettingsListFx = function(){
                    picker.log("Created and configured the settings list",true);
                };
            }
            var oCreateList = {
                __metadata: { type: 'SP.List' }, 
                AllowContentTypes: false, 
                BaseTemplate: 100, 
                ContentTypesEnabled: false, 
                Description: "Stores the settings for dsMagic pickers JS API", 
                Title: "dsMagic-pickers"
            };
            picker.ajax.create(_spPageContextInfo.webAbsoluteUrl+"/_api/web/lists", JSON.stringify(oCreateList), function(x, d){
                picker.log("dsMagic-picker settings list created",true);
                picker.log(x,true);
                picker.log(d,true);
                var responseData = JSON.parse(d);
                var listRestURL = responseData.d.__metadata.uri;
                var oSettings1 = {
                    __metadata: { type: 'SP.List' }, 
                    EnableVersioning: true,
                    EnableModeration: true,
                    EnableAttachments: false
                };
                picker.ajax.update(listRestURL, JSON.stringify(oSettings1), function(x1, d1){
                    picker.log("dsMagic-picker list - versioning enabled",true);
                    var oSettings2 = {
                        __metadata: { type: 'SP.List' }, 
                        MajorVersionLimit: 10,
                        DraftVersionVisibility: 2,
                        MajorWithMinorVersionsLimit: 25
                    };
                    picker.ajax.update(listRestURL, JSON.stringify(oSettings2), function(x2, d2){
                        picker.log("dsMagic-picker list - versioning limits set",true);
                        var fld_blbSettings = {
                            __metadata: { type: 'SP.FieldMultiLineText' }, 
                            Title: "blbSettings", 
                            FieldTypeKind: 3, 
                            Required: false, 
                            EnforceUniqueValues: false, 
                            StaticName: "blbSettings",
                            NumberOfLines: 1,
                            RichText: false,
                            AppendOnly: false,
                            RestrictedMode: true,
                            WikiLinking: false
                        };
                        picker.ajax.create(listRestURL+"/Fields", JSON.stringify(fld_blbSettings), function(xf1, df1){
                            picker.log("dsMagic-picker list - added field 'blbSettings'",true);
                            var fld_fieldName = {
                                __metadata: { type: 'SP.FieldText' }, 
                                Title: "fieldName", 
                                FieldTypeKind: 2, 
                                Required: false, 
                                EnforceUniqueValues: false, 
                                StaticName: "fieldName"
                            };
                            picker.ajax.create(listRestURL+"/Fields", JSON.stringify(fld_fieldName), function(xf2, df2){
                                picker.log("dsMagic-picker list - added field 'fieldName'",true);
                                if ( typeof(afterCreatingSettingsListFx) === "function" ){
                                    afterCreatingSettingsListFx();
                                }
                            }, function(){
                                picker.log("dsMagic-picker list - failed to add field 'fieldName'",true);
                            });    
                        }, function(){
                            picker.log("dsMagic-picker list - failed to add field 'blbSettings'",true);
                        });
                        
                    }, function(){
                        picker.log("failed to set version limits on dsMagic-picker",true);
                    });
                }, function(){
                    picker.log("failed to enable versioning on dsMagic-picker",true);
                });
            }, function(){
                picker.log("failed to create dsMagic-picker settings list",true);
            });
            /*
            ds.rest.list.create(100, sListDisplayName, "stores the settings for the forms on your site", function(cD, cS, cX){
                listDef = cD.d; 
                var oSettings1 = {
                    EnableVersioning: true,
                    EnableModeration: true,
                    EnableAttachments: false
                };
                ds.rest.list.update(sListDisplayName, oSettings1, function(){
                    picker.log("Enabled Versioning and disabled Attachments for |"+sListDisplayName+"||");
                    var oSettings2 = {
                        MajorVersionLimit: 10,
                        DraftVersionVisibility: 2,
                        MajorWithMinorVersionsLimit: 25
                    };
                    ds.rest.list.update(sListDisplayName, oSettings2, function(){
                        picker.log("Finished configuring Versioning for |"+sListDisplayName+"|");
                        var to1 = setTimeout(function(){
                            var oField = {
                                Description:"The form name, as it appears in the URL",
                                StaticName:"FormNameURL",
                                EnforceUniqueValues: false
                            };
                            ds.rest.list.addField(ds.lists[sListDisplayName].Id, "FormNameURL", 2, false, false, "FormNameURL", null, oField, function(){
                                picker.log("Added field |FormNameURL| to list");
                            });
                            var to2 = setTimeout(function(){
                                var oField = {
                                    StaticName:"blbRelRecs"
                                };
                                ds.rest.list.addField(ds.lists[sListDisplayName].Id, "blbRelRecs", 3, false, false, "blbRelRecs", null, oField, function(d,s,x){
                                    picker.log("Added field |blbRelRecs| to list");
                                    var body = {
                                        __metadata: { type: 'SP.FieldMultiLineText' }, 
                                        NumberOfLines: 1,
                                        RichText: false,
                                        AppendOnly: false,
                                        RestrictedMode: true,
                                        WikiLinking: false
                                    };
                                    ds.$.ajax({
                                        url: d.d.__metadata.uri,
                                        method: "POST",
                                        data: JSON.stringify(body),
                                        headers: {
                                            "X-HTTP-Method":"MERGE",
                                            "IF-MATCH": "*"
                                        }
                                    }).done(function(data,textStatus,jqXHR){
                                        ds.rest.lastCall.status = textStatus;
                                        ds.rest.lastCall.data = data;
                                        ds.rest.lastCall.jqXHR = jqXHR;
                                        picker.log("List field |blbRelRecs| updated via REST");
                                        picker.log(listDef.Fields,true);
                                        ds.$.ajax({
                                            url: listDef.Fields.__deferred.uri+"?$filter=StaticName eq 'Title'",
                                            method: "GET",
                                            data: null
                                        }).done(function(data,textStatus,jqXHR){
                                            picker.log("Got definition of existing list Title field...");
                                            picker.log(data,true);
                                            var body2 = {
                                                __metadata: { type: data.d.results[0].__metadata.type }, 
                                                Title: "ListNameURL",
                                                StaticName: data.d.results[0].Title
                                            };
                                            picker.log("Attempting update to Title field's Display name...");
                                            ds.$.ajax({
                                                url: data.d.results[0].__metadata.uri,
                                                method: "POST",
                                                data: JSON.stringify(body2),
                                                headers: {
                                                    "X-HTTP-Method":"MERGE",
                                                    "IF-MATCH": "*"
                                                }
                                            }).done(function(data,textStatus,jqXHR){
                                                ds.rest.lastCall.status = textStatus;
                                                ds.rest.lastCall.data = data;
                                                ds.rest.lastCall.jqXHR = jqXHR;
                                                picker.log("List field |Title| updated to |ListNameURL| via REST");
                                                if ( typeof(afterCreatingSettingsListFx) === "function" ) {
                                                    afterCreatingSettingsListFx();
                                                }
                                            }).fail(function(jqXHR, textStatus, errorThrown){
                                                ds.rest.lastCall.status = textStatus;
                                                ds.rest.lastCall.errorThrown = errorThrown;
                                                ds.rest.lastCall.jqXHR = jqXHR;
                                            });
                                        }).fail(function(jqXHR,textStatus,errorThrown){
                                            ds.rest.lastCall.status = textStatus;
                                            ds.rest.lastCall.errorThrown = errorThrown;
                                            ds.rest.lastCall.jqXHR = jqXHR;
                                        });
                                    }).fail(function(jqXHR, textStatus, errorThrown){
                                        ds.rest.lastCall.status = textStatus;
                                        ds.rest.lastCall.errorThrown = errorThrown;
                                        ds.rest.lastCall.jqXHR = jqXHR;
                                    });
                                });
                            },234);
                        },123);
                    });
                });
            });
            */
        },
        captureSettings: function(formURL, fieldName, blbSettings){
            /*
            var formURL = _spPageContextInfo.serverRequestPath;
            var fieldName = "PrimaryResource"
            var blbSettings = JSON.stringify({
                "fieldDisplayName":"PrimaryResource",
                "bCaptureProfileDetails":true,
                "arrMapping":[
                    {"fin":"ResourceAccount","userProperty":"picker.userProperties.LoginName"},
                    {"fin":"ResourceEmail","userProperty":"picker.userProperties.Email"},
                    {"fin":"ResourceGroups","userProperty":"picker.userProperties.Groups.results"},
                    {"fin":"ResourcePeers","userProperty":"picker.userProperties.Groups.results|Users.results"}
                ]
            });
            */
            var oItem = {
                __metadata: {type: 'SP.Data.DsMagicpickersListItem'},
                Title: formURL,
                fieldName: fieldName,
                blbSettings: JSON.stringify(blbSettings)
            };
            picker.ajax.create(_spPageContextInfo.webAbsoluteUrl +"/_api/web/lists/GetByTitle('dsMagic-pickers')/Items", JSON.stringify(oItem), function(x,d){
                picker.log("Created list item for Settings",true);
                picker.log(x,true);
                picker.log(d,true);
            }, function(){
                picker.log("Failed to create list item for Settings",true);
            });
        },
        getSettings: function(fxAfter, fxFailed){
            var restURL = _spPageContextInfo.webAbsoluteUrl +"/_api/web/lists/GetByTitle('dsMagic-pickers')/Items?$filter=Title eq '"+ _spPageContextInfo.serverRequestPath +"'";
            picker.ajax.read(restURL, function(x,d){
                picker.log("Got a page of settings",true);
                for ( var iD = 0; iD < d.d.results.length; iD++ ){
                    picker.settings.push(d.d.results[iD]);
                    var blbSettings = JSON.parse(d.d.results[iD].blbSettings);
                    pickerListeners.push(blbSettings);
                }
            }, function(xLast, dLast){
                picker.log("Last page of settings",true);
                //picker.setupListeners();
                if ( typeof(fxAfter) === "function" ) {
                    fxAfter();
                }
            }, function(){
                picker.log("Failed to get settings",true);
                if ( typeof(fxFailed) === "function" ) {
                    fxFailed();
                }
            });
        },
        updateSettings: function(itemID, blbSettings){
            var oItem = {
                __metadata: {type: 'SP.Data.DsMagicpickersListItem'},
                blbSettings: JSON.stringify(blbSettings)
            };
            picker.ajax.update(_spPageContextInfo.webAbsoluteUrl +"/_api/web/lists/GetByTitle('dsMagic-pickers')/Items("+ itemID +")", JSON.stringify(oItem), function(x,d){
                picker.log("Updated list item for Settings",true);
                picker.log(x,true);
                picker.log(d,true);
            }, function(){
                picker.log("Failed to update list item for Settings",true);
            });
        }
    },
    settings: [],
    userProperties: {},
    userGroups: {},
    getUserProperties: function(afterFx){
        /* current user */
        picker.ajax.read(
            _spPageContextInfo.webAbsoluteUrl +"/_api/web/CurrentUser?$expand=Groups,UserId,Groups/Users",
            function(xhr, data){
                picker.log(data);
                for ( userProp in data.d ){
                    try{
                        picker.userProperties[userProp] = data.d[userProp];
                    }
                    catch(err){
                        picker.log("failed to capture user prop |"+ userProp +"|");
                    }
                }
            },
            function(xhr, data){
                picker.log(data);
                if ( typeof(afterFx) === "function" ){
                    afterFx();
                }
            },
            function(){
                picker.log("failed to get user properties");
            }
        );
        
    },
    getUserPropertiesByLoginName: function(loginName, afterFx){
        picker.userProperties = {};
        picker.ajax.read(
            _spPageContextInfo.webAbsoluteUrl +"/_api/web/SiteUsers(@v)?@v='"+encodeURIComponent(loginName)+"'&$expand=Groups,UserId,Groups/Users",
            function(xhr, data){
                picker.log(data);
                for ( userProp in data.d ){
                    try{
                        picker.userProperties[userProp] = data.d[userProp];
                        if ( userProp === "Groups" ){
                            if ( typeof(data.d.Groups.results) !== "undefined" ){
                                for ( var iG = 0; iG < data.d.Groups.results.length; iG++ ){
                                    picker.userGroups[data.d.Groups.results[iG].Title] = data.d.Groups.results[iG];
                                }
                            }
                        }
                    }
                    catch(err){
                        picker.log("failed to capture user prop |"+ userProp +"|");
                    }
                }
            },
            function(xhr, data){
                picker.log(data);
                if ( typeof(afterFx) === "function" ){
                    afterFx();
                }
            },
            function(){
                picker.log("failed to get user properties");
            }
        );
    },
    getPageWebPartsAndFormFields: function(fxEach, afterFx){
        var collWPs = document.querySelectorAll("[id^='MSOZoneCell_WebPartWPQ']");
        for (var webPart = 0; webPart < collWPs.length; webPart++) {
            var wp = {
                id: collWPs[webPart].id,
                num: 0
            };
            try { wp.num = collWPs[webPart].id.replace("MSOZoneCell_WebPartWPQ", ""); } catch (err) {}
            try { wp.title = document.getElementById("WebPartTitleWPQ" + wp.num).innerHTML; } catch (err) {}
            try { wp.body = document.getElementById("WebPartWPQ" + wp.num).innerHTML; } catch (err) {}
            try { wp.guid = document.getElementById("WebPartWPQ" + wp.num).getAttribute("webpartid"); } catch (err) {}
            picker.pageWebParts.push(wp);
            /* check if we have a WPQ#FormCtx object defined for this web part; then loop thru its fields and capture them in picker.formFields */
            try{
                var wpCTX = window["WPQ"+wp.num+"FormCtx"];
                if ( typeof(picker.formFields["WPQ"+wp.num+"FormCtx"]) === "undefined" ){
                    picker.formFields["WPQ"+wp.num+"FormCtx"] = {};
                }
                for ( formField in wpCTX.ListSchema ) {
                    picker.formFields["WPQ"+wp.num+"FormCtx"][formField] = wpCTX.ListSchema[formField];
                }        
            }
            catch(err){
                try{console.log("WebPart |"+wp.num+"| is not a standard form");}catch(err){}
            }
            if (typeof(fxEach) === "function") {
                fxEach(wp);
            }
        }
        if ( typeof(afterFx) === "function" ){
            afterFx();
        }
    },
    getAllPickers: function(afterFx){
        for ( pickerControl in SPClientPeoplePicker.SPClientPeoplePickerDict ) {
            if ( typeof(picker.instances[SPClientPeoplePicker.SPClientPeoplePickerDict[pickerControl].TopLevelElementId]) === "undefined" ) {
                picker.instances[SPClientPeoplePicker.SPClientPeoplePickerDict[pickerControl].TopLevelElementId] = {
                    spcsom: SPClientPeoplePicker.SPClientPeoplePickerDict[pickerControl],
                    fieldDisplayName: "",
                    fin: "",
                    fieldGUID: SPClientPeoplePicker.SPClientPeoplePickerDict[pickerControl].TopLevelElementId.replace("_$ClientPeoplePicker",""),
                    arrMapping: []
                };
                picker.instances[SPClientPeoplePicker.SPClientPeoplePickerDict[pickerControl].TopLevelElementId].fieldGUID = picker.instances[SPClientPeoplePicker.SPClientPeoplePickerDict[pickerControl].TopLevelElementId].fieldGUID.substr(picker.instances[SPClientPeoplePicker.SPClientPeoplePickerDict[pickerControl].TopLevelElementId].fieldGUID.indexOf("_")+1,38);
                var pickerFormField = picker.findFormField("TopLevelElementId",SPClientPeoplePicker.SPClientPeoplePickerDict[pickerControl].TopLevelElementId);
                try{
                    picker.instances[SPClientPeoplePicker.SPClientPeoplePickerDict[pickerControl].TopLevelElementId].fieldDisplayName = pickerFormField.Title;
                    picker.instances[SPClientPeoplePicker.SPClientPeoplePickerDict[pickerControl].TopLevelElementId].fin = pickerFormField.Name;
                }
                catch(err){}
            }
        }
        if ( typeof(afterFx) === "function" ){
            afterFx();
        }
    },
    findFormField: function(propertyName, propertyValue){
        var oReturn = null;
        var bFoundFormField = false;
        for ( formWP in picker.formFields ) {
            for ( formField in picker.formFields[formWP] ) {
                if ( typeof(picker.formFields[formWP][formField][propertyName]) !== "undefined" ) {
                    if ( picker.formFields[formWP][formField][propertyName] === propertyValue ) {
                        oReturn = picker.formFields[formWP][formField];
                        bFoundFormField = true;
                        break;
                    }
                }    
            }
            if ( bFoundFormField === true ) {
                break;
            }
        }
        return oReturn;
    },
    showConfigButtons: function(afterFx){
        for ( control in picker.instances ) {
            var pickerControl = null, pickerWrapper = null, configButton = null, configPopup = null;
            pickerControl = document.getElementById(control);
            pickerControl.style.display = "inline-block";
            pickerControl.style.width = "90%";
            pickerWrapper = pickerControl.parentElement;
            pickerWrapper.style.display = "inline-block";
            pickerWrapper.style.width = "85%";

            pickerWrapper.parentElement.style.position = "relative";
            pickerWrapper.parentElement.style.top = "0px";
            pickerWrapper.parentElement.style.left = "0px";

            configButton = document.createElement("SPAN");
            configButton.innerHTML = "&#8286;";
            configButton.className = "pickerConfigButton buttonish";
            configButton.id = pickerControl.id +"_configButton";
            /*pickerWrapper.parentElement.appendChild(configButton);*/
            //https://developer.mozilla.org/en-US/docs/Web/API/Element/insertAdjacentElement
            pickerWrapper.insertAdjacentElement('afterend',configButton);

            configPopup = document.createElement("DIV");
            configPopup.id = pickerControl.id +"_configWrapper";
            configPopup.className = "pickerConfigPopup";
            var arrH = ["<DIV class='configPopupHeader'><span class='closeButton buttonish' id='"];
            arrH.push(control);
            arrH.push("_configCloseButton'>&times;</span></DIV><H3><DIV class='configForField'>");
            /*arrH.push(control);*/
            arrH.push(picker.findFormField("TopLevelElementId",control).Title);
            arrH.push("</DIV></H3><H4>Upon resolved user, populate form fields...</H4><TABLE><thead><tr><td>Form Field</td><td>User Property</td><td>Add</td><td>Remove</td></tr></thead><tbody id='");
            arrH.push(control);
            arrH.push("_configMappingRowsWrapper'><tr class='mappingRow'><td class='listField'><select class='listField'><option value='0'>Select list field</option>");
            for ( formWP in picker.formFields ) {
                for ( formField in picker.formFields[formWP] ) {
                    arrH.push("<option value='");
                    arrH.push(picker.formFields[formWP][formField].Id);
                    arrH.push("'>");
                    arrH.push(picker.formFields[formWP][formField].Name);
                    arrH.push("</option>");
                }
            }
            arrH.push("</select></td><td class='userProperty'>");
            /*
            arrH.push("<select class='userProperty'><option value='0'>Select user property</option>");
            for ( userProperty in picker.userProperties ) {
                arrH.push("<option value='");
                arrH.push(userProperty);
                arrH.push("'>");
                arrH.push(userProperty);
                arrH.push("</option>");
            }
            arrH.push("</select>");
            */
            arrH.push("<INPUT class='userProperty' type='text' placeholder='e.g. picker.userProperties.LoginName' value='' />");
            arrH.push("</td><td><span class='addNew buttonish'>&plus;&nbsp;Add</span></td><td><span class='removeMapping buttonish'>&minus;&nbsp;Remove</span></td></tr></tbody></TABLE><DIV class='configPopupFooter'><span class='saveButton buttonish' id='");
            arrH.push(control);
            arrH.push("_configSaveButton'>Save</span></DIV><DIV class='saveSettings'></DIV>");
            configPopup.innerHTML = arrH.join("");

            /*pickerWrapper.parentElement.appendChild(configPopup);*/
            configButton.insertAdjacentElement('afterend',configPopup);
            if ( pickerWrapper.parentElement.querySelectorAll(".ms-metadata").length > 0 ){
                configPopup.insertAdjacentElement('afterend',document.createElement("BR"));
            }
            
            configButton.addEventListener("click", function(e){
                this.nextSibling.style.display = "block";
            });

            configPopup.querySelector(".closeButton").addEventListener("click", function(e){
                this.parentElement.parentElement.style.display = "none";
            });

            configPopup.querySelector(".saveButton").addEventListener("click", function(e){
                picker.log("Save button clicked");
                picker.log(e);
                picker.log(this);
                var popupWrapper = picker.parentsUntilMatchingSelector(this, ".pickerConfigPopup");
                var fieldName = popupWrapper.querySelector(".configForField").innerText;
                var saveConfig = {};
                saveConfig.fieldDisplayName = fieldName;
                saveConfig.bCaptureProfileDetails = false;
                saveConfig.arrMapping = [];
                
                
                var collMapping = this.parentElement.previousSibling.querySelectorAll(".mappingRow");
                for ( var iR = 0; iR < collMapping.length; iR++ ) {
                    /*try{console.log(collMapping[iR]);}catch(err){}*/
                    var map = collMapping[iR];
                    var formFieldGUID = map.querySelector("SELECT.listField").value;
                    var formFieldName = document.querySelector("OPTION[value='"+ formFieldGUID +"']").innerText;
                    //var userProperty = map.querySelector("SELECT.userProperty").value;
                    var userProperty = map.querySelector("INPUT.userProperty").value;
                    picker.log("Capture user profile property |"+ userProperty +"| in field named |"+ formFieldName +"|");
                    /*
                    for ( formWP in picker.formFields ) {
                        for ( formField in picker.formFields[formWP] ) {
                            if ( typeof(picker.formFields[formWP][formField].TopLevelElementId) !== "undefined" ) {
                                if ( picker.formFields[formWP][formField].TopLevelElementId === control ) {
                                    arrH.push(picker.formFields[formWP][formField].Title);
                                    break;
                                }
                            }    
                        }
                    }
                    */
                    saveConfig.arrMapping.push({"fin":formFieldName, "userProperty": userProperty});
                    saveConfig.bCaptureProfileDetails = true;
                }
                popupWrapper.querySelector(".saveSettings").innerText = JSON.stringify(saveConfig);
                /* update the existing settings record for this field */
                var bFoundExisting = false;
                for ( var iSettings = 0; iSettings < picker.settings.length; iSettings++ ){
                    //picker.log(picker.settings[iSettings],true);
                    if ( picker.settings[iSettings].fieldName === fieldName ) {
                        //picker.log(picker.settings[iSettings].Id,true);
                        bFoundExisting = true;
                        picker.ajax.updateSettings(picker.settings[iSettings].Id, JSON.parse(popupWrapper.querySelector(".saveSettings").innerText));
                        break;
                    }
                }
                /* save new settings item if an existing one was not found */
                if ( bFoundExisting === false ){
                    picker.ajax.captureSettings(_spPageContextInfo.serverRequestPath, fieldName, JSON.parse(popupWrapper.querySelector(".saveSettings").innerText));
                }
            });
            
            configPopup.querySelector(".removeMapping").addEventListener("click",function(e){
                picker.log("remove");
                picker.log(e);
                picker.log(this);
                var parentRow = picker.parentsUntilMatchingSelector(this, ".mappingRow");
                picker.log(parentRow);
                parentRow.remove();
            });

            function fxAddNew(e){
                picker.log("AddNew button clicked");
                picker.log(e);
                picker.log(this);
                var newRow = picker.parentsUntilMatchingSelector(this, ".mappingRow:last-of-type").cloneNode(true);
                //newRow.querySelector("SELECT.listField OPTION[value='0']").setAttribute("selected","selected");
                newRow.querySelector("SELECT.listField").value = 0;
                //newRow.querySelector("SELECT.userProperty OPTION[value='0']").setAttribute("selected","selected");
                newRow.querySelector("INPUT.userProperty").value = "";
                picker.parentsUntilMatchingSelector(this, "[id*='_configMappingRowsWrapper']").appendChild(newRow);
                //configPopup.querySelector(".mappingRow:last-of-type .addNew").addEventListener("click",fxAddNew);
                picker.parentsUntilMatchingSelector(this, "[id*='_configMappingRowsWrapper']").querySelector(".mappingRow:last-of-type .addNew").addEventListener("click",fxAddNew);
                picker.parentsUntilMatchingSelector(this, "[id*='_configMappingRowsWrapper']").querySelector(".mappingRow:last-of-type .removeMapping").addEventListener("click",function(e){
                    picker.log("remove");
                    picker.log(e);
                    picker.log(this);
                    var parentRow = picker.parentsUntilMatchingSelector(this, ".mappingRow");
                    picker.log(parentRow);
                    parentRow.remove();
                });
            };
            
            
            configPopup.querySelector(".addNew").addEventListener("click",fxAddNew);


        }
        //picker.populateCurrentSettings();
    },
    appendConfigStyles: function(){
        var arrStyles = [".pickerConfigPopup {display: none; background-color: #e6e6e6; border: 1px solid black; padding: 5px 10px; margin-top: 3px;}"];
        arrStyles.push(".configPopupHeader {text-align: right;}");
        arrStyles.push(".configPopupFooter {margin-top: 3px;}");
        arrStyles.push(".buttonish {cursor: pointer; border: 1px solid black; padding: 1px 10px;}");
        arrStyles.push(".closeButton {font-size: 2.3em; border: 0px hidden transparent; padding: 0px 0px;}");
        arrStyles.push(".addNew {background-color: #40588e; color: #ffeee7; display: none;}");
        arrStyles.push(".removeMapping {background-color: #691818; color: #f1f1f1;}")
        arrStyles.push(".saveButton {background-color: #4e864d; color: #f1f1f1;}");
        arrStyles.push(".saveSettings {border: 3px double black; font-family: monospace; padding: 5px; margin-top: 8px;}");
        arrStyles.push(".pickerConfigButton {background-color: #e6e6a6; color: brown; border: 0px hidden transparent; padding-top: 3px; padding-bottom: 3px;}");
        arrStyles.push("[id*='_configMappingRowsWrapper'] > tr.mappingRow:last-of-type .addNew {display: inline-block;}");
        arrStyles.push("INPUT.userProperty {width: 280px;}");
        var elemStyles = document.createElement("STYLE");
        elemStyles.type = "text/css";
        elemStyles.innerHTML = arrStyles.join("");
        document.head.appendChild(elemStyles);
    },
    parentsUntilMatchingSelector: function(htmlElement, selector){
        var parent = htmlElement.parentElement;
        var lastParent = parent;
        var iBreak = 0;
        while ( iBreak < 100 && parent.querySelectorAll(selector).length < 1 ){
            lastParent = parent;
            parent = parent.parentElement;
            iBreak++;
        }
        if ( parent.querySelectorAll(selector).length > 0 ){
            return lastParent;
        }
        else {
            return false;
        }
    },
    setupListeners: function(){
        for ( var iL = 0; iL < pickerListeners.length; iL++ ){
            var listener = pickerListeners[iL];
            var pickerFormField = picker.findFormField("Title", listener.fieldDisplayName);
            if ( !pickerFormField === false ){
                var pickerTopLevelElementId = pickerFormField.TopLevelElementId;
                var pickerInstance = picker.instances[pickerTopLevelElementId];
                if ( listener.bCaptureProfileDetails === true ) {
                    pickerInstance.spcsom.OnUserResolvedClientScript = function(pickerId, arrUsers) {
                        picker.log(pickerId);
                        picker.log(arrUsers);
                        if ( pickerInstance.spcsom.HasResolvedUsers() === true ) {
                            for ( var iUser = 0; iUser < arrUsers.length; iUser++ ) {
                                picker.log(arrUsers[iUser]);
                                picker.getUserPropertiesByLoginName(arrUsers[iUser].Key, function(){
                                    for ( var iMapping = 0; iMapping < listener.arrMapping.length; iMapping++ ){
                                        var setField = picker.findFormField("Name", listener.arrMapping[iMapping].fin);
                                        picker.log(setField);
                                        var setFieldElement = document.querySelector("[id*='"+ setField.Id +"']");
                                        picker.log(setFieldElement);
                                        picker.log(listener.arrMapping[iMapping].userProperty);
                                        var setUserProperty = undefined, splitUserProperty = undefined, setUserSubProperty = undefined;
                                        try{
                                            if ( listener.arrMapping[iMapping].userProperty.indexOf("|") >= 0 ) {
                                                splitUserProperty = listener.arrMapping[iMapping].userProperty.split("|");
                                                setUserProperty = eval(splitUserProperty[0]);
                                                setUserSubProperty = splitUserProperty[1];
                                            }
                                            else{
                                                setUserProperty = eval(listener.arrMapping[iMapping].userProperty);
                                            }
                                        }catch(err){}
                                        try{
                                            if ( typeof(setUserProperty) === "object" ) {
                                                if ( typeof(setUserProperty.length) !== "undefined" && setUserProperty.length > 0 ) {
                                                    picker.log("Setting userproperty mapping by looping through the referenced JS array",true);
                                                    picker.log(setUserProperty,true);
                                                    for ( var iG = 0; iG < setUserProperty.length; iG++ ){
                                                        if ( typeof(setUserSubProperty) !== "undefined" ){
                                                            picker.log("Setting userproperty mapping by looping through sub-array",true);
                                                            var subLoop = eval("setUserProperty["+iG+"]."+ setUserSubProperty);
                                                            picker.log(subLoop,true);
                                                            for ( var iU = 0; iU < subLoop.length; iU++ ){
                                                                picker.log("picker adding person or group to field with fin |"+ listener.arrMapping[iMapping].fin +"|");
                                                                SPClientPeoplePicker.SPClientPeoplePickerDict[picker.findFormField("Name", listener.arrMapping[iMapping].fin).TopLevelElementId].AddUserKeys(subLoop[iU].LoginName);
                                                                //pickerInstance.spcsom.AddUserKeys(setUserProperty[iG].LoginName);
                                                            }

                                                        }
                                                        else {
                                                            picker.log("picker adding person or group to field with fin |"+ listener.arrMapping[iMapping].fin +"|");
                                                            SPClientPeoplePicker.SPClientPeoplePickerDict[picker.findFormField("Name", listener.arrMapping[iMapping].fin).TopLevelElementId].AddUserKeys(setUserProperty[iG].LoginName);
                                                            //pickerInstance.spcsom.AddUserKeys(setUserProperty[iG].LoginName);
                                                        }
                                                    }
                                                }
                                            }
                                        }catch(err){}
                                        if ( typeof(setUserProperty) !== "object" ) {
                                            picker.log("picker capturing simple user profile property value in form field with fin |"+ listener.arrMapping[iMapping].fin +"|");
                                            if ( setFieldElement.tagName === "INPUT" ) {
                                                setFieldElement.value = eval(listener.arrMapping[iMapping].userProperty);
                                            }
                                            else if ( setFieldElement.tagName === "SELECT" ) {
                                                if ( !setFieldElement.getAttribute("multiple") === true ) {
                                                    /* single value select control */
                                                    setFieldElement.value = eval(listener.arrMapping[iMapping].userProperty);
                                                }
                                                else {
                                                    /* multivalue select control */
                                                }
                                            }
                                            else if ( setFieldElement.tagName === "TEXTAREA" ) {
                                                setFieldElement.value = eval(listener.arrMapping[iMapping].userProperty);
                                            }
                                        }
                                    }    
                                });
                            }
                        }
                    }
                }
            }
        }
    },
    populateCurrentSettings: function(){
        for ( var iPL = 0; iPL < pickerListeners.length; iPL++ ){
            var listener = pickerListeners[iPL];
            picker.log(listener,true);
            var pickerFormField = picker.findFormField("Title", listener.fieldDisplayName);
            if ( !pickerFormField === false ){
                var pickerTopLevelElementId = pickerFormField.TopLevelElementId;
                var pickerInstance = picker.instances[pickerTopLevelElementId];
                var pickerConfigWrapper = document.getElementById(pickerTopLevelElementId +"_configWrapper");
                if ( listener.bCaptureProfileDetails === true ) {
                    var pickerConfigMappingTBODY = document.getElementById(pickerTopLevelElementId +"_configMappingRowsWrapper");
                    var mappingRow = pickerConfigMappingTBODY.querySelector(".mappingRow");
                    for ( var iMapping = 0; iMapping < listener.arrMapping.length; iMapping++ ){
                        picker.log(listener.arrMapping[iMapping],true);
                        var setField = picker.findFormField("Name", listener.arrMapping[iMapping].fin);
                        picker.log(setField);
                        var setFieldElement = document.querySelector("[id*='"+ setField.Id +"']");
                        picker.log(setFieldElement);
                        var userPropertyName = listener.arrMapping[iMapping].userProperty;
                        picker.log(userPropertyName);
                        /* set the values of our select controls that define our mapping */
                        if ( iMapping > 0 ) {
                            pickerConfigMappingTBODY.querySelector(".mappingRow:last-of-type .addNew").click();
                            var newRow = pickerConfigMappingTBODY.querySelector(".mappingRow:last-of-type");
                            try{newRow.querySelector("SELECT.listField OPTION[value='"+ setField.Id +"']").setAttribute("selected","selected");}catch(err){}
                            //try{newRow.querySelector("SELECT.userProperty OPTION[value='"+ userPropertyName +"']").setAttribute("selected","selected");}catch(err){}
                            try{newRow.querySelector("INPUT.userProperty").value = userPropertyName;}catch(err){}
                        }
                        else {
                            try{mappingRow.querySelector("SELECT.listField OPTION[value='"+ setField.Id +"']").setAttribute("selected","selected");}catch(err){}
                            //try{mappingRow.querySelector("SELECT.userProperty OPTION[value='"+ userPropertyName +"']").setAttribute("selected","selected");}catch(err){}
                            try{mappingRow.querySelector("INPUT.userProperty").value = userPropertyName;}catch(err){}
                        }
                    }
                }
            }
        }
    },
    showAvailableProperties: function(){
        var propsWrapper = document.createElement("ul");
        propsWrapper.setAttribute("id","propsWrapper");
        for ( prop in picker.userProperties ) {
            picker.log("["+ prop +"] = |" + JSON.stringify(picker.userProperties[prop]) +"|",true);
            var propOut = document.createElement("li");
            propOut.setAttribute("id","propOut"+prop);
            propOut.innerText = "["+ prop +"] = |" + JSON.stringify(picker.userProperties[prop]) +"|";
            propsWrapper.appendChild(propOut);
        }
        document.getElementById("DeltaPlaceHolderMain").insertAdjacentElement('beforebegin',propsWrapper);
        propsWrapper.addEventListener("click",function(e){
            this.remove();
        });
    },
    init: setTimeout(function(){
        picker.getUserProperties(function(){
            picker.getPageWebPartsAndFormFields(function(wp){
                picker.log("webpart detected");
                picker.log(wp);
            }, function(){
                picker.getAllPickers();
                picker.appendConfigStyles();
            });
            if ( picker.isPageInEditMode() === false ) {
                picker.showConfigButtons();
                picker.ajax.getSettings(function(){
                    picker.setupListeners();
                    picker.populateCurrentSettings();
                });
            }
        });
    },23)
}