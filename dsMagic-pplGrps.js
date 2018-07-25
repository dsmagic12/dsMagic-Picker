/*http://expertsoverlunch.com/sandbox/Lists/Tasks/NewForm.aspx*/
/*http://expertsoverlunch.com/sandbox/Lists/PeopleAndGroupPickers/NewForm.aspx*/
/*fails on calendar list: http://expertsoverlunch.com/sandbox/Lists/Meetings/NewForm.aspx */
var pplGrpsListeners = pplGrpsListeners || [];
var SPFieldTypes = {
    listFields: [],
    listFieldsByFIN: {},
    listFieldsByDisplay: {},
    listFieldsByGuid: {},
    listFieldsExpanded: [],
    ajax: {
        lastCall: {},
        create: function (restURL, object, fxCallback, fxFailed) {
            SPFieldTypes.ajax.lastCall = {
                xhr: null,
                readyState: null,
                data: null,
                status: null,
                url: restURL,
                error: null
            };
            var xhr = new XMLHttpRequest();
            xhr.open('POST', restURL, true);
            xhr.setRequestHeader("X-RequestDigest", document.getElementById("__REQUESTDIGEST").value);
            xhr.setRequestHeader("IF-MATCH", "*");
            xhr.setRequestHeader("X-HTTP-Method", "POST");
            xhr.setRequestHeader("accept", "application/json;odata=verbose");
            xhr.setRequestHeader("content-type", "application/json;odata=verbose");
            xhr.onreadystatechange = function () {
                SPFieldTypes.ajax.lastCall.readyState = xhr.readyState;
                SPFieldTypes.ajax.lastCall.status = xhr.status;
                if (xhr.readyState === 4) {
                    if (xhr.status !== 201) {
                        SPFieldTypes.ajax.lastCall.data = xhr.response;
                        if (typeof (fxFailed) === "function") {
                            fxFailed(xhr, xhr.response, xhr.status);
                        }
                    } 
                    else {
                        SPFieldTypes.ajax.lastCall.data = xhr.response;
                        if (typeof (fxCallback) === "function") {
                            fxCallback(xhr, xhr.response);
                        }
                    }
                }
            };
            xhr.send(object);
        },
        update: function (restURL, object, fxCallback, fxFailed) {
            SPFieldTypes.ajax.lastCall = {
                xhr: null,
                readyState: null,
                data: null,
                status: null,
                url: restURL,
                error: null
            };
            var xhr = new XMLHttpRequest();
            xhr.open('POST', restURL, true);
            xhr.setRequestHeader("X-RequestDigest", document.getElementById("__REQUESTDIGEST").value);
            xhr.setRequestHeader("IF-MATCH", "*");
            xhr.setRequestHeader("X-HTTP-Method", "MERGE");
            xhr.setRequestHeader("accept", "application/json;odata=verbose");
            xhr.setRequestHeader("content-type", "application/json;odata=verbose");
            xhr.onreadystatechange = function () {
                SPFieldTypes.ajax.lastCall.readyState = xhr.readyState;
                SPFieldTypes.ajax.lastCall.status = xhr.status;
                if (xhr.readyState === 4) {
                    if (xhr.status !== 200 && xhr.status !== 204) {
                        SPFieldTypes.ajax.lastCall.data = xhr.response;
                        if (typeof (fxFailed) === "function") {
                            fxFailed(xhr, xhr.response, xhr.status);
                        }
                    } 
                    else {
                        var resp = xhr.response;
                        SPFieldTypes.ajax.lastCall.data = resp;
                        if (typeof (fxCallback) === "function") {
                            fxCallback(xhr, resp);
                        }
                    }
                }
            };
            xhr.send(object);
        },
        delete: function (restURL, object, fxCallback) {
            SPFieldTypes.ajax.lastCall = {
                xhr: null,
                readyState: null,
                data: null,
                status: null,
                url: restURL,
                error: null
            };
            var xhr = new XMLHttpRequest();
            xhr.open('POST', restURL, true);
            xhr.setRequestHeader("X-RequestDigest", document.getElementById("__REQUESTDIGEST").value);
            xhr.setRequestHeader("IF-MATCH", "*");
            xhr.setRequestHeader("X-HTTP-Method", "DELETE");
            xhr.setRequestHeader("accept", "application/json;odata=verbose");
            xhr.setRequestHeader("content-type", "application/json;odata=verbose");
            xhr.onreadystatechange = function () {
                SPFieldTypes.ajax.lastCall.readyState = xhr.readyState;
                SPFieldTypes.ajax.lastCall.status = xhr.status;
                if (xhr.readyState === 4) {
                    if (xhr.status !== 200) {
                        SPFieldTypes.ajax.lastCall.data = xhr.response;
                        if (typeof (fxFailed) === "function") {
                            fxFailed(xhr, xhr.response, xhr.status);
                        }
                    } 
                    else {
                        var resp = xhr.response;
                        SPFieldTypes.ajax.lastCall.data = resp;
                        if (typeof (fxCallback) === "function") {
                            fxCallback(xhr, resp);
                        }
                    }
                }
            };
            xhr.send(object);
        },
        read: function (restURL, fxCallback, fxLastPage, fxFailed, sDataType) {
            if (typeof (sDataType) === "undefined") {
                var sDataType = "json";
            }
            SPFieldTypes.ajax.lastCall = {
                xhr: null,
                readyState: null,
                data: null,
                status: null,
                url: restURL,
                error: null
            };
            var xhr = new XMLHttpRequest();
            xhr.open('GET', restURL, true);
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
                    SPFieldTypes.log("SPFieldTypes.ajax.read... unknown sDataType value |" + sDataType + "|", true);
            }
            xhr.onreadystatechange = function () {
                SPFieldTypes.ajax.lastCall.readyState = xhr.readyState;
                SPFieldTypes.ajax.lastCall.status = xhr.status;
                if (xhr.readyState === 4) {
                    if (xhr.status !== 200) {
                        SPFieldTypes.ajax.lastCall.data = xhr.response;
                        if (typeof (fxFailed) === "function") {
                            fxFailed(xhr, xhr.response, xhr.status);
                        }
                    } 
                    else {
                        switch (sDataType) {
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
                        SPFieldTypes.ajax.lastCall.data = resp;
                        if (typeof (fxCallback) === "function") {
                            fxCallback(xhr, resp);
                        }
                        if (typeof (resp.d.__next) !== "undefined") {
                            SPFieldTypes.ajax.read(resp.d.__next, fxCallback, fxLastPage);
                        } else if (typeof (fxLastPage) === "function") {
                            fxLastPage(xhr, resp);
                        }
                    }
                }
            };
            SPFieldTypes.ajax.lastCall.xhr = xhr;
            xhr.send();
        }
    },
    log: function(message){
        try{
            console.log(message);
        }
        catch(er){
        }
    },
    fromFormBody: function(elemMsFormBody){
        /*SPFieldTypes.fromFormBody(document.getElementsByClassName("ms-formbody")[0])*/
        var firstComment = elemMsFormBody.innerHTML.substr(0,elemMsFormBody.innerHTML.indexOf("-->"));
        SPFieldTypes.log(firstComment,true);
        var FieldName = firstComment.substr(firstComment.indexOf("FieldName=")+11,200);
        FieldName = FieldName.substr(0,FieldName.indexOf("\n")-1);
        var FieldInternalName = firstComment.substr(firstComment.indexOf("FieldInternalName=")+19,200);
        FieldInternalName = FieldInternalName.substr(0,FieldInternalName.indexOf("\n")-1);
        var FieldType = firstComment.substr(firstComment.indexOf("FieldType=")+11,200);
        FieldType = FieldType.substr(0,FieldType.indexOf("\n")-1);
        SPFieldTypes.log(FieldName +"|"+ FieldInternalName +"|"+ FieldType,true);
        var fieldGUID = "";
        try{
            fieldGUID = WPQ2FormCtx.ListSchema[FieldInternalName].Id;
        }
        catch(err){
            fieldGUID = SPFieldTypes.listFields.filter(function(field){ return field.StaticName === FieldInternalName; })[0].Id;
        }
        switch (FieldType){
            case "SPFieldDateTime":
                if ( SPFieldTypes.listFieldsByFIN[FieldInternalName].DisplayFormat === 0 ) {
                    var oSPField = new SPFieldTypes[FieldType].dateOnly(FieldInternalName, fieldGUID, FieldName);
                }
                else if ( SPFieldTypes.listFieldsByFIN[FieldInternalName].DisplayFormat === 1 ) {
                    var oSPField = new SPFieldTypes[FieldType].dateAndTime(FieldInternalName, fieldGUID, FieldName);
                }                
                break;
            case "SPFieldNote":
                if ( SPFieldTypes.listFieldsByFIN[FieldInternalName].RichText === true ) {
                    var oSPField = new SPFieldTypes[FieldType].rich(FieldInternalName, fieldGUID, FieldName);
                }
                else if ( SPFieldTypes.listFieldsByFIN[FieldInternalName].RichText === false ) {
                    var oSPField = new SPFieldTypes[FieldType].plain(FieldInternalName, fieldGUID, FieldName);
                }
                break;
            case "SPFieldChoice":
                if ( SPFieldTypes.listFieldsByFIN[FieldInternalName].EditFormat === 1 ) {
                    var oSPField = new SPFieldTypes[FieldType].radio(FieldInternalName, fieldGUID, FieldName);
                }
                else if ( SPFieldTypes.listFieldsByFIN[FieldInternalName].EditFormat === 0 ) {
                    var oSPField = new SPFieldTypes[FieldType].dropdown(FieldInternalName, fieldGUID, FieldName);
                }
                break;
            case "SPFieldUser":
                if ( SPFieldTypes.listFieldsByFIN[FieldInternalName].SelectionMode === 1 ) {
                    var oSPField = new SPFieldTypes[FieldType].peopleOrGroups(FieldInternalName, fieldGUID, FieldName);
                }
                else if ( SPFieldTypes.listFieldsByFIN[FieldInternalName].SelectionMode === 0 ) {
                    var oSPField = new SPFieldTypes[FieldType].peopleOnly(FieldInternalName, fieldGUID, FieldName);
                }
                break;
            case "SPFieldUserMulti":
                if ( SPFieldTypes.listFieldsByFIN[FieldInternalName].SelectionMode === 1 ) {
                    var oSPField = new SPFieldTypes[FieldType].peopleOrGroups(FieldInternalName, fieldGUID, FieldName);
                }
                else if ( SPFieldTypes.listFieldsByFIN[FieldInternalName].SelectionMode === 0 ) {
                    var oSPField = new SPFieldTypes[FieldType].peopleOnly(FieldInternalName, fieldGUID, FieldName);
                }
                break;
            case "SPFieldUrl":
                var oSPField = new SPFieldTypes[FieldType].hyperlink(FieldInternalName, fieldGUID, FieldName);
                break;
            default:
                var oSPField = new SPFieldTypes[FieldType](FieldInternalName, fieldGUID, FieldName);
                oSPField.elemMain = oSPField.object("control");
        }
        
        /* add our field's rest definition to its FormCtxDefinition */
        try{
            oSPField.restDefinition = SPFieldTypes.listFieldsByFIN[FieldInternalName];
        }catch(err){
            oSPField.restDefinition = null;
        }
        /* add our field's FormCtxDefinition to its basic definition and all of its rest definitions */
        try{
            oSPField.FormCtxDefinition = WPQ2FormCtx.ListSchema[FieldInternalName];
            SPFieldTypes.listFieldsByFIN[FieldInternalName].oSPField = oSPField;
            SPFieldTypes.listFieldsByDisplay[FieldName].oSPField = oSPField;
            SPFieldTypes.listFieldsByGuid[fieldGUID].oSPField = oSPField;
        }catch(err){
            oSPField.FormCtxDefinition = null;
            SPFieldTypes.listFieldsByFIN[FieldInternalName].oSPField = oSPField;
            SPFieldTypes.listFieldsByDisplay[FieldName].oSPField = oSPField;
            SPFieldTypes.listFieldsByGuid[fieldGUID].oSPField = oSPField;
        }
        SPFieldTypes.log(oSPField,true);
        SPFieldTypes.listFieldsExpanded.push(oSPField);
    },
    getListFields: function(listGUID, fxCallback){
        if ( typeof(listGUID) === "undefined" ) { var listGUID = _spPageContextInfo.pageListId.replace("{","").replace("}",""); }
        SPFieldTypes.ajax.read(
            _spPageContextInfo.webAbsoluteUrl+"/_api/web/lists(guid'"+ listGUID +"')/Fields",
            function(x,d){
                for ( var iR = 0; iR < d.d.results.length; iR++ ){
                    var result = d.d.results[iR];
                    result.listGUID = listGUID;
                    SPFieldTypes.listFields.push(result);
                    SPFieldTypes.listFieldsByFIN[result.StaticName] = result;
                    SPFieldTypes.listFieldsByDisplay[result.Title] = result;
                    SPFieldTypes.listFieldsByGuid[result.Id] = result;
                }
            },
            function(xL,dL){
                if ( typeof(fxCallback) === 'function' ) { 
                    fxCallback(SPFieldTypes.listFields); 
                }
            },
            function(){
                SPFieldTypes.log("Failed");
            },
            "json"
        );
    },
    /*http://expertsoverlunch.com/sandbox/Lists/Field%20Types%20Demo/NewForm.aspx*/
    SPFieldText: function(fin, guid, displayName){
        return {
            fin: fin,
            guid: guid,
            fieldDisplayName: displayName,
            fieldType: "SPFieldText",
            fieldSubType: "",
            selectors: {
                control: "INPUT[type='text'][id='"+ fin +"_"+ guid +"_$TextField'], INPUT[type='text'][id$='_TextField'][title^='"+ displayName +"']"
            },
            elemMain: null,
            get: function(){
                return this.elemMain.value
            },
            set: function(newValue){
                this.elemMain.value = newValue;
            },
            object: function(selName){
                return document.querySelector(this.selectors[selName])
            }
        }
    },
    SPFieldDateTime: {
        dateOnly: function(fin,guid,displayName){
            return {
                fin: fin,
                guid: guid,
                fieldDisplayName: displayName,
                fieldType: "SPFieldDateTime",
                fieldSubType: "DateOnly",
                selectors: {
                    control: "INPUT[id='"+ fin +"_"+ guid +"_$DateTimeFieldDate']",
                    picker: "IMG[id='"+ fin +"_"+ guid +"_$DateTimeFieldDateDatePickerImage']"
                },
                elemMain: document.querySelector(this.selectors.control),
                elemPicker: document.querySelector(this.selectors.control).parentElement,
                get: function(){
                    return this.elemMain.value
                },
                set: function(newValue){
                    this.elemMain.value = newValue;
                }
            }
        },
        dateAndTime: function(fin,guid,displayName){
            return {
                fin: fin,
                guid: guid,
                fieldDisplayName: displayName,
                fieldType: "SPFieldDateTime",
                fieldSubType: "DateAndTime",
                selectors: {
                    control: "INPUT[id='"+ fin +"_"+ guid +"_$DateTimeFieldDate']",
                    picker: "IMG[id='"+ fin +"_"+ guid +"_$DateTimeFieldDateDatePickerImage']",
                    hours: "SELECT[id='"+ fin +"_"+ guid +"_$DateTimeFieldDateHours']",
                    minutes: "SELECT[id='"+ fin +"_"+ guid +"_$DateTimeFieldDateMinutes']"
                },
                elemMain: document.querySelector(this.selectors.control),
                elemPicker: document.querySelector(this.selectors.control).parentElement,
                elemHours: document.querySelector(this.selectors.hours),
                elemMinutes: document.querySelector(this.selectors.minutes),
                get: function(){
                    return this.elemMain.value
                },
                set: function(newValue){
                    this.elemMain.value = newValue;
                }
            }
        }
    },
    SPFieldNumber: function(fin, guid,displayName){
        return {
            fin: fin,
            guid: guid,
            fieldDisplayName: displayName,
            fieldType: "SPFieldNumber",
            fieldSubType: "",
            selectors: {
                control: "INPUT[id='"+ fin +"_"+ guid +"_$NumberField']"
            },
            elemMain: document.querySelector(this.selectors.control),
            get: function(){
                return this.elemMain.value
            },
            set: function(newValue){
                this.elemMain.value = newValue;
            }
        }
    },
    SPFieldCurrency: function(fin,guid,displayName){
        return {
            fin: fin,
            guid: guid,
            fieldDisplayName: displayName,
            fieldType: "SPFieldCurrency",
            fieldSubType: "",
            selectors: {
                control: "INPUT[id='"+ fin +"_"+ guid +"_$CurrencyField']"
            },
            elemMain: document.querySelector(this.selectors.control),
            get: function(){
                return this.elemMain.value
            },
            set: function(newValue){
                this.elemMain.value = newValue;
            }
        }
    },
    SPFieldNote: {
        plain: function(fin,guid,displayName){
            return {
                fin: fin,
                guid: guid,
                fieldDisplayName: displayName,
                fieldType: "SPFieldNote",
                fieldSubType: "Plain",
                selectors: {
                    control: "TEXTAREA[id='"+ fin +"_"+ guid +"_$TextField']"
                },
                elemMain: document.querySelector(this.selectors.control),
                get: function(){
                    return this.elemMain.value
                },
                set: function(newValue){
                    this.elemMain.value = newValue;
                }
            }
        },
        rich: function(fin,guid,displayName){
            return {
                fin: fin,
                guid: guid,
                fieldDisplayName: displayName,
                fieldType: "SPFieldNote",
                fieldSubType: "Rich",
                selectors: {
                    control: "DIV[id='"+ fin +"_"+ guid +"_$TextField_inplacerte']"
                },
                elemMain: document.querySelector(this.selectors.control),
                get: function(){
                    return this.elemMain.innerHTML
                },
                set: function(newValue){
                    this.elemMain.innerHTML = newValue;
                }
            }
        }
    },
    SPFieldChoice: {
        dropdown: function(fin,guid,displayName){
            return {
                fin: fin,
                guid: guid,
                fieldDisplayName: displayName,
                fieldType: "SPFieldChoice",
                fieldSubType: "Dropdown",
                selectors: {
                    control: "SELECT[id='"+ fin +"_"+ guid +"_$DropDownChoice']",
                    choices: "SELECT[id='"+ fin +"_"+ guid +"_$DropDownChoice'] OPTION",
                    existingValueIndicator: "INPUT[id='"+ fin +"_"+ guid +"_DropDownButton",
                    fillInIndicator: "INPUT[id='"+ fin +"_"+ guid +"_FillInButton",
                    fillIn: "INPUT[id='"+ fin +"_"+ guid +"_$FillInChoice"
                },
                elemMain: document.querySelector(this.selectors.control),
                get: function(){
                    return this.elemMain.value
                },
                set: function(newValue){
                    this.elemMain.value = newValue;
                }
            }
        },
        radio: function(fin,guid,displayName){
            return {
                fin: fin,
                guid: guid,
                fieldDisplayName: displayName,
                fieldType: "SPFieldChoice",
                fieldSubType: "Radio",
                selectors: {
                    control: "TABLE[id='"+ fin +"_"+ guid +"_ChoiceRadioTable']",
                    choices: "INPUT[name='"+ fin +"_"+ guid +"_$RadioButtonChoiceField']",
                    fillInIndicator: "INPUT[id='"+ fin +"_"+ guid +"_$RadioButtonChoiceFieldFillInRadio",
                    fillIn: "INPUT[id='"+ fin +"_"+ guid +"_$RadioButtonChoiceFieldFillInText"
                },
                elemMain: document.querySelector(this.selectors.control),
                get: function(){
                    return this.elemMain.value
                },
                set: function(newValue){
                    this.elemMain.value = newValue;
                }
            }
        }
    },
    SPFieldMultiChoice: function(fin,guid,displayName){
        return {
            fin: fin,
            guid: guid,
            fieldDisplayName: displayName,
            fieldType: "SPFieldMultiChoice",
            fieldSubType: "",
            selectors: {
                control: "TABLE[id='"+ fin +"_"+ guid +"_MultiChoiceTable']",
                choices: "TABLE[id='"+ fin +"_"+ guid +"_MultiChoiceTable'] INPUT[type='checkbox'][id^='"+ fin +"_"+ guid +"_MultiChoiceOption_']",
                fillInIndicator: "INPUT[id='"+ fin +"_"+ guid +"FillInRadio",
                fillIn: "INPUT[id='"+ fin +"_"+ guid +"FillInText"
            },
            elemMain: document.querySelector(this.selectors.control),
            get: function(){
                return this.elemMain.value
            },
            set: function(newValue){
                this.elemMain.value = newValue;
            }
        }
    },
    SPFieldLookup: function(fin,guid,displayName){
        return {
            fin: fin,
            guid: guid,
            fieldDisplayName: displayName,
            fieldType: "SPFieldLookup",
            fieldSubType: "",
            selectors: {
                control: "SELECT[id='"+ fin +"_"+ guid +"_$LookupField']"
            },
            elemMain: document.querySelector(this.selectors.control),
            get: function(){
                return this.elemMain.value
            },
            set: function(newValue){
                this.elemMain.value = newValue;
            }
        }
    },
    SPFieldLookupMulti: function(fin,guid,displayName){
        return {
            fin: fin,
            guid: guid,
            fieldDisplayName: displayName,
            fieldType: "SPFieldLookupMulti",
            fieldSubType: "",
            selectors: {
                possible: "SELECT[id='"+ fin +"_"+ guid +"_SelectCandidate']",
                selected: "SELECT[id='"+ fin +"_"+ guid +"_SelectResult']",
                addButton: "SELECT[id='"+ fin +"_"+ guid +"_AddButton']",
                removeButton: "SELECT[id='"+ fin +"_"+ guid +"_RemoveButton']"
            },
            elemPossible: document.querySelector(this.selectors.possible),
            elemSelected: document.querySelector(this.selectors.selected),
            get: function(){
                return this.elemSelected.childNodes;
            },
            set: function(arrNewValue){
                for ( var iV = 0; iV < arrNewValue.length; iV++ ){
                    var lookupOption = document.querySelector(this.selectors.possible +" OPTION[value='"+ arrNewValue[iV] +"']");
                    pplGrps.log("need to move this to selected",true);
                    pplGrps.log(lookupOption,true);
                }
            }
        }
    },
    SPFieldBoolean: function(fin,guid,displayName){
        return {
            fin: fin,
            guid: guid,
            fieldDisplayName: displayName,
            fieldType: "SPFieldBoolean",
            fieldSubType: "",
            selectors: {
                control: "INPUT[id='"+ fin +"_"+ guid +"_$BooleanField']"
            },
            elemMain: document.querySelector(this.selectors.control),
            get: function(){
                return this.elemMain.getAttribute("checked")
            },
            set: function(newValue){
                if ( newValue === true ){
                    this.elemMain.setAttribute("checked","checked");
                }
                else {
                    this.elemMain.removeAttribute("checked");
                }
            }
        }
    },
    SPFieldAllDayEvent: function(fin,guid,displayName){
        return {
            fin: fin,
            guid: guid,
            fieldDisplayName: displayName,
            fieldType: "SPFieldAllDayEvent",
            fieldSubType: "",
            selectors: {
                control: "INPUT[id='"+ fin +"_"+ guid +"_AllDayEventField']"
            },
            elemMain: document.querySelector(this.selectors.control),
            get: function(){
                return this.elemMain.getAttribute("checked")
            },
            set: function(newValue){
                if ( newValue === true ){
                    this.elemMain.setAttribute("checked","checked");
                }
                else {
                    this.elemMain.removeAttribute("checked");
                }
            }
        }
    },
    SPFieldRecurrence: function(fin,guid,displayName){
        return {
            fin: fin,
            guid: guid,
            fieldDisplayName: displayName,
            fieldType: "SPFieldRecurrence",
            fieldSubType: "",
            selectors: {
                control: "INPUT[id='"+ fin +"_"+ guid +"_RecurrenceField']",
                recurrenceTypeDaily: "INPUT[id*='_RecurrenceDataField_'][id$='recurrencePatternType2']",
                recurrenceTypeWeekly: "INPUT[id*='_RecurrenceDataField_'][id$='recurrencePatternType3']",
                recurrenceTypeMonthly: "INPUT[id*='_RecurrenceDataField_'][id$='recurrencePatternType4']",
                recurrenceTypeYearly: "INPUT[id*='_RecurrenceDataField_'][id$='recurrencePatternType5']",
                dailyPatternXDays: "INPUT[id*='_RecurrenceDataField_'][id$='dailyRecurType0']",
                dailyPatternXDaysValue: "INPUT[id*='_RecurrenceDataField_'][id$='daily_dayFrequency']",
                dailyPatternEveryWeekday: "INPUT[id*='_RecurrenceDataField_'][id$='dailyRecurType1']",
                weeklyWeekFrequencyValue: "INPUT[id*='_RecurrenceDataField_'][id$='weekly_weekFrequency']",
                weeklyOnSunday: "INPUT[id*='_RecurrenceDataField_'][id$='weekly_multiDays_0']",
                weeklyOnMondy: "INPUT[id*='_RecurrenceDataField_'][id$='weekly_multiDays_1']",
                weeklyOnTuesday: "INPUT[id*='_RecurrenceDataField_'][id$='weekly_multiDays_2']",
                weeklyOnWednesday: "INPUT[id*='_RecurrenceDataField_'][id$='weekly_multiDays_3']",
                weeklyOnThursday: "INPUT[id*='_RecurrenceDataField_'][id$='weekly_multiDays_4']",
                weeklyOnFriday: "INPUT[id*='_RecurrenceDataField_'][id$='weekly_multiDays_5']",
                weeklyOnSaturday: "INPUT[id*='_RecurrenceDataField_'][id$='weekly_multiDays_6']",
                monthlyDayXMonth: "INPUT[id*='_RecurrenceDataField_'][id$='monthlyRecurType0']",
                monthlyDayXMonthDayValue: "INPUT[id*='_RecurrenceDataField_'][id$='monthly_day']",
                monthlyDayXMonthMonthValue: "INPUT[id*='_RecurrenceDataField_'][id$='monthly_monthFrequency']",
                monthlyByDay: "INPUT[id*='_RecurrenceDataField_'][id$='monthlyRecurType1']",
                monthlyByDayWeekOfMonthValue: "SELECT[id*='_RecurrenceDataField_'][id$='monthlyByDay_weekOfMonth']",
                monthlyByDayDayValue: "SELECT[id*='_RecurrenceDataField_'][id$='monthlyByDay_day']",
                monthlyByDayFrequencyValue: "INPUT[id*='_RecurrenceDataField_'][id$='monthlyByDay_monthFrequency']",
                yearly: "SELECT[id*='_RecurrenceDataField_'][id$='yearlyRecurType0']",
                yearlyMonth: "SELECT[id*='_RecurrenceDataField_'][id$='yearly_month']",
                yearlyDay: "SELECT[id*='_RecurrenceDataField_'][id$='yearly_day']",
                yearlyByDay: "SELECT[id*='_RecurrenceDataField_'][id$='yearlyRecurType1']",
                yearlyByDayWeekOfMonthValue: "SELECT[id*='_RecurrenceDataField_'][id$='yearlyByDay_weekOfMonth']",
                yearlyByDayDayValue: "SELECT[id*='_RecurrenceDataField_'][id$='yearlyByDay_day']",
                yearlyByDayMonthValue: "SELECT[id*='_RecurrenceDataField_'][id$='yearlyByDay_month']",
                recurrenceStartDate: "INPUT[id*='_RecurrenceDataField_'][id*='windowStart_windowStartDate']",
                recurrenceEndNoDate: "INPUT[id*='_RecurrenceDataField_'][id*='endDateRangeType0']",
                recurrenceEndOccurrences: "INPUT[id*='_RecurrenceDataField_'][id*='endDateRangeType1']",
                recurrenceEndOccurrencesValue: "INPUT[id*='_RecurrenceDataField_'][id*='repeatInstances']",
                recurrenceEndByDate: "INPUT[id*='_RecurrenceDataField_'][id*='endDateRangeType2']",
                recurrenceEndByDateValue: "INPUT[id*='_RecurrenceDataField_'][id*='windowEnd_windowEndDate']"
            },
            elemMain: document.querySelector(this.selectors.control),
            get: function(){
                return this.elemMain.getAttribute("checked")
            },
            set: function(newValue){
                if ( newValue === true ){
                    this.elemMain.setAttribute("checked","checked");
                }
                else {
                    this.elemMain.removeAttribute("checked");
                }
            }
        }
    },
    SPFieldOverbook: function(fin,guid,displayName){
        return {
            fin: fin,
            guid: guid,
            fieldDisplayName: displayName,
            fieldType: "SPFieldOverbook",
            fieldSubType: "",
            selectors: {
                control: "INPUT[id='"+ fin +"_"+ guid +"_btnCheckOverbook']"
            },
            elemMain: document.querySelector(this.selectors.control),
            get: function(){},
            set: function(){}
        }
    },
    SPFieldFreeBusy: function(fin,guid,displayName){
        return {
            fin: fin,
            guid: guid,
            fieldDisplayName: displayName,
            fieldType: "SPFieldFreeBusy",
            fieldSubType: "",
            selectors: {
                control: "DIV[id='FreeBusyRootDiv']"
            },
            elemMain: document.querySelector(this.selectors.control),
            get: function(){},
            set: function(){}
        }
    },
    SPFieldFacilities: function(fin, guid,displayName){
        return {
            fin: fin,
            guid: guid,
            fieldDisplayName: displayName,
            fieldType: "SPFieldFacilities",
            fieldSubType: "",
            selectors: {
                control: "SELECT[type='text'][id='"+ fin +"_"+ guid +"_SelectCandidate']"
            },
            elements: {
                main: document.querySelector(this.selectors.control)
            },
            get: function(){
                return this.elements.main.value
            },
            set: function(newValue){
                this.elements.main.value = newValue;
            }
        }
    },
    SPFieldUser: {
        peopleOnly: function(fin,guid,displayName){
            return {
                fin: fin,
                guid: guid,
                fieldDisplayName: displayName,
                fieldType: "SPFieldUser",
                fieldSubType: "PeopleOnly",
                allowMultiple: false,
                selectors: {
                    control: "DIV[id='"+ fin +"_"+ guid +"_$ClientPeoplePicker']",
                    initialHelpText: "SPAN[id='"+ fin +"_"+ guid +"_$ClientPeoplePicker_InitialHelpText']",
                    waitImage: "IMG[id='"+ fin +"_"+ guid +"_$ClientPeoplePicker_WaitImage']",
                    resolvedList: "SPAN[id='"+ fin +"_"+ guid +"_$ClientPeoplePicker_ResolvedList']",
                    editorInput: "INPUT[id='"+ fin +"_"+ guid +"_$ClientPeoplePicker_EditorInput']",
                    fieldDisplayName: displayName,
                    legacy: "DIV[id$='_UserField_upLevelDiv'][title='"+this.fieldDisplayName+"']",
                    legacyOuterTable: "TABLE[id$='_UserField_OuterTable']",
                    legacyCheckNames: "A[id$='_UserField_checkNames']",
                    legacyBrowse: "A[id$='_UserField_browse']"
                },
                elemMain: function(){ return document.querySelector(this.selectors.control);},
                get: function(){
                    return this.elemMain.value;
                },
                set: function(newValue ){
                    this.elemMain.value = newValue;
                }
            }
        },
        peopleOrGroups: function(fin,guid,displayName){
            return {
                fin: fin,
                guid: guid,
                fieldDisplayName: displayName,
                fieldType: "SPFieldUser",
                fieldSubType: "PeopleOrGroups",
                allowMultiple: false,
                selectors: {
                    control: "DIV[id='"+ fin +"_"+ guid +"_$ClientPeoplePicker']",
                    initialHelpText: "SPAN[id='"+ fin +"_"+ guid +"_$ClientPeoplePicker_InitialHelpText']",
                    waitImage: "IMG[id='"+ fin +"_"+ guid +"_$ClientPeoplePicker_WaitImage']",
                    resolvedList: "SPAN[id='"+ fin +"_"+ guid +"_$ClientPeoplePicker_ResolvedList']",
                    editorInput: "INPUT[id='"+ fin +"_"+ guid +"_$ClientPeoplePicker_EditorInput']",
                    fieldDisplayName: displayName,
                    legacy: "DIV[id$='_UserField_upLevelDiv'][title='"+this.fieldDisplayName+"']",
                    legacyOuterTable: "TABLE[id$='_UserField_OuterTable']",
                    legacyCheckNames: "A[id$='_UserField_checkNames']",
                    legacyBrowse: "A[id$='_UserField_browse']"
                },
                elemMain: document.querySelector(this.selectors.control),
                get: function(){
                    return this.elemMain.value;
                },
                set: function(newValue ){
                    this.elemMain.value = newValue;
                }
            }
        }
    },
    SPFieldUserMulti: {
        peopleOnly: function(fin,guid,displayName){
            return {
                fin: fin,
                guid: guid,
                fieldDisplayName: displayName,
                fieldType: "SPFieldUser",
                fieldSubType: "PeopleOnly",
                allowMultiple: true,
                selectors: {
                    control: "DIV[id='"+ fin +"_"+ guid +"_$ClientPeoplePicker']",
                    initialHelpText: "SPAN[id='"+ fin +"_"+ guid +"_$ClientPeoplePicker_InitialHelpText']",
                    waitImage: "IMG[id='"+ fin +"_"+ guid +"_$ClientPeoplePicker_WaitImage']",
                    resolvedList: "SPAN[id='"+ fin +"_"+ guid +"_$ClientPeoplePicker_ResolvedList']",
                    editorInput: "INPUT[id='"+ fin +"_"+ guid +"_$ClientPeoplePicker_EditorInput']",
                    fieldDisplayName: displayName,
                    legacy: "DIV[id$='_UserField_upLevelDiv'][title='"+this.fieldDisplayName+"']",
                    legacyOuterTable: "TABLE[id$='_UserField_OuterTable']",
                    legacyCheckNames: "A[id$='_UserField_checkNames']",
                    legacyBrowse: "A[id$='_UserField_browse']"
                },
                elemMain: document.querySelector(this.selectors.control),
                get: function(){
                    return this.elemMain.value;
                },
                set: function(newValue ){
                    this.elemMain.value = newValue;
                }
            }
        },
        peopleOrGroups: function(fin,guid,displayName){
            return {
                fin: fin,
                guid: guid,
                fieldDisplayName: displayName,
                fieldType: "SPFieldUser",
                fieldSubType: "PeopleOrGroups",
                allowMultiple: true,
                selectors: {
                    control: "DIV[id='"+ fin +"_"+ guid +"_$ClientPeoplePicker']",
                    initialHelpText: "SPAN[id='"+ fin +"_"+ guid +"_$ClientPeoplePicker_InitialHelpText']",
                    waitImage: "IMG[id='"+ fin +"_"+ guid +"_$ClientPeoplePicker_WaitImage']",
                    resolvedList: "SPAN[id='"+ fin +"_"+ guid +"_$ClientPeoplePicker_ResolvedList']",
                    editorInput: "INPUT[id='"+ fin +"_"+ guid +"_$ClientPeoplePicker_EditorInput']",
                    fieldDisplayName: displayName,
                    legacy: "DIV[id$='_UserField_upLevelDiv'][title='"+this.fieldDisplayName+"']",
                    legacyOuterTable: "TABLE[id$='_UserField_OuterTable']",
                    legacyCheckNames: "A[id$='_UserField_checkNames']",
                    legacyBrowse: "A[id$='_UserField_browse']"
                },
                elemMain: document.querySelector(this.selectors.control),
                get: function(){
                    return this.elemMain.value;
                },
                set: function(newValue ){
                    this.elemMain.value = newValue;
                }
            }
        }
    },
    SPFieldUrl: {
        hyperlink: function(fin,guid,displayName){
            return {
                fin: fin,
                guid: guid,
                fieldDisplayName: displayName,
                fieldType: "SPFieldURL",
                fieldSubType: "Hyperlink",
                selectors: {
                    control: "INPUT[id='"+ fin +"_"+ guid +"_$UrlFieldUrl']",
                    description: "INPUT[id='"+ fin +"_"+ guid +"_$UrlFieldDescription']"
                },
                elemMain: document.querySelector(this.selectors.control),
                elemDescription: document.querySelector(this.selectors.description),
                get: function(){
                    try{
                        return "<a href='"+ this.elemMain.value +"'>"+ this.elemDescription.value +"</a>";
                    }
                    catch(err){
                        return "<a href='"+ this.elemMain.value +"'>"+ this.elemMain.value +"</a>";
                    }
                },
                set: function(newValueUrl, newValueDescription ){
                    this.elemMain.value = newValueUrl;
                    this.elemDescription.value = newValueDescription;
                }
            }
        },
        picture: function(fin,guid,displayName){
            return {
                fin: fin,
                guid: guid,
                fieldDisplayName: displayName,
                fieldType: "SPFieldURL",
                fieldSubType: "Picture",
                selectors: {
                    control: "INPUT[id='"+ fin +"_"+ guid +"_$UrlFieldUrl']",
                    description: "INPUT[id='"+ fin +"_"+ guid +"_$UrlFieldDescription']"
                },
                elemMain: document.querySelector(this.selectors.control),
                elemDescription: document.querySelector(this.selectors.description),
                get: function(){
                    try{
                        return "<a href='"+ this.elemMain.value +"'>"+ this.elemDescription.value +"</a>";
                    }
                    catch(err){
                        return "<a href='"+ this.elemMain.value +"'>"+ this.elemMain.value +"</a>";
                    }
                },
                set: function(newValueUrl, newValueDescription ){
                    this.elemMain.value = newValueUrl;
                    this.elemDescription.value = newValueDescription;
                }
            }
        }
    }
};
var pplGrps = pplGrps || {
    instances: {},
    legacyInstances: {},
    formFields: {},
    pageWebParts: [],
    pickerControls: [],
    bDebug: false,
    editMode: false,
    userCanEditSettings: false,
    PeopleManagerServiceUnavailable: false,
    currentSettings: {},
    listFields: [],
    log: function (message, bIgnoreDebugReq) {
        if (typeof (bIgnoreDebugReq) === "undefined") {
            var bIgnoreDebugReq = false;
        }
        if (pplGrps.bDebug === true || bIgnoreDebugReq === true) {
            try {
                console.log(message);
            } catch (err) {}
        }
    },
    isPageInEditMode: function () {
        /*https://sharepoint.stackexchange.com/questions/149096/a-way-to-identify-when-page-is-in-edit-mode-for-javascript-purposes*/
        var result = (window.MSOWebPartPageFormName != undefined) && ((document.forms[window.MSOWebPartPageFormName] && document.forms[window.MSOWebPartPageFormName].MSOLayout_InDesignMode && ("1" == document.forms[window.MSOWebPartPageFormName].MSOLayout_InDesignMode.value)) || (document.forms[window.MSOWebPartPageFormName] && document.forms[window.MSOWebPartPageFormName]._wikiPageMode && ("Edit" == document.forms[window.MSOWebPartPageFormName]._wikiPageMode.value)));
        pplGrps.editMode = result || false;
        return result || false;
    },
    ajax: {
        lastCall: {},
        create: function (restURL, object, fxCallback, fxFailed) {
            pplGrps.ajax.lastCall = {
                xhr: null,
                readyState: null,
                data: null,
                status: null,
                url: restURL,
                error: null
            };
            var xhr = new XMLHttpRequest();
            xhr.open('POST', restURL, true);
            xhr.setRequestHeader("X-RequestDigest", document.getElementById("__REQUESTDIGEST").value);
            xhr.setRequestHeader("IF-MATCH", "*");
            xhr.setRequestHeader("X-HTTP-Method", "POST");
            xhr.setRequestHeader("accept", "application/json;odata=verbose");
            xhr.setRequestHeader("content-type", "application/json;odata=verbose");
            xhr.onreadystatechange = function () {
                pplGrps.ajax.lastCall.readyState = xhr.readyState;
                pplGrps.ajax.lastCall.status = xhr.status;
                if (xhr.readyState === 4) {
                    if (xhr.status !== 201) {
                        pplGrps.ajax.lastCall.data = xhr.response;
                        if (typeof (fxFailed) === "function") {
                            fxFailed(xhr, xhr.response, xhr.status);
                        }
                    } else {
                        /*var resp = JSON.parse(xhr.response);*/
                        pplGrps.ajax.lastCall.data = xhr.response;
                        if (typeof (fxCallback) === "function") {
                            fxCallback(xhr, xhr.response);
                        }
                    }
                }
            };
            xhr.send(object);
        },
        update: function (restURL, object, fxCallback, fxFailed) {
            pplGrps.ajax.lastCall = {
                xhr: null,
                readyState: null,
                data: null,
                status: null,
                url: restURL,
                error: null
            };
            var xhr = new XMLHttpRequest();
            xhr.open('POST', restURL, true);
            xhr.setRequestHeader("X-RequestDigest", document.getElementById("__REQUESTDIGEST").value);
            xhr.setRequestHeader("IF-MATCH", "*");
            xhr.setRequestHeader("X-HTTP-Method", "MERGE");
            xhr.setRequestHeader("accept", "application/json;odata=verbose");
            xhr.setRequestHeader("content-type", "application/json;odata=verbose");
            xhr.onreadystatechange = function () {
                pplGrps.ajax.lastCall.readyState = xhr.readyState;
                pplGrps.ajax.lastCall.status = xhr.status;
                if (xhr.readyState === 4) {
                    if (xhr.status !== 200 && xhr.status !== 204) {
                        pplGrps.ajax.lastCall.data = xhr.response;
                        if (typeof (fxFailed) === "function") {
                            fxFailed(xhr, xhr.response, xhr.status);
                        }
                    } else {
                        var resp = xhr.response;
                        pplGrps.ajax.lastCall.data = resp;
                        if (typeof (fxCallback) === "function") {
                            fxCallback(xhr, resp);
                        }
                    }
                }
            };
            xhr.send(object);
        },
        delete: function (restURL, object, fxCallback) {
            pplGrps.ajax.lastCall = {
                xhr: null,
                readyState: null,
                data: null,
                status: null,
                url: restURL,
                error: null
            };
            var xhr = new XMLHttpRequest();
            xhr.open('POST', restURL, true);
            xhr.setRequestHeader("X-RequestDigest", document.getElementById("__REQUESTDIGEST").value);
            xhr.setRequestHeader("IF-MATCH", "*");
            xhr.setRequestHeader("X-HTTP-Method", "DELETE");
            xhr.setRequestHeader("accept", "application/json;odata=verbose");
            xhr.setRequestHeader("content-type", "application/json;odata=verbose");
            xhr.onreadystatechange = function () {
                pplGrps.ajax.lastCall.readyState = xhr.readyState;
                pplGrps.ajax.lastCall.status = xhr.status;
                if (xhr.readyState === 4) {
                    if (xhr.status !== 200) {
                        pplGrps.ajax.lastCall.data = xhr.response;
                        if (typeof (fxFailed) === "function") {
                            fxFailed(xhr, xhr.response, xhr.status);
                        }
                    } else {
                        var resp = xhr.response;
                        pplGrps.ajax.lastCall.data = resp;
                        if (typeof (fxCallback) === "function") {
                            fxCallback(xhr, resp);
                        }
                    }
                }
            };
            xhr.send(object);
        },
        read: function (restURL, fxCallback, fxLastPage, fxFailed, sDataType, bCache) {
            if (typeof (sDataType) === "undefined") {
                var sDataType = "json";
            }
            if ( typeof(bCache) === "undefined" ){
                var bCache = true;
            }
            pplGrps.ajax.lastCall = {
                xhr: null,
                readyState: null,
                data: null,
                status: null,
                url: restURL,
                error: null
            };
            var xhr = new XMLHttpRequest();
            xhr.open('GET', restURL, true);
            if  ( bCache === true ){
                var now = new Date();
                /* 4 hours later */
                var later = new Date(now.valueOf() + (1000 * 60 * 60 * 4));
                xhr.setRequestHeader("Expires", later);
                xhr.setRequestHeader("Last-Modified", now);
                xhr.setRequestHeader("Cache-Control", "Public");
            }
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
                    pplGrps.log("pplGrps.ajax.read... unknown sDataType value |" + sDataType + "|", true);
            }

            xhr.onreadystatechange = function () {
                pplGrps.ajax.lastCall.readyState = xhr.readyState;
                pplGrps.ajax.lastCall.status = xhr.status;
                if (xhr.readyState === 4) {
                    if (xhr.status !== 200) {
                        pplGrps.ajax.lastCall.data = xhr.response;
                        if (typeof (fxFailed) === "function") {
                            fxFailed(xhr, xhr.response, xhr.status);
                        }
                    } else {
                        switch (sDataType) {
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
                        pplGrps.ajax.lastCall.data = resp;
                        if (typeof (fxCallback) === "function") {
                            fxCallback(xhr, resp);
                        }
                        if (typeof (resp.d.__next) !== "undefined") {
                            pplGrps.ajax.read(resp.d.__next, fxCallback, fxLastPage);
                        } else if (typeof (fxLastPage) === "function") {
                            fxLastPage(xhr, resp);
                        }
                    }
                }
            };
            pplGrps.ajax.lastCall.xhr = xhr;
            xhr.send();
        },
        createSettingsList: function (afterCreatingSettingsListFx) {
            var sListDisplayName = "dsMagic-pplGrps";
            var listDef = null;
            if (typeof (afterCreatingSettingsListFx) === "undefined") {
                var afterCreatingSettingsListFx = function () {
                    pplGrps.log("Created and configured the settings list", true);
                };
            }
            var oCreateList = {
                __metadata: {
                    type: 'SP.List'
                },
                AllowContentTypes: false,
                BaseTemplate: 100,
                ContentTypesEnabled: false,
                Description: "Stores the settings for dsMagic pickers JS API",
                Title: "dsMagic-pplGrps"
            };
            pplGrps.ajax.create(_spPageContextInfo.webAbsoluteUrl + "/_api/web/lists", JSON.stringify(oCreateList), function (x, d) {
                pplGrps.log("dsMagic-pplGrps settings list created", true);
                pplGrps.log(x);
                pplGrps.log(d);
                var responseData = JSON.parse(d);
                var listRestURL = responseData.d.__metadata.uri;
                var oSettings1 = {
                    __metadata: {
                        type: 'SP.List'
                    },
                    EnableVersioning: true,
                    EnableModeration: false,
                    EnableAttachments: false,
                    MajorVersionLimit: 10,
                    DraftVersionVisibility: 0,
                    EnableMinorVersions: false,
                    ForceCheckout: false
                };
                pplGrps.ajax.update(listRestURL, JSON.stringify(oSettings1), function (x1, d1) {
                    pplGrps.log("dsMagic-pplGrps list - versioning enabled", true);
                    /*
                    var oSettings2 = {
                        __metadata: {
                            type: 'SP.List'
                        },
                        MajorVersionLimit: 10,
                        DraftVersionVisibility: 0,
                        MajorWithMinorVersionsLimit: 0,
                        EnableMinorVersions: false
                    };
                    pplGrps.ajax.update(listRestURL, JSON.stringify(oSettings2), function (x2, d2) {
                        pplGrps.log("dsMagic-pplGrps list - versioning limits set", true);
                    */
                        var fld_blbSettings = {
                            __metadata: {
                                type: 'SP.FieldMultiLineText'
                            },
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
                        pplGrps.ajax.create(listRestURL + "/Fields", JSON.stringify(fld_blbSettings), function (xf1, df1) {
                            pplGrps.log("dsMagic-pplGrps list - added field 'blbSettings'", true);
                            var fld_fieldName = {
                                __metadata: {
                                    type: 'SP.FieldText'
                                },
                                Title: "fieldName",
                                FieldTypeKind: 2,
                                Required: false,
                                EnforceUniqueValues: false,
                                StaticName: "fieldName"
                            };
                            pplGrps.ajax.create(listRestURL + "/Fields", JSON.stringify(fld_fieldName), function (xf2, df2) {
                                pplGrps.log("dsMagic-pplGrps list - added field 'fieldName'", true);
                                if (typeof (afterCreatingSettingsListFx) === "function") {
                                    afterCreatingSettingsListFx();
                                }
                                if (confirm("Settings list created. Would you like to refresh the page?") === true) {
                                    window.location.reload();
                                }
                            }, function () {
                                pplGrps.log("dsMagic-pplGrps list - failed to add field 'fieldName'", true);
                            });
                        }, function () {
                            pplGrps.log("dsMagic-pplGrps list - failed to add field 'blbSettings'", true);
                        });
                    /*
                    }, function () {
                        pplGrps.log("failed to set version limits on dsMagic-pplGrps", true);
                    });
                    */
                }, function () {
                    pplGrps.log("failed to enable versioning on dsMagic-pplGrps", true);
                });
            }, function () {
                pplGrps.log("failed to create dsMagic-pplGrps settings list", true);
            });
        },
        captureSettings: function (formURL, fieldName, blbSettings) {
            var existingSettings = pplGrps.checkIfSettingsExistForField(fieldName);
            if ( existingSettings > 0 ){
                pplGrps.log("Updating existing settings list item",true);
                pplGrps.ajax.updateSettings(pplGrps.currentSettings[fieldName].Id, blbSettings);
            }
            else {
                pplGrps.log("Saving new settings list item",true);
                var oItem = {
                    __metadata: {
                        type: 'SP.Data.DsMagicpplGrpsListItem'
                    },
                    Title: formURL,
                    fieldName: fieldName,
                    blbSettings: JSON.stringify(blbSettings)
                };
                pplGrps.ajax.create(
                    _spPageContextInfo.webAbsoluteUrl + "/_api/web/lists/GetByTitle('dsMagic-pplGrps')/Items", 
                    JSON.stringify(oItem), 
                    function (x, d) {
                        pplGrps.log("Created list item for Settings");
                        pplGrps.log(x);
                        pplGrps.log(d);
                        if (confirm("Settings have been saved. Would you like to refresh the page?") === true) {
                            window.location.reload();
                        }
                    }, function (x,s,e) {
                        pplGrps.log("Failed to create list item for Settings", true);
                        pplGrps.log(x, true);
                        pplGrps.log(s, true);
                        pplGrps.log(e, true);
                    }
                );
            }
        },
        getSettings: function (fxAfter, fxFailed) {
            var restURL = _spPageContextInfo.webAbsoluteUrl + "/_api/web/lists/GetByTitle('dsMagic-pplGrps')/Items?$filter=Title eq '" + _spPageContextInfo.serverRequestPath + "'";
            pplGrps.ajax.read(restURL, function (x, d) {
                pplGrps.log("Got a page of settings", true);
                for (var iD = 0; iD < d.d.results.length; iD++) {
                    pplGrps.settings.push(d.d.results[iD]);
                    var blbSettings = JSON.parse(d.d.results[iD].blbSettings);
                    pplGrpsListeners.push(blbSettings);
                    pplGrps.currentSettings[d.d.results[iD].fieldName] = {
                        blbSettings: blbSettings,
                        Id: d.d.results[iD].Id
                    };
                }
            }, function (xLast, dLast) {
                pplGrps.log("Last page of settings", true);
                //pplGrps.setupListeners();
                if (typeof (fxAfter) === "function") {
                    fxAfter();
                }
            }, function () {
                pplGrps.log("Failed to get settings", true);
                if (typeof (fxFailed) === "function") {
                    fxFailed();
                }
            });
        },
        checkIfSettingsExistForField: function(fin){
            var finSettingsId = 0;
            for ( fieldInternalName in pplGrps.currentSettings ){
                if ( fieldInternalName === fin ) {
                    finSettingsId = pplGrps.currentSettings[fieldInternalName].Id;
                    break;
                }
            }
            return finSettingsId;
        },
        updateSettings: function (itemID, blbSettings) {
            var oItem = {
                __metadata: {
                    type: 'SP.Data.DsMagicpplGrpsListItem'
                },
                blbSettings: JSON.stringify(blbSettings)
            };
            pplGrps.ajax.update(
                _spPageContextInfo.webAbsoluteUrl + "/_api/web/lists/GetByTitle('dsMagic-pplGrps')/Items(" + itemID + ")", 
                JSON.stringify(oItem), 
                function (x, d) {
                    pplGrps.log("Updated list item for Settings", true);
                    pplGrps.log(x, true);
                    pplGrps.log(d, true);
                    if (confirm("Settings have been saved. Would you like to refresh the page?") === true) {
                        window.location.reload();
                    }
                }, function () {
                    pplGrps.log("Failed to update list item for Settings", true);
                }
            );
        },
        checkIfSettingsListExists: function (fxExists, fxDoesNotExist) {
            var restURL = _spPageContextInfo.webAbsoluteUrl + "/_api/web/lists/GetByTitle('dsMagic-pplGrps')";
            pplGrps.ajax.read(restURL, function (x, d) {
                pplGrps.log("Settings list exists", true);
                pplGrps.lists['dsMagic-pplGrps'] = d.d;
                pplGrps.ajax.getListPermissions("pplGrps.lists['dsMagic-pplGrps']", function (settingsList) {
                    pplGrps.log(settingsList.EffectiveBasePermissions, true);
                    pplGrps.ajax.checkForItemEditAddDelete('dsMagic-pplGrps', function (canEditAndAdd) {
                        pplGrps.userCanEditSettings = canEditAndAdd;
                        if (typeof (fxExists) === "function") {
                            fxExists(canEditAndAdd);
                        }
                    });
                });
            }, undefined, function () {
                pplGrps.log("Settings list does not exist", true);
                if (typeof (fxDoesNotExist) === "function") {
                    fxDoesNotExist();
                }
            });
        },
        getListPermissions: function (strCaptureResultsIn, afterFx) {
            var restURL = eval(strCaptureResultsIn + ".__metadata.uri");
            restURL += "?$select=EffectiveBasePermissions";
            var fxCallback = function (xhr, data) {
                eval(strCaptureResultsIn + ".EffectiveBasePermissions = " + JSON.stringify(data.d.EffectiveBasePermissions) + ";");
                if (typeof (afterFx) === "function") {
                    afterFx(eval(strCaptureResultsIn));
                }
            }
            pplGrps.ajax.read(restURL, fxCallback);
        },
        checkForItemEditAddDelete: function (restlistname, afterFx) {
            // supposed to check like this:
            // http://1.dsmagicsp.cloudappsportal.com/_api/web/lists(guid'd11a71ff-4170-4725-a804-9d76153d0ebf')/getusereffectivepermissions(@user)?@user=%27i%3A0%23%2Ew%7Cshakir%5Cjohn%27
            // or for me
            // http://1.dsmagicsp.cloudappsportal.com/_api/web/lists(guid'd11a71ff-4170-4725-a804-9d76153d0ebf')/getusereffectivepermissions(@user)?@user=%27i%3A0%23%2Ew%7Cschauer%5Cdaniel%27
            // unencoded, the last parameter looks like this
            // http://1.dsmagicsp.cloudappsportal.com/_api/web/lists(guid'd11a71ff-4170-4725-a804-9d76153d0ebf')/getusereffectivepermissions(@user)?@user='i:0#.w|shakir\john'
            // "http://expertsoverlunch.com/sandbox/_api/web/lists/GetByTitle('dsMagic-pplGrps')/getusereffectivepermissions(@user)?@user='"+ encodeURIComponent(userLogin) +"'"
            var up = new SP.BasePermissions();
            up.set(SP.PermissionKind.viewListItems);
            up.set(SP.PermissionKind.editListItems);
            up.set(SP.PermissionKind.addListItems);
            up.set(SP.PermissionKind.deleteListItems);
            // High must be < up.$5_1 to correctly detect no access for anonymous users
            // but must be > up.$5_1 to correctly detect appropriate access for daniel logged in
            if (typeof (_spPageContextInfo.userId) === "undefined") {
                if (pplGrps.lists[restlistname].EffectiveBasePermissions.High < up.$5_1) {
                    if (typeof (afterFx) === "function") {
                        afterFx(true);
                    }
                    return true;
                } else {
                    /*
                    if (typeof(afterFx) === "function") { afterFx(false); }
                    return false;
                    */
                    if (typeof (afterFx) === "function") {
                        afterFx(true);
                    }
                    return true;
                }
            } else {
                if (pplGrps.lists[restlistname].EffectiveBasePermissions.High > up.$5_1) {
                    if (typeof (afterFx) === "function") {
                        afterFx(true);
                    }
                    return true;
                } else {
                    /*
                    if (typeof(afterFx) === "function") { afterFx(false); }
                    return false;
                    */
                    if (typeof (afterFx) === "function") {
                        afterFx(true);
                    }
                    return true;
                }
            }
        }
    },
    lists: {},
    settings: [],
    userProperties: {},
    userGroups: {},
    getUserProperties: function (afterFx) {
        /* current user */
        pplGrps.ajax.read(
            _spPageContextInfo.webAbsoluteUrl + "/_api/web/CurrentUser?$expand=Groups,UserId,Groups/Users",
            function (xhr, data) {
                pplGrps.log(data);
                for (userProp in data.d) {
                    try {
                        pplGrps.userProperties[userProp] = data.d[userProp];
                    } catch (err) {
                        pplGrps.log("failed to capture user prop |" + userProp + "|");
                    }
                }
            },
            function (xhr, data) {
                pplGrps.log(data);
                if (typeof (afterFx) === "function") {
                    afterFx();
                }
            },
            function () {
                pplGrps.log("failed to get user properties");
            }
        );

    },
    getUserPropertiesByLoginName: function (loginName, afterFx, pickerID, userKey) {
        pplGrps.userProperties = {};
        pplGrps.ajax.read(
            _spPageContextInfo.webAbsoluteUrl + "/_api/web/SiteUsers(@v)?@v='" + encodeURIComponent(loginName) + "'&$expand=Groups,UserId,Groups/Users",
            function (xhr, data) {
                pplGrps.log(data);
                for (userProp in data.d) {
                    try {
                        pplGrps.userProperties[userProp] = data.d[userProp];
                        try{pplGrps.instances[pickerID].users[userKey][userProp] = data.d[userProp];}catch(err){}
                        if (userProp === "Groups") {
                            if (typeof (data.d.Groups.results) !== "undefined") {
                                for (var iG = 0; iG < data.d.Groups.results.length; iG++) {
                                    pplGrps.userGroups[data.d.Groups.results[iG].Title] = data.d.Groups.results[iG];
                                    //try{pplGrps.instances[pickerID].groups[data.d.Groups.results[iG].Key] = data.d.Groups.results[iG];}catch(err){}
                                }
                            }
                        }
                    } catch (err) {
                        pplGrps.log("failed to capture user prop |" + userProp + "|");
                    }
                }
                if (typeof (afterFx) === "function") {
                    if ( typeof(pplGrps.instances[pickerID]) !== "undefined" ){
                        if ( typeof(pplGrps.instances[pickerID].users[userKey]) !== "undefined" ){
                            afterFx(pplGrps.instances[pickerID].users[userKey]);
                        }
                    }
                    else {
                        afterFx(pplGrps.userProperties);
                    }
                }
            },
            function (xhr, data) {
                pplGrps.log(data);

            },
            function () {
                pplGrps.log("failed to get user properties");
            }
        );
    },
    getUserProfileFromPeopleManager: function(loginName, afterFx, pickerID, userKey){
        /*http://expertsoverlunch.com/sandbox/_api/SP.UserProfiles.PeopleManager/getPropertiesFor(@v)?@v=%27i%3A0%23.f%7Cexpertsoverlunchmp%7Cdaniel%27*/
        if ( pplGrps.PeopleManagerServiceUnavailable === false ){
            pplGrps.log("Requesting user profile via SP.UserProfiles.PeopleManager service",true);
            pplGrps.ajax.read(
                _spPageContextInfo.webAbsoluteUrl + "/_api/SP.UserProfiles.PeopleManager/getPropertiesFor(@v)?@v=%27" + loginName + "%27",
                function (xhr, data) {
                    var returnUserObject = null;
                    /* first expand user info under pplGrps.userProperties */
                    pplGrps.userProperties["PeopleManagerProfile"] = data.d;
                    if ( typeof(data.d.UserProfileProperties) !== "undefined" ){
                        var arrUserProfileProps = data.d.UserProfileProperties;
                        for ( var iUPP = 0; iUPP < arrUserProfileProps.length; iUPP++ ){
                            profileProp = arrUserProfileProps[iUPP];
                            pplGrps.userProperties[profileProp.Key] = profileProp.Value;
                        }
                    }
                    pplGrps.log("Expanded SP user at pplGrps.userProperties via SP.UserProfiles.PeopleManager service",true);
                    pplGrps.log(pplGrps.userProperties,true);
                    returnUserObject = pplGrps.userProperties;
                    /* now expand user info under pplGrps.instances[pickerID].users[userKey] */
                    if ( typeof(pickerID) !== "undefined" && typeof(userKey) !== "undefined" ) {
                        if ( typeof(pplGrps.instances[pickerID]) !== "undefined" ){
                            if ( typeof(pplGrps.instances[pickerID].users[userKey]) !== "undefined" ){
                                pplGrps.instances[pickerID].users[userKey]["PeopleManagerProfile"] = data.d;
                                if ( typeof(data.d.UserProfileProperties) !== "undefined" ){
                                    var arrUserProfileProps = data.d.UserProfileProperties;
                                    for ( var iUPP = 0; iUPP < arrUserProfileProps.length; iUPP++ ){
                                        profileProp = arrUserProfileProps[iUPP];
                                        pplGrps.instances[pickerID].users[userKey][profileProp.Key] = profileProp.Value;
                                    }
                                }
                                pplGrps.log("Expanded SP user at pplGrps.instances['"+ pickerID +"'].users['"+ userKey +"'] via SP.UserProfiles.PeopleManager service",true);
                                returnUserObject = pplGrps.instances[pickerID].users[userKey];
                            }    
                        }
                    }
                    if (typeof (afterFx) === "function") {
                        afterFx(returnUserObject);
                    }
                },
                function (xhr, data) {
                    pplGrps.log("Processed last page of rest results for the user profile retrieved via SP.UserProfiles.PeopleManager service",true);
                },
                function (xhr, data, status) {
                    pplGrps.log("Failed to get user profile via SP.UserProfiles.PeopleManager service",true);
                    pplGrps.log(xhr,true);
                    pplGrps.log(data,true);
                    pplGrps.log(status,true);
                    pplGrps.PeopleManagerServiceUnavailable = true;
                }
            );
        }
        else {
            pplGrps.log("The SP.UserProfiles.PeopleManager service seems to be unavailable, returning basic user profile",true);
            var returnUserObject = null;
            if ( typeof(pplGrps.userProperties) !== "undefined" ){
                returnUserObject = pplGrps.userProperties;
            }
            if ( typeof(pickerID) !== "undefined" && typeof(userKey) !== "undefined" ) {
                if ( typeof(pplGrps.instances[pickerID]) !== "undefined" ){
                    if ( typeof(pplGrps.instances[pickerID].users[userKey]) !== "undefined" ){
                        returnUserObject = pplGrps.userProperties;
                    }
                }
            }
            if (typeof (afterFx) === "function") {
                afterFx(returnUserObject);
            }
        }
    },
    getCurrentUserProfileFromPeopleManager: function(afterFx, pickerID, userKey){
        /*http://expertsoverlunch.com/sandbox/_api/SP.UserProfiles.PeopleManager/getPropertiesFor(@v)?@v=%27i%3A0%23.f%7Cexpertsoverlunchmp%7Cdaniel%27*/
        if ( pplGrps.PeopleManagerServiceUnavailable === false ){
            pplGrps.log("Requesting user profile via SP.UserProfiles.PeopleManager service",true);
            pplGrps.ajax.read(
                _spPageContextInfo.webAbsoluteUrl + "/_api/SP.UserProfiles.PeopleManager/getmyproperties",
                function (xhr, data) {
                    var returnUserObject = null;
                    /* first expand user info under pplGrps.userProperties */
                    pplGrps.userProperties["PeopleManagerProfile"] = data.d;
                    if ( typeof(data.d.UserProfileProperties) !== "undefined" ){
                        var arrUserProfileProps = data.d.UserProfileProperties;
                        for ( var iUPP = 0; iUPP < arrUserProfileProps.length; iUPP++ ){
                            profileProp = arrUserProfileProps[iUPP];
                            pplGrps.userProperties[profileProp.Key] = profileProp.Value;
                        }
                    }
                    pplGrps.log("Expanded SP user at pplGrps.userProperties via SP.UserProfiles.PeopleManager service",true);
                    pplGrps.log(pplGrps.userProperties,true);
                    returnUserObject = pplGrps.userProperties;
                    /* now expand user info under pplGrps.instances[pickerID].users[userKey] */
                    if ( typeof(pickerID) !== "undefined" && typeof(userKey) !== "undefined" ) {
                        if ( typeof(pplGrps.instances[pickerID]) !== "undefined" ){
                            if ( typeof(pplGrps.instances[pickerID].users[userKey]) !== "undefined" ){
                                pplGrps.instances[pickerID].users[userKey]["PeopleManagerProfile"] = data.d;
                                if ( typeof(data.d.UserProfileProperties) !== "undefined" ){
                                    var arrUserProfileProps = data.d.UserProfileProperties;
                                    for ( var iUPP = 0; iUPP < arrUserProfileProps.length; iUPP++ ){
                                        profileProp = arrUserProfileProps[iUPP];
                                        pplGrps.instances[pickerID].users[userKey][profileProp.Key] = profileProp.Value;
                                    }
                                }
                                pplGrps.log("Expanded SP user at pplGrps.instances['"+ pickerID +"'].users['"+ userKey +"'] via SP.UserProfiles.PeopleManager service",true);
                                returnUserObject = pplGrps.instances[pickerID].users[userKey];
                            }    
                        }
                    }
                    if (typeof (afterFx) === "function") {
                        afterFx(returnUserObject);
                    }
                },
                function (xhr, data) {
                    pplGrps.log("Processed last page of rest results for the user profile retrieved via SP.UserProfiles.PeopleManager service",true);
                },
                function (xhr, data, status) {
                    pplGrps.log("Failed to get user profile via SP.UserProfiles.PeopleManager service",true);
                    pplGrps.log(xhr,true);
                    pplGrps.log(data,true);
                    pplGrps.log(status,true);
                    pplGrps.PeopleManagerServiceUnavailable = true;
                }
            );
        }
        else {
            pplGrps.log("The SP.UserProfiles.PeopleManager service seems to be unavailable, returning basic user profile",true);
            var returnUserObject = null;
            if ( typeof(pplGrps.userProperties) !== "undefined" ){
                returnUserObject = pplGrps.userProperties;
            }
            if ( typeof(pickerID) !== "undefined" && typeof(userKey) !== "undefined" ) {
                if ( typeof(pplGrps.instances[pickerID]) !== "undefined" ){
                    if ( typeof(pplGrps.instances[pickerID].users[userKey]) !== "undefined" ){
                        returnUserObject = pplGrps.userProperties;
                    }
                }
            }
            if (typeof (afterFx) === "function") {
                afterFx(returnUserObject);
            }
        }
    },
    getPageWebPartsAndFormFields: function (fxEach, afterFx) {
        var collWPs = document.querySelectorAll("[id^='MSOZoneCell_WebPartWPQ']");
        for (var webPart = 0; webPart < collWPs.length; webPart++) {
            var wp = {
                id: collWPs[webPart].id,
                num: 0
            };
            try {
                wp.num = collWPs[webPart].id.replace("MSOZoneCell_WebPartWPQ", "");
            } catch (err) {}
            try {
                wp.title = document.getElementById("WebPartTitleWPQ" + wp.num).innerHTML;
            } catch (err) {}
            try {
                wp.body = document.getElementById("WebPartWPQ" + wp.num).innerHTML;
            } catch (err) {}
            try {
                wp.guid = document.getElementById("WebPartWPQ" + wp.num).getAttribute("webpartid");
            } catch (err) {}
            pplGrps.log(wp,true);
            pplGrps.pageWebParts.push(wp);
            /* check if we have a WPQ#FormCtx object defined for this web part; then loop thru its fields and capture them in pplGrps.formFields */
            try {
                var wpCTX = window["WPQ" + wp.num + "FormCtx"];
                if (typeof (pplGrps.formFields["WPQ" + wp.num + "FormCtx"]) === "undefined") {
                    pplGrps.formFields["WPQ" + wp.num + "FormCtx"] = {};
                }
                for (formField in wpCTX.ListSchema) {
                    pplGrps.formFields["WPQ" + wp.num + "FormCtx"][formField] = wpCTX.ListSchema[formField];
                }
            } catch (err) {
                pplGrps.log("WebPart |" + wp.num + "| is not a standard form",true);
                if (typeof (pplGrps.formFields["dsWPQ"+webPart+"FormCtx"]) === "undefined" ) {
                    pplGrps.formFields["dsWPQ"+webPart+"FormCtx"] = {}
                }
                var collFormFields = collWPs[webPart].querySelectorAll(".ms-formbody");
                pplGrps.log("WebPart |" + wp.num + "| contains "+ collFormFields.length,true);
                for ( var iField = 0; iField < collFormFields.length; iField++ ){
                    var oField = collFormFields[iField];
                    pplGrps.log(oField);
                    var firstComment = oField.innerHTML.substr(0,oField.innerHTML.indexOf("-->"));
                    pplGrps.log(firstComment);
                    var FieldName = firstComment.substr(firstComment.indexOf("FieldName=")+11,200);
                    FieldName = FieldName.substr(0,FieldName.indexOf("\n")-1);
                    var FieldInternalName = firstComment.substr(firstComment.indexOf("FieldInternalName=")+19,200);
                    FieldInternalName = FieldInternalName.substr(0,FieldInternalName.indexOf("\n")-1);
                    var FieldType = firstComment.substr(firstComment.indexOf("FieldType=")+11,200);
                    FieldType = FieldType.substr(0,FieldType.indexOf("\n")-1);
                    pplGrps.log(FieldName +"|"+ FieldInternalName +"|"+ FieldType);
                    pplGrps.formFields["dsWPQ"+webPart+"FormCtx"][FieldInternalName] = {
                        Description: "",
                        Direction: "none",
                        FieldType: FieldType,
                        Hidden: false,
                        IMEMode: null,
                        Id: "",
                        Name: FieldInternalName,
                        ReadOnlyField: false,
                        Required: false,
                        Title: FieldName,
                        Type: FieldType
                    };
                    pplGrps.log(pplGrps.formFields["dsWPQ"+webPart+"FormCtx"][FieldInternalName]);
                    if ( FieldType.indexOf("User") >= 0 || FieldType.indexOf("Group") >= 0 ){
                        /* is picker */
                        pplGrps.formFields["dsWPQ"+webPart+"FormCtx"][FieldInternalName].Id = oField.querySelectorAll(".ms-usereditor")[0].id;
                        pplGrps.formFields["dsWPQ"+webPart+"FormCtx"][FieldInternalName].TopLevelElementId = oField.querySelectorAll("[id$='_UserField_upLevelDiv']")[0].id;
                        
                        var fxSetOnKeyUp = function(controlID){
                            pplGrps.log("Setup event listener on picker hidden input onchange for control |"+ controlID+"|",true);
                            document.getElementById(controlID).onkeyup = function onkeyup(event) {
                                pplGrps.log(controlID);
                                pplGrps.log(event);
                                if ( event.which === 13 ){
                                    // user pressed 'enter'
                                    var wfTimeout = 0;
                                    var wfResolution = setInterval(function(){
                                        wfTimeout = wfTimeout + 1;
                                        var collEntries = document.querySelectorAll("[id='"+controlID+"'] .ms-entity-unresolved, [id='"+controlID+"'] .ms-entity-resolved");
                                        if ( collEntries.length > 0 ){
                                            pplGrps.log("detected resolution",true);
                                            pplGrps.instances[controlID].spcsom.OnUserResolvedClientScript(controlID, pplGrps.instances[controlID].spcsom.GetAllUserInfo())
                                            clearInterval(wfResolution);
                                        }
                                        if ( wfTimeout > 10 ){
                                            pplGrps.log("timedout",true);
                                            clearInterval(wfResolution);
                                        }
                                    },1000);
                                    //setTimeout(function(){
                                    //    pplGrps.instances[controlID].spcsom.OnUserResolvedClientScript(controlID, pplGrps.instances[controlID].spcsom.GetAllUserInfo())
                                    //},2000);
                                }
                                //else if ( event.which === 186 ){
                                //    // user pressed 'semicolon'
                                //    setTimeout(function(){
                                //        pplGrps.instances[controlID].spcsom.OnUserResolvedClientScript(controlID, pplGrps.instances[controlID].spcsom.GetAllUserInfo())
                                //    },2000);
                                //}
                                return event;
                            }
                        };
                        fxSetOnKeyUp(pplGrps.formFields["dsWPQ"+webPart+"FormCtx"][FieldInternalName].TopLevelElementId);

                    }
                }
            }
            if (typeof (fxEach) === "function") {
                fxEach(wp);
            }
        }
        if (typeof (afterFx) === "function") {
            afterFx();
        }
    },
    getAllPickers: function (afterFx) {
        if ( typeof(SPClientPeoplePicker) !== "undefined" ) {
            for (pickerControl in SPClientPeoplePicker.SPClientPeoplePickerDict) {
                if (typeof (pplGrps.instances[SPClientPeoplePicker.SPClientPeoplePickerDict[pickerControl].TopLevelElementId]) === "undefined") {
                    //pplGrps.instances[SPClientPeoplePicker.SPClientPeoplePickerDict[pickerControl].TopLevelElementId] = {
                    pplGrps.instances[pickerControl] = {
                        spcsom: SPClientPeoplePicker.SPClientPeoplePickerDict[pickerControl],
                        fieldDisplayName: "",
                        fin: "",
                        fieldGUID: SPClientPeoplePicker.SPClientPeoplePickerDict[pickerControl].TopLevelElementId.replace("_$ClientPeoplePicker", ""),
                        arrMapping: [],
                        users: {},
                        groups: {}
                    };
                    //pplGrps.instances[SPClientPeoplePicker.SPClientPeoplePickerDict[pickerControl].TopLevelElementId].fieldGUID = pplGrps.instances[SPClientPeoplePicker.SPClientPeoplePickerDict[pickerControl].TopLevelElementId].fieldGUID.substr(pplGrps.instances[SPClientPeoplePicker.SPClientPeoplePickerDict[pickerControl].TopLevelElementId].fieldGUID.indexOf("_") + 1, 38);
                    pplGrps.instances[pickerControl].fieldGUID = pplGrps.instances[pickerControl].fieldGUID.substr(pplGrps.instances[pickerControl].fieldGUID.indexOf("_") + 1, 38);
                    //var pickerFormField = pplGrps.findFormField("TopLevelElementId", SPClientPeoplePicker.SPClientPeoplePickerDict[pickerControl].TopLevelElementId);
                    var pickerFormField = pplGrps.findFormField("TopLevelElementId", pickerControl);
                    try {
                        //pplGrps.instances[SPClientPeoplePicker.SPClientPeoplePickerDict[pickerControl].TopLevelElementId].fieldDisplayName = pickerFormField.Title;
                        pplGrps.instances[pickerControl].fieldDisplayName = pickerFormField.Title;
                        //pplGrps.instances[SPClientPeoplePicker.SPClientPeoplePickerDict[pickerControl].TopLevelElementId].fin = pickerFormField.Name;
                        pplGrps.instances[pickerControl].fin = pickerFormField.Name;
                    } catch (err) {}
                }
            }
            if (typeof (afterFx) === "function") {
                afterFx();
            }
        }
        else {
            pplGrps.pickerControls = [];
            var collLegacyPickers = document.querySelectorAll(".ms-inputuserfield[id$='_upLevelDiv']");
            var thisPicker = function(elemPicker, fxLastPage){
                pplGrps.ajax.read(
                    _spPageContextInfo.webAbsoluteUrl +"/_api/web/lists(guid'"+ _spPageContextInfo.pageListId.replace("{","").replace("}","") +"')/Fields/GetByTitle('"+ elemPicker.getAttribute("title") +"')",
                    function(x,d){
                        var iPCIndex = pplGrps.pickerControls.push(d.d);
                        var elem = elemPicker;
                        pplGrps.legacyInstances[elem.id] = {
                            formFieldId: elem.id,
                            spcsom: {
                                formFieldId: elem.id,
                                HasResolvedUsers: function(){
                                    return pplGrps.legacyInstances[this.formFieldId].formField.querySelectorAll(".ms-entity-resolved").length > 0
                                },
                                OnUserResolvedClientScript: function(controlId, arrEntries){
                                    pplGrps.log("OnUserResolvedClientScript fired",true);
                                    pplGrps.log(controlId,true);
                                    pplGrps.log(arrEntries,true);
                                    
                                },
                                AddUserKeys: function(loginName){
                                    pplGrps.legacyInstances[this.formFieldId].formField.innerText = loginName;
                                    pplGrps.legacyInstances[this.formFieldId].spcsom.checkNames();
                                    /*
                                    var fieldCheckNames = document.getElementById(this.formFieldId.replace("upLevelDiv","checkNames"));
                                    fieldCheckNames.click();
                                    */
                                },
                                GetAllUserInfo: function(){
                                    var collResolved = document.querySelectorAll("[id='"+this.formFieldId+"'] .ms-entity-resolved");
                                    var arrReturn = [];
                                    for ( var iR = 0; iR < collResolved.length; iR++ ){
                                        arrReturn.push({
                                            "AccountName":collResolved[iR].id.substr(4,collResolved[iR].id.length-4),
                                            "LoginName": collResolved[iR].getAttribute("title"),
                                            "Name": collResolved[iR].childNodes[0],
                                            "EntityData":  collResolved[iR].childNodes[0].childNodes[0],
                                            "DisplayText": collResolved[iR].childNodes[0].getAttribute("displaytext"),
                                            "Key": collResolved[iR].id.substr(4,collResolved[iR].id.length-4),
                                            "IsResolved": true
                                        });
                                    }
                                    /*
                                    getting this sent
                                    http://expertsoverlunch.com/sandbox/_api/web/SiteUsers(@v)?@v=%27expertsoverlunchMP%3ADaniel%27&$expand=Groups,UserId,Groups/Users
                                    should look like this
                                    http://expertsoverlunch.com/sandbox/_api/web/SiteUsers(@v)?@v=%27i%3A0%23.f%7Cexpertsoverlunchmp%7Cmarcus%27&$expand=Groups,UserId,Groups/Users
                                    */
                                    return arrReturn;
                                },
                                checkNames: function onclick(event){
                                    //pplGrps.log(event,true);
                                    pplGrps.log(this.formFieldId,true);
                                    document.getElementById(this.formFieldId.replace("_UserField_upLevelDiv","_UserField_checkNames")).click();
                                    /*
                                    if(!ValidatePickerControl('ctl00_ctl26_g_e1c47bb7_99a4_482b_9c13_94ae882b8a90_ctl00_ctl05_ctl09_ctl00_ctl00_ctl04_ctl00_ctl00_UserField')){ ShowValidationError(); return false;} 
                                    var arg=getUplevel('ctl00_ctl26_g_e1c47bb7_99a4_482b_9c13_94ae882b8a90_ctl00_ctl05_ctl09_ctl00_ctl00_ctl04_ctl00_ctl00_UserField'); 
                                    var ctx='ctl00_ctl26_g_e1c47bb7_99a4_482b_9c13_94ae882b8a90_ctl00_ctl05_ctl09_ctl00_ctl00_ctl04_ctl00_ctl00_UserField';
                                    EntityEditorSetWaitCursor(ctx);
                                    WebForm_DoCallback(
                                         'ctl00$ctl26$g_e1c47bb7_99a4_482b_9c13_94ae882b8a90$ctl00$ctl05$ctl09$ctl00$ctl00$ctl04$ctl00$ctl00$UserField',
                                        arg,
                                        EntityEditorHandleCheckNameResult,
                                        ctx,
                                        EntityEditorHandleCheckNameError,
                                        true
                                    );
                                    return false;
                                    */
                                    //setTimeout(function(){
                                        pplGrps.legacyInstances[this.formFieldId].spcsom.OnUserResolvedClientScript(this.formFieldId, pplGrps.legacyInstances[this.formFieldId].spcsom.GetAllUserInfo());
                                    //},123);
                                }
                            },
                            fieldDisplayName: pplGrps.pickerControls[iPCIndex-1].Title,
                            fin: pplGrps.pickerControls[iPCIndex-1].InternalName,
                            fieldGUID: pplGrps.pickerControls[iPCIndex-1].Id,
                            arrMapping: [],
                            formField: elem,
                            users: {},
                            groups: {}
                        }
                        pplGrps.instances[elem.id] = pplGrps.legacyInstances[elem.id];
                    }, 
                    fxLastPage, 
                    function(){
                        pplGrps.log("Failed to get this list's settings",true);
                    }, 
                    "json"
                );
            }
            for ( var iLP = 0; iLP < collLegacyPickers.length; iLP++ ){
                if ( iLP+1 === collLegacyPickers.length ){
                    var fxLastPage = afterFx;
                }
                else {
                    var fxLastPage = undefined;
                }
                
                thisPicker(collLegacyPickers[iLP], fxLastPage);
            }
        }
    },
    findFormField: function (propertyName, propertyValue) {
        var oReturn = null;
        var bFoundFormField = false;
        for (formWP in pplGrps.formFields) {
            for (formField in pplGrps.formFields[formWP]) {
                if (typeof (pplGrps.formFields[formWP][formField][propertyName]) !== "undefined") {
                    if (pplGrps.formFields[formWP][formField][propertyName] === propertyValue) {
                        oReturn = pplGrps.formFields[formWP][formField];
                        bFoundFormField = true;
                        break;
                    }
                }
            }
            if (bFoundFormField === true) {
                break;
            }
        }
        return oReturn;
    },
    showConfigButtons: function (afterFx) {
        for (control in pplGrps.instances) {
            var pickerControl = null,
                pickerWrapper = null,
                configButton = null,
                configPopup = null;
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
            configButton.className = "pplGrpsConfigButton buttonish";
            configButton.id = pickerControl.id + "_configButton";
            /*pickerWrapper.parentElement.appendChild(configButton);*/
            //https://developer.mozilla.org/en-US/docs/Web/API/Element/insertAdjacentElement
            pickerWrapper.insertAdjacentElement('afterend', configButton);

            configPopup = document.createElement("DIV");
            configPopup.id = pickerControl.id + "_configWrapper";
            configPopup.className = "pplGrpsConfigPopup";
            var arrH = ["<DIV class='configPopupHeader'><span class='closeButton buttonish' id='"];
            arrH.push(control);
            arrH.push("_configCloseButton'>&times;</span></DIV><H3><DIV class='configForField'>");
            /*arrH.push(control);*/
            arrH.push(pplGrps.findFormField("TopLevelElementId", control).Title);
            arrH.push("</DIV></H3><H4>Upon resolved user, populate form fields...</H4><TABLE><thead><tr><td>Form Field</td><td>User Property&nbsp;<span class='userPropertyTooltip'><img border='0' alt='help with user property syntax' title='pplGrps.userProperties.propertyName\nor\npplGrps.userProperties.propertyArray.results|subArrayName.results' src='");
            arrH.push(_spPageContextInfo.siteAbsoluteUrl);
            arrH.push("/_layouts/15/images/helpicon.gif'/></span></td><td>Preview</td><td>Add</td><td>Remove</td></tr></thead><tbody id='");
            arrH.push(control);
            arrH.push("_configMappingRowsWrapper'><tr class='mappingRow'><td class='listField'><select class='listField'><option value='0'>Select list field</option>");
            for (formWP in pplGrps.formFields) {
                for (formField in pplGrps.formFields[formWP]) {
                    arrH.push("<option value='");
                    arrH.push(pplGrps.formFields[formWP][formField].Id);
                    arrH.push("'>");
                    arrH.push(pplGrps.formFields[formWP][formField].Name);
                    arrH.push("</option>");
                }
            }
            arrH.push("</select></td><td class='userProperty'>");
            /*
            arrH.push("<select class='userProperty'><option value='0'>Select user property</option>");
            for ( userProperty in pplGrps.userProperties ) {
                arrH.push("<option value='");
                arrH.push(userProperty);
                arrH.push("'>");
                arrH.push(userProperty);
                arrH.push("</option>");
            }
            arrH.push("</select>");
            */
            arrH.push("<INPUT class='userProperty' type='text' placeholder='e.g. pplGrps.userProperties.LoginName' value='' />");
            arrH.push("</td><td><span class='userPropertyValuePreview'></span></td><td><span class='addNew buttonish'>&plus;&nbsp;Add</span></td><td><span class='removeMapping buttonish'>&minus;&nbsp;Remove</span></td></tr></tbody></TABLE><DIV class='configPopupFooter'><span class='saveButton buttonish' id='");
            arrH.push(control);
            arrH.push("_configSaveButton'>Save</span></DIV><DIV class='saveSettings'></DIV>");
            configPopup.innerHTML = arrH.join("");

            /*pickerWrapper.parentElement.appendChild(configPopup);*/
            configButton.insertAdjacentElement('afterend', configPopup);
            if (pickerWrapper.parentElement.querySelectorAll(".ms-metadata").length > 0) {
                configPopup.insertAdjacentElement('afterend', document.createElement("BR"));
            }

            configButton.addEventListener("click", function (e) {
                this.nextSibling.style.display = "block";
            });

            configPopup.querySelector("INPUT.userProperty").addEventListener("keyup", function (e) {
                try {
                    if (this.value.indexOf("|") < 0) {
                        if (typeof (eval(this.value)) === "object") {
                            if (typeof (eval(this.value + ".length")) !== "undefined") {
                                this.parentElement.parentElement.querySelector("SPAN.userPropertyValuePreview").innerText = "array";
                            } else {
                                this.parentElement.parentElement.querySelector("SPAN.userPropertyValuePreview").innerText = eval(this.value);
                            }
                        } else {
                            this.parentElement.parentElement.querySelector("SPAN.userPropertyValuePreview").innerText = eval(this.value);
                        }
                    } else {
                        if (typeof (eval(this.value.split("|")[0] + "[0]." + this.value.split("|")[1])) === "object") {
                            if (typeof (eval(this.value.split("|")[0] + "[0]." + this.value.split("|")[1] + ".length")) !== "undefined") {
                                this.parentElement.parentElement.querySelector("SPAN.userPropertyValuePreview").innerText = "array";
                            } else {
                                this.parentElement.parentElement.querySelector("SPAN.userPropertyValuePreview").innerText = eval(this.value.split("|")[0] + "[0]." + this.value.split("|")[1]);
                            }
                        } else {
                            this.parentElement.parentElement.querySelector("SPAN.userPropertyValuePreview").innerText = eval(this.value.split("|")[0] + "[0]." + this.value.split("|")[1]);
                        }

                    }
                } catch (err) {
                    pplGrps.log("failed to show userPropertyValue preview", true);
                    pplGrps.log(err, true);
                }
            });

            configPopup.querySelector(".closeButton").addEventListener("click", function (e) {
                this.parentElement.parentElement.style.display = "none";
            });

            configPopup.querySelector(".saveButton").addEventListener("click", function (e) {
                pplGrps.log("Save button clicked");
                pplGrps.log(e);
                pplGrps.log(this);
                var popupWrapper = pplGrps.parentsUntilMatchingSelector(this, ".pplGrpsConfigPopup");
                var fieldName = popupWrapper.querySelector(".configForField").innerText;
                var saveConfig = {};
                saveConfig.fieldDisplayName = fieldName;
                saveConfig.bCaptureProfileDetails = false;
                saveConfig.arrMapping = [];


                var collMapping = this.parentElement.previousSibling.querySelectorAll(".mappingRow");
                for (var iR = 0; iR < collMapping.length; iR++) {
                    /*try{console.log(collMapping[iR]);}catch(err){}*/
                    var map = collMapping[iR];
                    var formFieldGUID = map.querySelector("SELECT.listField").value;
                    var formFieldName = document.querySelector("OPTION[value='" + formFieldGUID + "']").innerText;
                    //var userProperty = map.querySelector("SELECT.userProperty").value;
                    var userProperty = map.querySelector("INPUT.userProperty").value;
                    pplGrps.log("Capture user profile property |" + userProperty + "| in field named |" + formFieldName + "|");
                    /*
                    for ( formWP in pplGrps.formFields ) {
                        for ( formField in pplGrps.formFields[formWP] ) {
                            if ( typeof(pplGrps.formFields[formWP][formField].TopLevelElementId) !== "undefined" ) {
                                if ( pplGrps.formFields[formWP][formField].TopLevelElementId === control ) {
                                    arrH.push(pplGrps.formFields[formWP][formField].Title);
                                    break;
                                }
                            }    
                        }
                    }
                    */
                    saveConfig.arrMapping.push({
                        "fin": formFieldName,
                        "userProperty": userProperty
                    });
                    saveConfig.bCaptureProfileDetails = true;
                }
                popupWrapper.querySelector(".saveSettings").innerText = JSON.stringify(saveConfig);
                /* update the existing settings record for this field */
                var bFoundExisting = false;
                for (var iSettings = 0; iSettings < pplGrps.settings.length; iSettings++) {
                    //pplGrps.log(pplGrps.settings[iSettings],true);
                    if (pplGrps.settings[iSettings].fieldName === fieldName) {
                        //pplGrps.log(pplGrps.settings[iSettings].Id,true);
                        bFoundExisting = true;
                        pplGrps.ajax.updateSettings(pplGrps.settings[iSettings].Id, JSON.parse(popupWrapper.querySelector(".saveSettings").innerText));
                        break;
                    }
                }
                /* save new settings item if an existing one was not found */
                if (bFoundExisting === false) {
                    pplGrps.ajax.captureSettings(_spPageContextInfo.serverRequestPath, fieldName, JSON.parse(popupWrapper.querySelector(".saveSettings").innerText));
                }
            });

            configPopup.querySelector(".removeMapping").addEventListener("click", function (e) {
                pplGrps.log("remove");
                pplGrps.log(e);
                pplGrps.log(this);
                var parentRow = pplGrps.parentsUntilMatchingSelector(this, ".mappingRow");
                pplGrps.log(parentRow);
                parentRow.remove();
            });

            function fxAddNew(e) {
                pplGrps.log("AddNew button clicked");
                pplGrps.log(e);
                pplGrps.log(this);
                var newRow = pplGrps.parentsUntilMatchingSelector(this, ".mappingRow:last-of-type").cloneNode(true);
                //newRow.querySelector("SELECT.listField OPTION[value='0']").setAttribute("selected","selected");
                newRow.querySelector("SELECT.listField").value = 0;
                //newRow.querySelector("SELECT.userProperty OPTION[value='0']").setAttribute("selected","selected");
                newRow.querySelector("INPUT.userProperty").value = "";
                pplGrps.parentsUntilMatchingSelector(this, "[id*='_configMappingRowsWrapper']").appendChild(newRow);
                //configPopup.querySelector(".mappingRow:last-of-type .addNew").addEventListener("click",fxAddNew);
                pplGrps.parentsUntilMatchingSelector(this, "[id*='_configMappingRowsWrapper']").querySelector(".mappingRow:last-of-type .addNew").addEventListener("click", fxAddNew);
                pplGrps.parentsUntilMatchingSelector(this, "[id*='_configMappingRowsWrapper']").querySelector(".mappingRow:last-of-type .removeMapping").addEventListener("click", function (e) {
                    pplGrps.log("remove");
                    pplGrps.log(e);
                    pplGrps.log(this);
                    var parentRow = pplGrps.parentsUntilMatchingSelector(this, ".mappingRow");
                    pplGrps.log(parentRow);
                    parentRow.remove();
                });
                pplGrps.parentsUntilMatchingSelector(this, "[id*='_configMappingRowsWrapper']").querySelector(".mappingRow:last-of-type INPUT.userProperty").addEventListener("keyup", function (e) {
                    pplGrps.log("preview");
                    pplGrps.log(e);
                    pplGrps.log(this);
                    try {
                        if (this.value.indexOf("|") < 0) {
                            if (typeof (eval(this.value)) === "object") {
                                if (typeof (eval(this.value + ".length")) !== "undefined") {
                                    this.parentElement.parentElement.querySelector("SPAN.userPropertyValuePreview").innerText = "array";
                                } else {
                                    this.parentElement.parentElement.querySelector("SPAN.userPropertyValuePreview").innerText = eval(this.value);
                                }
                            } else {
                                this.parentElement.parentElement.querySelector("SPAN.userPropertyValuePreview").innerText = eval(this.value);
                            }
                        } else {
                            if (typeof (eval(this.value.split("|")[0] + "[0]." + this.value.split("|")[1])) === "object") {
                                if (typeof (eval(this.value.split("|")[0] + "[0]." + this.value.split("|")[1] + ".length")) !== "undefined") {
                                    this.parentElement.parentElement.querySelector("SPAN.userPropertyValuePreview").innerText = "array";
                                } else {
                                    this.parentElement.parentElement.querySelector("SPAN.userPropertyValuePreview").innerText = eval(this.value.split("|")[0] + "[0]." + this.value.split("|")[1]);
                                }
                            } else {
                                this.parentElement.parentElement.querySelector("SPAN.userPropertyValuePreview").innerText = eval(this.value.split("|")[0] + "[0]." + this.value.split("|")[1]);
                            }

                        }
                    } catch (err) {
                        pplGrps.log("failed to show userPropertyValue preview", true);
                        pplGrps.log(err, true);
                    }
                });
            };


            configPopup.querySelector(".addNew").addEventListener("click", fxAddNew);


        }
        if (typeof (afterFx) === "function") {
            afterFx();
        }
        //pplGrps.populateCurrentSettings();
    },
    appendConfigStyles: function () {
        var arrStyles = [".pplGrpsConfigPopup {display: none; background-color: #e6e6e6; border: 1px solid black; padding: 5px 10px; margin-top: 3px; width: fit-content;}"];
        arrStyles.push(".configPopupHeader {text-align: right;}");
        arrStyles.push(".configPopupFooter {margin-top: 3px;}");
        arrStyles.push(".buttonish {cursor: pointer; border: 1px solid black; padding: 1px 10px;}");
        arrStyles.push(".closeButton {font-size: 2.3em; border: 0px hidden transparent; padding: 0px 0px;}");
        arrStyles.push(".addNew {background-color: #40588e; color: #ffeee7; display: none;}");
        arrStyles.push(".removeMapping {background-color: #691818; color: #f1f1f1;}")
        arrStyles.push(".saveButton {background-color: #4e864d; color: #f1f1f1;}");
        arrStyles.push(".saveSettings {border: 3px double black; font-family: monospace; padding: 5px; margin-top: 8px;}");
        arrStyles.push(".pplGrpsConfigButton {background-color: #e6e6a6; color: brown; border: 0px hidden transparent; padding-top: 3px; padding-bottom: 3px;}");
        arrStyles.push("[id*='_configMappingRowsWrapper'] > tr.mappingRow:last-of-type .addNew {display: inline-block;}");
        arrStyles.push("INPUT.userProperty {width: 280px;}");
        var elemStyles = document.createElement("STYLE");
        elemStyles.type = "text/css";
        elemStyles.innerHTML = arrStyles.join("");
        document.head.appendChild(elemStyles);
    },
    parentsLoop: {
        current: null,
        parent: null,
        lastParent: null,
        selector: "",
        breakCounter: 0
    },
    parentsUntilMatchingSelector: function (htmlElement, selector) {
        pplGrps.parentsLoop.current = htmlElement;
        var parent = htmlElement.parentElement;
        var lastParent = parent;
        var iBreak = 0;
        pplGrps.parentsLoop.selector = selector;
        while (iBreak < 100 && parent.querySelectorAll(selector).length < 1 && parent.tagName !== "BODY" ) {
            pplGrps.parentsLoop.breakCounter = iBreak;
            pplGrps.parentsLoop.selector = selector;
            lastParent = parent;
            pplGrps.parentsLoop.lastParent = parent;
            parent = parent.parentElement;
            pplGrps.parentsLoop.parent = parent.parentElement;
            iBreak++;
        }
        if (parent.querySelectorAll(selector).length > 0) {
            return lastParent;
        } else {
            return false;
        }
    },
    setupListeners: function () {
        for (var iL = 0; iL < pplGrpsListeners.length; iL++) {
            pplGrps.log("configuring listener |"+ iL +"|",true);
            var listener = pplGrpsListeners[iL];
            var pickerFormField = pplGrps.findFormField("Title", listener.fieldDisplayName);
            if (!pickerFormField === false) {
                if (listener.bCaptureProfileDetails === true) {
                    var fxGenResolvedScript = function(oListener, oPickerFormField){
                        pplGrps.instances[oPickerFormField.TopLevelElementId].spcsom.OnUserResolvedClientScript = function (pickerControlId, arrUsers) {
                            var pickerTopLevelElementId = pickerControlId;
                            var pickerInstance = pplGrps.instances[pickerTopLevelElementId];
                            pplGrps.log("pplGrps.instances["+pickerControlId+"].spcsom.OnUserResolvedClientScript function fired!",true);
                            pplGrps.log("pplGrps.instances["+pickerControlId+"].spcsom.HasResolvedUsers() = |"+ pickerInstance.spcsom.HasResolvedUsers() +"|",true);
                            pplGrps.log("pplGrps.instances["+pickerControlId+"].spcsom.GetAllUserInfo().length = |"+ pickerInstance.spcsom.GetAllUserInfo().length +"|");
                            pplGrps.log(arrUsers,true);
                            pplGrps.log(pickerInstance.spcsom.GetAllUserInfo());
                            if (pickerInstance.spcsom.HasResolvedUsers() === false) {
                                pplGrps.log("pplGrps.instances["+pickerControlId+"] has no resolved users!",true);
                                
                            }
                            else if (pickerInstance.spcsom.HasResolvedUsers() === true) {
                                for (var iUser = 0; iUser < arrUsers.length; iUser++) {
                                    /* capture the user profile properties exposed via the person/group picker control */
                                    pickerInstance.users[arrUsers[iUser].Key] = arrUsers[iUser];
                                    pplGrps.log(arrUsers[iUser]);
                                    /* make a rest call to request user properties from the user's sharepoint profile */
                                    pplGrps.getUserPropertiesByLoginName(arrUsers[iUser].Key, function (oUser) {
                                        pplGrps.log(oUser);
                                        for (var iMapping = 0; iMapping < oListener.arrMapping.length; iMapping++) {
                                            var setField = pplGrps.findFormField("Name", oListener.arrMapping[iMapping].fin);
                                            pplGrps.log(setField,true);
                                            var setFieldElement = document.querySelector("[id*='" + setField.Id + "']");
                                            if ( setFieldElement === null ){
                                                switch (setField.FieldType){
                                                    case "SPFieldDateTime":
                                                        setFieldElement = document.querySelector(".ms-formbody INPUT[id$='DateTimeField_DateTimeFieldDate'][title^='"+ setField.Name +"']");
                                                        break;
                                                    case "SPFieldNote":
                                                        setFieldElement = document.querySelector(".ms-formbody TEXTAREA[title^='"+ setField.Name +"']");
                                                        break;
                                                    case "SPFieldChoice":
                                                        break;
                                                    case "SPFieldMultiChoice":
                                                        break;
                                                    case "SPFieldLookup":
                                                        break;
                                                    case "SPFieldLookupMulti":
                                                        break;
                                                    case "SPFieldAllDayEvent":
                                                        break;
                                                    case "SPFieldRecurrence":
                                                        break;
                                                    case "SPFieldOverbook":
                                                        break;
                                                    case "SPFieldFreeBusy":
                                                        break;
                                                    case "SPFieldFacilities":
                                                        break;
                                                    case "SPFieldUser":
                                                        setFieldElement = document.querySelector(".ms-formbody DIV.sp-peoplepicker-topLevel[title^='"+ setField.Name +"'], .ms-formbody DIV.ms-inputuserfield[title^='"+ setField.Name +"']");
                                                        break;
                                                    case "SPFieldUserMulti":
                                                        setFieldElement = document.querySelector(".ms-formbody DIV.sp-peoplepicker-topLevel[title^='"+ setField.Name +"'], .ms-formbody DIV.ms-inputuserfield[title^='"+ setField.Name +"']");
                                                        break;
                                                    default: 
                                                        setFieldElement = document.querySelector(".ms-formbody INPUT[title^='"+ setField.Name +"']");
                                                }
                                            }
                                            pplGrps.log(setFieldElement);
                                            pplGrps.log(setFieldElement.tagName);
                                            pplGrps.log(setFieldElement.className);
                                            pplGrps.log(oListener.arrMapping[iMapping].userProperty);
                                            var setUserProperty = undefined,
                                                splitUserProperty = undefined,
                                                setUserSubProperty = undefined;
                                            try {
                                                if (oListener.arrMapping[iMapping].userProperty.indexOf("|") >= 0) {
                                                    splitUserProperty = oListener.arrMapping[iMapping].userProperty.split("|");
                                                    setUserProperty = eval(splitUserProperty[0]);
                                                    setUserSubProperty = splitUserProperty[1];
                                                } else {
                                                    setUserProperty = eval(oListener.arrMapping[iMapping].userProperty);
                                                }
                                            } catch (err) {}
                                            pplGrps.log(setUserProperty);
                                            pplGrps.log(splitUserProperty);
                                            pplGrps.log(setUserSubProperty);
                                            /* capture the sharepoint user profile details retrieved via rest under this picker instance's user object */
                                            for ( userProp in oUser ){
                                                pickerInstance.users[oUser.LoginName][userProp] = oUser[userProp];
                                            }
                                            try {
                                                if (typeof (setUserProperty) === "object") {
                                                    if (typeof (setUserProperty.length) !== "undefined" && setUserProperty.length > 0) {
                                                        pplGrps.log("Setting userproperty mapping by looping through the referenced JS array");
                                                        pplGrps.log(setUserProperty);
                                                        for (var iG = 0; iG < setUserProperty.length; iG++) {
                                                            if (typeof (setUserSubProperty) !== "undefined") {
                                                                pplGrps.log("Setting userproperty mapping by looping through sub-array");
                                                                var subLoop = eval("setUserProperty[" + iG + "]." + setUserSubProperty);
                                                                pplGrps.log(subLoop);
                                                                pplGrps.emptyPicker(pplGrps.findFormField("Name", oListener.arrMapping[iMapping].fin).TopLevelElementId);
                                                                for (var iU = 0; iU < subLoop.length; iU++) {
                                                                    pplGrps.log("pplGrps adding person or group to field with fin |" + oListener.arrMapping[iMapping].fin + "|", true);
                                                                    /*SPClientPeoplePicker.SPClientPeoplePickerDict[pplGrps.findFormField("Name", oListener.arrMapping[iMapping].fin).TopLevelElementId].AddUserKeys(subLoop[iU].LoginName);*/
                                                                    //pickerInstance.spcsom.AddUserKeys(setUserProperty[iG].LoginName);
                                                                    pplGrps.instances[pplGrps.findFormField("Name", oListener.arrMapping[iMapping].fin).TopLevelElementId].spcsom.AddUserKeys(subLoop[iU].LoginName);
                                                                }

                                                            } else {
                                                                pplGrps.log("pplGrps adding person or group to field with fin |" + oListener.arrMapping[iMapping].fin + "|", true);
                                                                /*SPClientPeoplePicker.SPClientPeoplePickerDict[pplGrps.findFormField("Name", oListener.arrMapping[iMapping].fin).TopLevelElementId].AddUserKeys(setUserProperty[iG].LoginName);*/
                                                                pplGrps.emptyPicker(pplGrps.findFormField("Name", oListener.arrMapping[iMapping].fin).TopLevelElementId);
                                                                pplGrps.instances[pplGrps.findFormField("Name", oListener.arrMapping[iMapping].fin).TopLevelElementId].spcsom.AddUserKeys(setUserProperty[iG].LoginName);
                                                                //pickerInstance.spcsom.AddUserKeys(setUserProperty[iG].LoginName);
                                                            }
                                                        }
                                                    }
                                                }
                                            } catch (err) {}
                                            if (typeof (setUserProperty) !== "object") {
                                                pplGrps.log("pplGrps capturing simple user profile property value in form field with fin |" + oListener.arrMapping[iMapping].fin + "|",true);
                                                pplGrps.log(setFieldElement,true);
                                                if (setFieldElement.tagName === "INPUT") {
                                                    setFieldElement.value = eval(oListener.arrMapping[iMapping].userProperty);
                                                } 
                                                else if (setFieldElement.tagName === "SELECT") {
                                                    if (!setFieldElement.getAttribute("multiple") === true) {
                                                        /* single value select control */
                                                        setFieldElement.value = eval(oListener.arrMapping[iMapping].userProperty);
                                                    } else {
                                                        /* multivalue select control */
                                                    }
                                                } 
                                                else if (setFieldElement.tagName === "TEXTAREA") {
                                                    setFieldElement.value = eval(oListener.arrMapping[iMapping].userProperty);
                                                }
                                                else if (setFieldElement.tagName === "DIV" /*&& setFieldElement.className.indexOf("sp-peoplepicker-topLevel") >= 0*/ ) {
                                                    /* picker */
                                                    pplGrps.emptyPicker(setFieldElement.id);
                                                    pplGrps.instances[setFieldElement.id].spcsom.AddUserKeys(eval(oListener.arrMapping[iMapping].userProperty));
                                                    
                                                }
                                                else if (setFieldElement.tagName === "SPAN" && setFieldElement.className.indexOf("ms-usereditor") >= 0 ) {
                                                    //pplGrps.emptyPicker(setFieldElement.id);
                                                    pplGrps.legacyInstances[setFieldElement.id+"_upLevelDiv"].spcsom.AddUserKeys(eval(oListener.arrMapping[iMapping].userProperty));
                                                }
                                            }
                                        }
                                    });
                                }
                            }
                        }
                    }
                    fxGenResolvedScript(listener, pickerFormField);
                    pplGrps.log("set instance's spcsom.OnUserResolvedClientScript for |"+ pickerFormField.TopLevelElementId +"|",true);
                }
            }
        }
    },
    emptyPicker: function(instanceName){
        var oPicker = pplGrps.instances[instanceName].spcsom;
        var collUsers = oPicker.GetAllUserInfo();
        for ( var iUser = 0; iUser < collUsers.length; iUser++ ) {
            oPicker.DeleteProcessedUser(document.getElementById(oPicker.TopLevelElementId +"_ResolvedList").childNodes[0])
        }
    },
    populateCurrentSettings: function () {
        for (var iPL = 0; iPL < pplGrpsListeners.length; iPL++) {
            var listener = pplGrpsListeners[iPL];
            pplGrps.log(listener);
            var pickerFormField = pplGrps.findFormField("Title", listener.fieldDisplayName);
            if (!pickerFormField === false) {
                var pickerTopLevelElementId = pickerFormField.TopLevelElementId;
                var pickerInstance = pplGrps.instances[pickerTopLevelElementId];
                pickerInstance.arrMapping = listener.arrMapping;
                var pickerConfigWrapper = document.getElementById(pickerTopLevelElementId + "_configWrapper");
                if (listener.bCaptureProfileDetails === true) {
                    var pickerConfigMappingTBODY = document.getElementById(pickerTopLevelElementId + "_configMappingRowsWrapper");
                    pplGrps.log(pickerConfigMappingTBODY,true);
                    var mappingRow = pickerConfigMappingTBODY.querySelector(".mappingRow");
                    for (var iMapping = 0; iMapping < listener.arrMapping.length; iMapping++) {
                        pplGrps.log(listener.arrMapping[iMapping]);
                        var setField = pplGrps.findFormField("Name", listener.arrMapping[iMapping].fin);
                        pplGrps.log(setField);
                        var setFieldElement = document.querySelector("[id*='" + setField.Id + "']");
                        pplGrps.log(setFieldElement);
                        var userPropertyName = listener.arrMapping[iMapping].userProperty;
                        pplGrps.log(userPropertyName);
                        /* set the values of our select controls that define our mapping */
                        if (iMapping > 0) {
                            pickerConfigMappingTBODY.querySelector(".mappingRow:last-of-type .addNew").click();
                            var newRow = pickerConfigMappingTBODY.querySelector(".mappingRow:last-of-type");
                            try {
                                newRow.querySelector("SELECT.listField OPTION[value='" + setField.Id + "']").setAttribute("selected", "selected");
                            } catch (err) {}
                            //try{newRow.querySelector("SELECT.userProperty OPTION[value='"+ userPropertyName +"']").setAttribute("selected","selected");}catch(err){}
                            try {
                                newRow.querySelector("INPUT.userProperty").value = userPropertyName;
                            } catch (err) {}
                            if (userPropertyName.length > 0) {
                                var e = new Event("keyup");
                                e.which = 16;
                                newRow.querySelector("INPUT.userProperty").dispatchEvent(e);
                            }
                        } else {
                            try {
                                mappingRow.querySelector("SELECT.listField OPTION[value='" + setField.Id + "']").setAttribute("selected", "selected");
                            } catch (err) {}
                            //try{mappingRow.querySelector("SELECT.userProperty OPTION[value='"+ userPropertyName +"']").setAttribute("selected","selected");}catch(err){}
                            try {
                                mappingRow.querySelector("INPUT.userProperty").value = userPropertyName;
                            } catch (err) {}
                            if (userPropertyName.length > 0) {
                                var e = new Event("keyup");
                                e.which = 16;
                                mappingRow.querySelector("INPUT.userProperty").dispatchEvent(e);
                            }
                        }
                    }
                }
            }
        }
    },
    showAvailableProperties: function () {
        var propsWrapper = document.createElement("ul");
        propsWrapper.setAttribute("id", "propsWrapper");
        for (prop in pplGrps.userProperties) {
            pplGrps.log("[" + prop + "] = |" + JSON.stringify(pplGrps.userProperties[prop]) + "|", true);
            var propOut = document.createElement("li");
            propOut.setAttribute("id", "propOut" + prop);
            propOut.innerText = "[" + prop + "] = |" + JSON.stringify(pplGrps.userProperties[prop]) + "|";
            propsWrapper.appendChild(propOut);
        }
        document.getElementById("DeltaPlaceHolderMain").insertAdjacentElement('beforebegin', propsWrapper);
        propsWrapper.addEventListener("click", function (e) {
            this.remove();
        });
    },
    loadScriptsForLegacyPickerConversion: function(afterFx){
        /*
        _spPageContextInfo.siteAbsoluteUrl +"/_layouts/15/clienttemplates.js"
        _spPageContextInfo.siteAbsoluteUrl +"/_layouts/15/clientforms.js"
        _spPageContextInfo.siteAbsoluteUrl +"/_layouts/15/clientpeoplepicker.js"
        _spPageContextInfo.siteAbsoluteUrl +"/_layouts/15/autofill.js"
        */
       var arrSPScripts = [_spPageContextInfo.siteAbsoluteUrl +"/_layouts/15/clienttemplates.js",_spPageContextInfo.siteAbsoluteUrl +"/_layouts/15/clientforms.js",_spPageContextInfo.siteAbsoluteUrl +"/_layouts/15/clientpeoplepicker.js",_spPageContextInfo.siteAbsoluteUrl +"/_layouts/15/autofill.js"];
       for ( var iSPS = 0; iSPS < arrSPScripts.length; iSPS++){
           var scriptElem = document.createElement("SCRIPT");
           scriptElem.src = arrSPScripts[iSPS];
           document.head.appendChild(scriptElem);
       }
       if ( typeof(afterFx) === 'function' ){
           afterFx();
       }
       return true;
    },
    getFieldDefinition: function(fieldGUID, fxCallback){
        /*pplGrps.getFieldDefinition(pplGrps.legacyInstances["ctl00_ctl26_g_e1c47bb7_99a4_482b_9c13_94ae882b8a90_ctl00_ctl05_ctl13_ctl00_ctl00_ctl04_ctl00_ctl00_UserField_upLevelDiv"].fieldGUID);*/
        pplGrps.ajax.read(
            _spPageContextInfo.webAbsoluteUrl+"/_api/web/lists(guid'"+ _spPageContextInfo.pageListId.replace("{","").replace("}","") +"')/Fields('"+ fieldGUID +"')", 
            function(x,d){
                pplGrps.log(d,true);
                if ( typeof(fxCallback) === 'function' ) { 
                    fxCallback(d.d); 
                }
            },
            undefined,
            function(){
                pplGrps.log("Failed");
            },
            "json"
        );
    },
    convertLegacyPicker: function(peoplePickerElementId, oField, defaultValue){
        /*
        pplGrps.getFieldDefinition(
            pplGrps.legacyInstances["ctl00_ctl26_g_e1c47bb7_99a4_482b_9c13_94ae882b8a90_ctl00_ctl05_ctl09_ctl00_ctl00_ctl04_ctl00_ctl00_UserField_upLevelDiv"].fieldGUID, function(oField){
            pplGrps.convertLegacyPicker(
                "ctl00_ctl26_g_e1c47bb7_99a4_482b_9c13_94ae882b8a90_ctl00_ctl05_ctl09_ctl00_ctl00_ctl04_ctl00_ctl00_UserField_upLevelDiv", 
                oField, 
                null
            );
        });
        */
        if ( typeof(defaultValue) === "undefined" ){ var defaultValue = null; }
        /*
        if ( typeof(SPClientPeoplePicker) === "undefined" ) {
            var elem = document.createElement("DIV");
            elem.className = "sp-peoplepicker-topLevel";
            elem.id = oField.InternalName+ "_"+ oField.Id +"_$ClientPeoplePicker";
            elem.setAttribute("Title",oField.Title);
            elem.setAttribute("spclientpeoplepicker","true");
            document.getElementById(peoplePickerElementId).insertAdjacentElement('beforebegin',elem);
        
            SP.SOD.executeFunc('clientpeoplepicker.js', 'SPClientPeoplePicker', function(){
                var schema = {};
                schema['PrincipalAccountType'] = 'User,DL,SecGroup,SPGroup';
                schema['SearchPrincipalSource'] = 15;
                schema['ResolvePrincipalSource'] = 15;
                schema['AllowMultipleValues'] = oField.AllowMultipleValues;
                schema['MaximumEntitySuggestions'] = 50;
                schema['Width'] = document.getElementById(peoplePickerElementId).offsetWidth+'px';
                SPClientPeoplePicker.InitializeStandalonePeoplePicker(elem.id, defaultValue, schema);
                document.getElementById(peoplePickerElementId).style.display = "none";
            });
        }
        */
       var schema = {};
       schema['PrincipalAccountType'] = 'User,DL,SecGroup,SPGroup';
       schema['SearchPrincipalSource'] = 15;
       schema['ResolvePrincipalSource'] = 15;
       schema['AllowMultipleValues'] = oField.AllowMultipleValues;
       schema['MaximumEntitySuggestions'] = 50;
       schema['Width'] = document.getElementById(peoplePickerElementId).offsetWidth+'px';
       SPClientPeoplePicker.InitializeStandalonePeoplePicker(peoplePickerElementId, defaultValue, schema);
    },
    onSharePointReady: function () {
        pplGrps.log("SP.ClientContext detected via SP.SOD.executeFuc... fired dsU.onSharePointReady", true);
        pplGrps.getUserProperties(function () {
            pplGrps.getPageWebPartsAndFormFields(function (wp) {
                pplGrps.log("webpart detected");
                pplGrps.log(wp);
            }, function () {
                pplGrps.getAllPickers();
                pplGrps.appendConfigStyles();
            });
            if (pplGrps.isPageInEditMode() === false) {
                pplGrps.ajax.checkIfSettingsListExists(function (userCanEditSettings) {
                    if (userCanEditSettings === true) {
                        pplGrps.showConfigButtons(function () {
                            pplGrps.ajax.getSettings(function () {
                                pplGrps.log("Got dsMagic-pplGrps settings for this page");
                                pplGrps.setupListeners();
                                pplGrps.populateCurrentSettings();
                            });
                        });
                    } else {
                        pplGrps.ajax.getSettings(function () {
                            pplGrps.log("Got dsMagic-pplGrps settings for this page");
                            pplGrps.setupListeners();
                            //pplGrps.populateCurrentSettings();
                        });
                    }
                    pplGrps.log("Settings list already exists for dsMagic-pplGrps JS API");

                }, function () {
                    pplGrps.ajax.createSettingsList(function () {
                        pplGrps.log("Settings list created for dsMagic-pplGrps JS API", true);
                    });
                })
            }
        });
    },
    onLoad: setTimeout(function () {
        pplGrps.log("dsMagic-pplGrps Namespace loaded", true);
        if (document.readyState !== "complete") {
            document.onreadystatechange = function () {
                if (document.readyState === "complete") {
                    pplGrps.log("document.onreadystatechange event fired and document.readyState is 'complete'", true);
                    SP.SOD.executeFunc('sp.js', 'SP.ClientContext', pplGrps.onSharePointReady);
                }
            }
        } else {
            pplGrps.log("document.onreadystatechange event fired and document.readyState is 'complete'", true);
            SP.SOD.executeFunc('sp.js', 'SP.ClientContext', pplGrps.onSharePointReady);
        }
    }, 23)
}