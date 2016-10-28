/* nSPGetListItems
 * ----------------------------------------------------------------------------
 * a JavaScript convenience library to get items from SharePoint lists using CSOM/JSOM
 *
 * version : 1.0
 * url     : https://github.com/imthenachoman/nSPGetListItems
 * author  : Anchal Nigam
 * e-mail  : imthenachoman@gmail.com
 *
 * Copyright (c) 2015 Anchal Nigam (imthenachoman@gmail.com)
 * Licensed under the MIT license: http://opensource.org/licenses/MIT
 * ----------------------------------------------------------------------------
 */
var nSPGetListItems = nSPGetListItems || (function()
{
    // track the user's requests
    var instanceTracker = [];
    
    // conversion functions for SP data objects
    var fieldTypeValueFunctions = {
        "UserMulti" : function(row, fieldData, listData)
        {
            // user fields are the same as normal lookup fields
            return fieldTypeValueFunctions["LookupMulti"](row, fieldData, listData);
        },
        "LookupMulti" : function(row, fieldData, listData)
        {
            // core data
            var ret = {
                "name" : fieldData.displayName,
                "type" : fieldData.typeName,
                "lookupType" : fieldData.lookupValueType,
                "data" : null
            };
            
            // if we have data iterate through them
            var data = row[fieldData.internalName];
            if(data && data.length)
            {
                ret.data = [];
                for(var i = 0, numLookups = data.length; i < numLookups; ++i)
                {
                    ret.data.push({
                        "value"        : fieldData.lookupValueOf(data[i].get_lookupValue()),
                        "lookupListID" : data[i].get_lookupId(),
                    });
                }
            }
            
            return ret;
        },
        "User" : function(row, fieldData, listData)
        {
            // user fields are the same as normal lookup fields
            return fieldTypeValueFunctions["Lookup"](row, fieldData, listData);
        },
        "Lookup" : function(row, fieldData, listData)
        {
            return {
                "name" : fieldData.displayName,
                "type" : fieldData.typeName,
                "lookupType" : fieldData.lookupValueType || "unknown",
                "data" : row[fieldData.internalName] ? {
                    "value"        : fieldData.lookupValueOf(row[fieldData.internalName].get_lookupValue()),
                    "lookupListID" : row[fieldData.internalName].get_lookupId()
                } : null
            };
        },
        "Title" : function(row, fieldData, listData)
        {
            return {
                "name" : fieldData.displayName,
                "type" : "Title",
                "data" : row["Title"]
            };
        },
        "FileType" : function(row, fieldData, listData)
        {
            return {
                "name" : fieldData.displayName,
                "type" : fieldData.typeName,
                "data" : row.FSObjType == "1" ? "folder" : (row.File_x0020_Type || "file")
            };

        },
        "ContentTypeID" : function(row, fieldData, listData)
        {
            return {
                "name" : fieldData.displayName,
                "type" : fieldData.typeName,
                "data" : row.ContentTypeId.toString()
            };
        },
        "File" : function(row, fieldData, listData)
        {
            return {
                "name" : fieldData.displayName,
                "type" : "File",
                "data" : {
                    "fileName" : row["FileLeafRef"],
                    "fileDirectory" : row["FileDirRef"],
                    "filePath" : row["FileRef"]
                },
            };
        },
        "URL" : function(row, fieldData, listData)
        {
            return {
                "name" : fieldData.displayName,
                "type" : fieldData.typeName,
                "data" : {
                    "url"         : row[fieldData.internalName].get_url(),
                    "description" : row[fieldData.internalName].get_description() 
                }
            };
        },
        "Boolean" : function(row, fieldData, listData)
        {
            return {
                "name" : fieldData.displayName,
                "type" : fieldData.typeName,
                "data" : row[fieldData.internalName] === true || row[fieldData.internalName] === "1"
            };

        },
        "FileSize" : function(row, fieldData, listName)
        {
            return {
                "name" : fieldData.displayName,
                "type" : fieldData.typeName,
                "data" : row["File_x0020_Size"] ? parseFloat(row["File_x0020_Size"]) : null
            };
        },
        "Number" : function(row, fieldData, listData)
        {
            return {
                "name" : fieldData.displayName,
                "type" : fieldData.typeName,
                "data" : parseFloat(row[fieldData.internalName])
            };

        },
        "default" : function(row, fieldData, listData)
        {
            return {
                "name" : fieldData.displayName,
                "type" : fieldData.typeName,
                "data" : row[fieldData.internalName]
            };
        },
        // special function for lookup fields
        "LookupFieldDateTime" : function(value)
        {
            // convert a UTC date string to a JS date object
            var utcDateParts = value.split(/-|t|:|z/ig);
            utcDateParts[1]--;
            var date = new Date(utcDateParts[0], utcDateParts[1], utcDateParts[2], utcDateParts[3], utcDateParts[4], utcDateParts[5]);
            return new Date(date.getTime() - (date.getTimezoneOffset() * 60 * 1000));
        },
        "LookupFieldNumber" : function(value)
        {
            return parseFloat(value);
        },
        "LookupFieldDefault" : function(value)
        {
            return value;
        }
    };
    
    // needed change the field type for certain fields
    var mySPFieldTypeMappings = {
        "ID_Counter" : "Number",
        "Edit_Computed" : false,
        "DocIcon_Computed" : "FileType",
        "LinkFilenameNoMenu_Computed" : "File",
        "LinkFilename_Computed" : "File",
        "ContentType_Computed" : "ContentTypeID",
        "FileSizeDisplay_Computed" : "FileSize",
        "ItemChildCount_Lookup" : "Number",
        "FolderChildCount_Lookup" : "Number",
        "Title_Text" : "Title",
        "LinkTitleNoMenu_Computed" : "Title",
        "LinkTitle_Computed" : "Title",
        "Attachments_Attachments" : "Boolean",
        "_UIVersionString_Text" : "Number",
    };
    
    // the main call
    var main = function(options)
    {
        // we need options
        if(!options) return false;
        // we have to have a list name or a list GUID
        if(!("listName" in options) && !("listGUID" in options)) return false;
        
        // create a place to hold all the data
        var listData = {
            "ID"                     : instanceTracker.length,
            "listName"               : options.listName,
            "listGUID"               : options.listGUID,
            "viewName"               : options.viewName,
            "viewGUID"               : options.viewGUID,
            "webURL"                 : options.webURL,
            "onListLoadComplete"     : options.onListLoadComplete,
            "onViewLoadComplete"     : options.onViewLoadComplete,
            "onListItemLoadComplete" : options.onListItemLoadComplete,
            "onError"                : options.onError,
            "interval"               : options.interval || 0,
            "myData"                 : options.myData
        };
        
        // add it to the tracker
        instanceTracker.push(listData);
        
        // get the data
        ExecuteOrDelayUntilScriptLoaded(function()
        {
            // load the web URL we want
            listData.clientContext = listData.webURL ? new SP.ClientContext(listData.webURL) : SP.ClientContext.get_current();
            
            // get the list
            listData.list = options.listName ? listData.clientContext.get_web().get_lists().getByTitle(listData.listName) : listData.clientContext.get_web().get_lists().getById(listData.listGUID);
            
            // get the list content types
            listData.contentTypes = listData.list.get_contentTypes();
            
            // get the list fields
            listData.listFields = listData.list.get_fields();
            
            // get the view the user wants or the default view if the user didn't specify a view
            listData.view = listData.viewName ? listData.list.get_views().getByTitle(listData.viewName) : (listData.viewGUID ? listData.list.get_views().getById(listData.viewGUID) : listData.list.getView(""));
            
            // get the fields for the view
            listData.viewFields = listData.view.get_viewFields();
            
            // load everything
            listData.clientContext.load(listData.list);
            
            // for the list we want the effective permissions for the current user
            listData.clientContext.load(listData.list, "EffectiveBasePermissions");
            listData.clientContext.load(listData.contentTypes);
            listData.clientContext.load(listData.listFields);
            listData.clientContext.load(listData.view);
            // we want the ViewQuery so we can get the query ML
            listData.clientContext.load(listData.view, "ViewQuery");
            listData.clientContext.load(listData.viewFields);
            
            // we're at the current stage
            listData.stage = 1;
            
            // do it to it it
            // send this list data as the this context so we can reference it in the success/error functions
            listData.clientContext.executeQueryAsync(Function.createDelegate(listData, listDetailsLoaded), Function.createDelegate(listData, onError));
        }, "sp.js");
        
        return listData.ID;
    };
    
    // called by executeQueryAsync calls on errors 
    var onError = function(sender, args)
    {
        // get the list data
        var listData = this;
        
        // if the user has an error function then run it
        if(listData.onError) listData.onError(listData.stage, args.get_message(), listData.myData);
    };
    
    // run when a list is loaded
    var listDetailsLoaded = function(sender, args)
    {
        // get the list data
        var listData = this;
        
        // if the user wants to know when the list data is loaded
        if(listData.onListLoadComplete)
        {
            var listContentTypes = {}, spListContentTypes = listData.contentTypes.get_data();
            
            // get the list content types
            for(var i = 0, numContentTypes = spListContentTypes.length; i < numContentTypes; ++i)
            {
                var currentContentType = spListContentTypes[i];
                listContentTypes[currentContentType.get_id().toString()] = {
                    "name"        : currentContentType.get_name(),
                    "description" : currentContentType.get_description(),
                    "forms" : {
                        "new"     : currentContentType.get_newFormUrl(),
                        "edit"    : currentContentType.get_editFormUrl(),
                        "disp"    : currentContentType.get_displayFormUrl()
                    }
                };
            }
            
            // get the list permissions
            var listPermissions = listData.list.get_effectiveBasePermissions();
            
            // give the user the relevant list data
            listData.onListLoadComplete({
                "listName"     : listData.list.get_title(),
                "listGUID"     : listData.list.get_id().toString(),
                "viewName"     : listData.view.get_title(),
                "viewGUID"     : listData.view.get_id().toString(),
                "listPermissions"  : {
                    "add"      : listPermissions.has(SP.PermissionKind.addListItems),
                    "edit"     : listPermissions.has(SP.PermissionKind.editListItems),
                    "delete"   : listPermissions.has(SP.PermissionKind.deleteListItems),
                    "view"     : listPermissions.has(SP.PermissionKind.viewListItems),
                },
                "listForms"        : {
                    "new"      : listData.list.get_defaultNewFormUrl(),
                    "edit"     : listData.list.get_defaultEditFormUrl(),
                    "disp"     : listData.list.get_defaultDisplayFormUrl()
                },
                "contentTypes" : listContentTypes,
                "myData"       : listData.myData
            });
        }
        
        // get all the fields in the list and save them so we can reference them later
        var listFields = {}, fieldEnumerator = listData.listFields.getEnumerator();
        while(fieldEnumerator.moveNext())
        {
            var thisField = fieldEnumerator.get_current();
            listFields[thisField.get_internalName()] = thisField;
        }
        
        // get all the fields in the view
        var myViewFields = [], theirViewFields = [], viewLookupFieldsToGet = [], spViewFields = listData.viewFields.get_data();
        for(var i = 0, numFields = spViewFields.length; i < numFields; ++i)
        {
            // get the field internal name
            var thisViewFieldInternalName = spViewFields[i];
            
            // find the associated field data
            var thisViewFieldData = listFields[thisViewFieldInternalName];
            
            // get the field settings
            var fieldSettings = getFieldSettings(thisViewFieldData);
            
            // if we didn't get anything then skip
            if(fieldSettings === false) continue;
            
            // save the data for internal use
            var internalFieldSettings = {
                "typeName" : fieldSettings.typeName,
                "displayName" : fieldSettings.displayName,
                "internalName" : fieldSettings.internalName,
                "valueOf" : fieldTypeValueFunctions[fieldSettings.typeName] || fieldTypeValueFunctions["default"]
            };
            
            // for lookup and user fields we want to get the lookup list
            if(fieldSettings.typeName == "Lookup" || fieldSettings.typeName == "LookupMulti" || fieldSettings.typeName == "User" || fieldSettings.typeName == "UserMulti")
            {
                // the list GUID and field internal name
                var lookupList = thisViewFieldData.get_lookupList();
                var lookupListFieldInternalName = thisViewFieldData.get_lookupField();
                
                // set a default value function
                internalFieldSettings.lookupValueOf = fieldTypeValueFunctions["LookupFieldDefault"];
                
                // if we have a lookup list
                if(lookupList)
                {
                    // place to store source list information
                    fieldSettings.sourceList = {};
                    var lookupListData = {
                        "myFieldSettings" : internalFieldSettings,
                        "userFieldDataLookupFieldSettings" : fieldSettings.sourceList,
                        // get the list
                        "list" : listData.clientContext.get_web().get_lists().getById(lookupList)
                    };
                    
                    // if we have a field then get the field too
                    if(lookupListFieldInternalName) lookupListData.field = lookupListData.list.get_fields().getByInternalNameOrTitle(lookupListFieldInternalName);
                    
                    // add it to the lst of data we want to pull
                    viewLookupFieldsToGet.push(lookupListData);
                }
            }
            
            // save the field data
            myViewFields.push(internalFieldSettings);
            theirViewFields.push(fieldSettings);
        }
        
        // store all the field data
        listData.fields = myViewFields;
        listData.numFields = myViewFields.length;
        
        // get the view CAML
        var viewXML = listData.view.get_viewQuery(), camlQuery = new SP.CamlQuery();
        camlQuery.set_viewXml('<View><Query>' + viewXML + '</Query><RowLimit>0</RowLimit></View>');
        
        // if we have lookup fields we need to get it
        if(viewLookupFieldsToGet.length)
        {
            // load the 
            for(var i = 0, numLookups = viewLookupFieldsToGet.length; i < numLookups; ++i)
            {
                var lookupListData = viewLookupFieldsToGet[i];
                listData.clientContext.load(lookupListData.list);
                listData.clientContext.load(lookupListData.list, "EffectiveBasePermissions");
                if(lookupListData.field)
                {
                    listData.clientContext.load(lookupListData.field);
                }
            }
            
            listData.stage = 2;
            // make the call
            listData.clientContext.executeQueryAsync(Function.createDelegate(viewLookupFieldsToGet, function(sender, args)
            {
                // iterate through all the lookup lists and save the data
                for(var i = 0, numLookups = viewLookupFieldsToGet.length; i < numLookups; ++i)
                {
                    var lookupListData = viewLookupFieldsToGet[i];
                    var myFieldSettings = lookupListData.myFieldSettings;
                    var userFieldDataLookupFieldSettings = lookupListData.userFieldDataLookupFieldSettings;
                    var list = lookupListData.list;
                    var field = lookupListData.field;
                    
                    // get list permissions
                    var listPermissions = list.get_effectiveBasePermissions();
                    userFieldDataLookupFieldSettings.listPermissions = {
                        "add"    : listPermissions.has(SP.PermissionKind.addListItems),
                        "edit"   : listPermissions.has(SP.PermissionKind.editListItems),
                        "delete" : listPermissions.has(SP.PermissionKind.deleteListItems),
                        "view"   : listPermissions.has(SP.PermissionKind.viewListItems),
                    };
                    
                    // get list form url
                    userFieldDataLookupFieldSettings.listForms = {
                        "new"  : list.get_defaultNewFormUrl(),
                        "edit" : list.get_defaultEditFormUrl(),
                        "disp" : list.get_defaultDisplayFormUrl()
                    };
                    
                    // if we found the field then get its info
                    if(field)
                    {
                        userFieldDataLookupFieldSettings.sourceField = getFieldSettings(field);
                        myFieldSettings.lookupValueType = userFieldDataLookupFieldSettings.sourceField.typeName;
                        switch(userFieldDataLookupFieldSettings.sourceField.typeName)
                        {
                            case "Date":
                            case "DateTime":
                                myFieldSettings.lookupValueOf = fieldTypeValueFunctions["LookupFieldDateTime"]; 
                                break;
                            case "Number":
                                myFieldSettings.lookupValueOf = fieldTypeValueFunctions["LookupFieldNumber"];
                                break
                            default:
                                myFieldSettings.lookupValueOf = fieldTypeValueFunctions["LookupFieldDefault"];
                        }
                    }
                }
                
                // load the data
                if(listData.onViewLoadComplete) camlQuery = listData.onViewLoadComplete(theirViewFields, viewXML, camlQuery, listData.myData) || camlQuery;
                listData.camlQuery = camlQuery;
                listData.listItems = listData.list.getItems(listData.camlQuery);
                
                listData.stage = 3;
                refresh(listData.ID);
            }), Function.createDelegate(listData, onError));
        }
        // otherwise just load the data
        else
        {
            if(listData.onViewLoadComplete) camlQuery = listData.onViewLoadComplete(theirViewFields, viewXML, camlQuery, listData.myData) || camlQuery;
            listData.camlQuery = camlQuery;
            listData.listItems = listData.list.getItems(listData.camlQuery);
                        
            listData.stage = 3;
            refresh(listData.ID);
        }
    };
    
    var getFieldSettings = function(field)
    {
        // core field settings
        var fieldSettings = {
            "typeID"            : field.get_fieldTypeKind(),
            "typeName"          : field.get_typeAsString(),
            "internalName"      : field.get_internalName(),
            "displayName"       : field.get_title(),
            "staticName"        : field.get_staticName()
        };
        
        // we need to map certain types of fields based on known data
        var spFieldTypeMapping = mySPFieldTypeMappings[fieldSettings.internalName + "_" + fieldSettings.typeName];
        
        // we need to skip some fields because we don't know how to process them so return false if that is the case
        if(spFieldTypeMapping === false) return false;
        else if(spFieldTypeMapping) fieldSettings.typeName = spFieldTypeMapping
        
        // for DateTime fields we want to see if it is just date or date and time
        if(fieldSettings.typeName == "DateTime")
        {
            // if the date format is 0 then we're just date
            if(field.get_displayFormat)
            {
                if(field.get_displayFormat() == "0") 
                {
                    fieldSettings.typeName = "Date";
                }
            }
            // if we don't have the displayformat then we have to use the schema XML
            else if(field.get_schemaXml)
            {
                var dateFormat = field.get_schemaXml().match(/\bFormat="(.*?)"\s/);
                if(dateFormat && dateFormat[1] == "DateOnly")
                {
                    fieldSettings.typeName = "Date";
                }
            }
        }
        // for calculated fields we want to get the display type
        else if(fieldSettings.typeName == "Calculated")
        {
            switch(field.get_outputType ? field.get_outputType() : 2)
            {
                case 2:
                    fieldSettings.typeName = "Text";
                    break;
                case 9:
                    fieldSettings.typeName = "Number";
                    break;
                case 10:
                    fieldSettings.typeName = "Currency";
                    break;
                case 4:
                    if(field.get_dateFormat()) fieldSettings.typeName = "DateTime";
                    else fieldSettings.typeName = "Date";
                    break;
                case 8:
                    fieldSettings.typeName = "Boolean";
                    break;
            }
        }
        
        return fieldSettings;
    };
    
    // list items loaded
    var listItemsLoaded = function(sender, args)
    {
        var listData = this, rows = [], fields = listData.fields, numFields = listData.numFields, listItems = listData.listItems.get_data();
        // iterate through the returned items
        for(var i = 0, numRows = listItems.length; i < numRows; ++i)
        {
            // store the data for each row/item
            var row = [];
            var thisRowData = listItems[i].get_fieldValues();
            var thisRowPermissions = listItems[i].get_effectiveBasePermissions();
            
            for(var j = 0; j < numFields; ++j)
            {
                var thisField = fields[j];
                row.push(thisField.valueOf(thisRowData, thisField, listData));
            }
            
            rows.push(
            {
                "ID" : thisRowData.ID,
                "permissions" :
                {
                    "edit" : thisRowPermissions.has(SP.PermissionKind.editListItems),
                    "delete" : thisRowPermissions.has(SP.PermissionKind.deleteListItems)
                },
                "data" : row
            });
        }
        
        // if the user wants the data then give it to them
        if(listData.onListItemLoadComplete) listData.onListItemLoadComplete(rows, listData.myData);
        
        // if we want to set an interval then set it
        if(listData.interval) listData.timerID = setTimeout(function()
        {
            refresh(listData.ID);
        }, listData.interval);
    };
    
    // refresh the data
    var refresh = main.refresh = function(trackerID)
    {
        // get the list
        var listData = instanceTracker[trackerID];
        
        // clear any running timer
        clearTimeout(listData.timerID);
        
        // load the data
        listData.clientContext.load(listData.listItems);
        listData.clientContext.load(listData.listItems, "Include(EffectiveBasePermissions)");
        
        // get the data
        listData.clientContext.executeQueryAsync(Function.createDelegate(listData, listItemsLoaded), Function.createDelegate(listData, onError));
    };
    
    // change directory for doc libs
    main.cd = function(trackerID, newDirectory)
    {
        var listData = instanceTracker[trackerID];
        
        // if the new dir is .. then go up one level
        if(newDirectory == "..")
        {
            newDirectory = "";
            var currentDirectory = listData.camlQuery.get_folderServerRelativeUrl();
            if(currentDirectory)
            {
                newDirectory = currentDirectory.replace(/\/[^\/]+$/, "");
            }
        }
        
        if(newDirectory)
        {
            listData.camlQuery.set_folderServerRelativeUrl(newDirectory);
            listData.listItems = listData.list.getItems(listData.camlQuery);
            refresh(trackerID);
            return true;
        }
        return false;
    };
    
    // stop a timer
    main.stopInterval = function(trackerID)
    {
        var listData = instanceTracker[trackerID];
        clearTimeout(listData.timerID);
    };
    
    return main;
})();
