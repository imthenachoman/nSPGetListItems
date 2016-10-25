# nSPGetListItems

a JavaScript convenience library to get items from SharePoint lists using CSOM/JSOM

https://github.com/imthenachoman/nSPGetListItems

## Table of Contents

 1. [Overview](#overview)
 2. [Reference](#reference)
	 1. function [onListLoadComplete](#onlistloadcomplete)
	 2. function [onViewLoadComplete](#onviewloadcomplete)
	 3. function [onListItemLoadComplete](#onlistitemloadcomplete)
	 4. function [onError](#onerror)
 4. [Examples](#examples)
 
## Overview

When developing something for SharePoint the majority of the time I am getting items from a list or document library using JavaScript. I got tired of having to write the same code over and over again so I wrote the helper library/utility `nSPGetListItems`.

The idea is simple: call `nSPGetListItems` with your [options](#reference)  (a list or document library name at minimum), and get back details about the list, the view, and, of course, the list items.

Additionally, `nSPGetListItems` lets you refresh the data asynchronously so you don't have to code for that. 

## Reference

The main function is called `nSPGetListItems` and it takes one `key : value` object. The available `key`s are:

Key | Required | Type | Default Value | Description | Example
--- | --- | --- | --- | --- | ---
`listName` or `listGUID` | **yes** | string | | the name or GUID of the list | <ul><li>`"Announcements"`</li><li>`"{1c7c0498-6f1c-4ec1-8ee6-dd9959f3c52d}"`</li></ul>
`viewName` or `viewGUID` | | string | the list's default view | the name or GUID of the view you want to use to determine the fields and items to pull | <ul><li>`"All Documents"`</li><li>`"{dce68293-70e5-4c47-acda-72e6236b8f65}"`</li></ul>
`webURL` | | string | the current web | the web URL you want to pull from | <ul><li>`"/someSite"`</li><li>`"/someSite/subSite"`</li></ul>
`onListLoadComplete` | | function | n/a | the function to call when the list details are loaded | see [onListLoadComplete](#onlistloadcomplete) below
`onViewLoadComplete` | | function | n/a | the function to call when the view details are loaded | see [onViewLoadComplete](#onviewloadcomplete) below
`onListItemLoadComplete` | | function | n/a | the function to call when the list items are loaded | see [onListItemLoadComplete](#onlistitemloadcomplete) below
`onError` | | function | n/a | the function to call if there is an error at any stage | see [onError](#onerror) below
`interval` | | number | n/a | 

### onListLoadComplete

The first thing `nSPGetListItems` does is pull relevant data about the list based on the options you provided. If you provide a `onListLoadComplete` function it will be called with an object with the following properties:

    nSPGetListItems({
        "onListLoadComplete" : function(listData)
    });

listData:

 - `listName`
 - `listGUID`
 - `viewName`
 - `viewGUID`
 - `listPermissions` - boolean values to indicate if the current user has `add`, `edit`, `delete`, and `view` permissions on the list
 - `listForms` - URLs to the `new`, `edit`, and `disp` forms for the list
 - `contentTypes` - an array of all the content types available in the list with the following properties:
   - `name`
   - `description`
   - `forms` - URLs to the `new`, `edit`, and `disp` forms for the content type (may be empty strings)

### onViewLoadComplete

Next `nSPGetListItems` will get details about the view you requested or the default view for the list. If you provide a `onViewLoadComplete` function it will be called with three parameters and expects an `SP.CamlQuery` return.

    nSPGetListItems({
        "onViewLoadComplete" : function(viewFields, viewXML, camlQueru)
        {
		   return camlQuery;
        }
    });


 - `viewFields` - an array of objects providing details about each column/field, in order, from the view:
   - `typeID` - SharePoint field type number (see https://msdn.microsoft.com/en-us/library/office/jj245640.aspx)
   - `typeName` - SharePoint field type name (see https://msdn.microsoft.com/en-us/library/office/jj245640.aspx)
   - `internalName` - the internal name of the field
   - `displayName` - the display name of the field
   - `staticName` - the static name of the field
   - `sourceList` - if `typeName` is `Lookup` then this will have details about the lookup list and field
     - `listPermissions` - same as [onListLoadComplete](#onlistloadcomplete)
     - `listForms` - same as [onListLoadComplete](#onlistloadcomplete)
     - `sourceField` - the details of the lookup field
       - `typeID` - same as above
       - `typeName` - same as above
       - `internalName` - same as above
       - `displayName` - same as above
       - `staticName` - same as above
 - `viewXML` - the XML of the view
 - `camlQuery` - an `SP.CamlQuery` object if you want to do something custom

### onListItemLoadComplete

a

### onError

a

## Examples

a
