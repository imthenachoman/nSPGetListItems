# nSPGetListItems

a JavaScript convenience library to get items from SharePoint lists using CSOM/JSOM

https://github.com/imthenachoman/nSPGetListItems

## Table of Contents

 1. [Overview](#overview)
 2. [Reference](#reference)
 3. [Features](#features)
 4. [Examples](#examples)
 
## Overview

When writing code for JavaScript the majority of the time I am getting items from a list or document library. I got tired of having to write the same code over and over again so I wrote the helper library/utility `nSPGetListItems`.

The idea is simple: call `nSPGetListItems` with your [options](#reference)  (a list or document library name at minimum), and get back details about the list, the view, and, of course, the list items.

Additionally, `nSPGetListItems` lets you refresh the data asynchronously so you don't have to code for that. 

## Reference

The main function is called `nSPGetListItems` and it takes one `key : value` object. The available `key`s are:

Key | Required | Type | Default Value | Description | Example
--- | --- | --- | --- | --- | ---
`listName` or `listGUID` | **yes** | string | | the name or GUID of the list | <ul><li>`"Announcements"`</li><li>`"{1c7c0498-6f1c-4ec1-8ee6-dd9959f3c52d}"`</li></ul>
`viewName` or `viewGUID` | | string | the list's default view | the name or GUID of the view you want to use to determine the fields and items to pull | <ul><li>`"All Documents"`</li><li>`"{dce68293-70e5-4c47-acda-72e6236b8f65}"`</li></ul>
`webURL` | | string | the current web | the web URL you want to pull from | <ul><li>`"/someSite"`</li><li>`"/someSite/subSite"`</li></ul>
`onListLoadComplete` | | function | n/a | the function to call when the list details are loaded | see [onListLoadComplete](#onListLoadComplete) below
`onViewLoadComplete` | | function | n/a | the function to call when the view details are loaded | see [onViewLoadComplete](#onViewLoadComplete) below
`onListItemLoadComplete` | | function | n/a | the function to call when the list items are loaded | see [onListItemLoadComplete](#onListItemLoadComplete) below
`onError` | | function | n/a | the function to call if there is an error at any stage | see [onError](#onError) below
`interval` | | number | n/a | 

### onListLoadComplete

a

### onViewLoadComplete

a

### onListItemLoadComplete

a

### onError

