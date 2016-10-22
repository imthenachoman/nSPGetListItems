# nSPGetListItems

a JavaScript convenience library to get items from SharePoint lists using CSOM/JSOM

https://github.com/imthenachoman/nSPGetListItems

## Table of Contents

 1. [Overview](#overview)
 2. [Reference](#reference)
 3. [Features](#features)
 4. [Examples](#examples)
 
## Overview

When writing code for JavaScript the majority of the time I am getting items from a list. I got tired of having to write the same code over and over again so I wrote the helper library/utility `nSPGetListItems`.

The idea is simple: call `nSPGetListItems` with your [options](#reference)  a list name/GUID at minimum), and get back details about the list, the view, and, of course, the list items.

Additionally, `nSPGetListItems` lets you refresh the data asynchronously so you don't have to code for that. 

## Reference

The main function is called `nSPGetListItems` and it takes one `key : value` object. The available `key`s are:

Key | Required | Type | Default Value | Description | Example
--- | --- | --- | --- | --- | ---
`listName` or `listGUID` | **yes** | string | | the name or GUID of the list | `"Announcements"` or `"{1c7c0498-6f1c-4ec1-8ee6-dd9959f3c52d}"`
`viewName` or `viewGUID` | | string | the list's default view | the name or GUID of the view you want to use to determine the fields and items to pull
`webURL` | | string | the current web | the 
`onListLoadComplete` | | function | n/a | 
`onViewLoadComplete` | | function | n/a | 
`onListItemLoadComplete` | | function | n/a | 
`onError` | | function | n/a | 
`interval` | | number | n/a | 

