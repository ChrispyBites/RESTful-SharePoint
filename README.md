# RESTful SharePoint
A library of jquery and javascript functions used to make life with the REST APIs a little easier.

This library attempts to abstract usage of SharePoint's REST API and will assist you in interacting with SharePoint objects outside the context of SP's UI (i.e., forms, etc.).  This library requires the jQuery library (tested with at least 1.11.2) and the underscore library (for the time being).

Unless otherwise indicated, all functions return an array of JSON objects that correspond to whatever you asked for. 

There are three main sections:
    MAIN - All of the primary functions that make everything else go.
    OPS	- The actual REST operations themselves.
    UTIL - Utility functions that format responses or do other calcs.

##MAIN
The functions in this section set up functionality for the rest of the library.  Currently we have:
	
The primary restObject definition: You'll start with this on the page by creating an api object using the web address that has the resources you want to mess around with:
	
	var MySite = new restObject('[siteURL]');
	
From there, you can chain any of the functions below to the api object, like so:
	MySite.getListData('MyList','field1,field2,field3');
	MySite.getUsersFromGroup('MyGroup');
	
The primary GET function sends a GET call through AJAX.  If you write new GET functions, you MUST reference this and provide an appropriately structured endpoint through relativeUrl.
	
The primary POST function sends a POST call through AJAX.  If you write new POST functions, you MUST reference this and provide an appropriately structured endpoint and an HTTP request body object. Note that the originating POST operation must be placed on a SharePoint page with jQuery, otherwise the digest grab won't work and the POST function will fail..

There is a secondary POST function used to handle delete requests; headers are formatted just differently enough that I felt like it needed it's own AJAX function.

##OPS											
These are the actual operations.  This section is split into subsections for each API:						
* LIST: Work with lists and list items.
* FILE: Work with files, folders, and libraries.
* FIELD: Work with fields.
* USER: Work with site users, groups, and roles.
* PROFILE: Work with the user profile service.	
* WEB: Work with websites.		
* SOCIAL: Work with the social service.
* SEARCH: Search REST API calls.

###LIST

####getLists
Get a list of all the lists at the target web site.  By default, this is going to give you just the title and guid for the lists.  Use getListData to get more information about a specific list.  This operation takes no parameters.

####getListData
Query a list for data. This function requires an options object formatted thusly:
```
	{
		list : A string value for the list's display name (not it's URL).
		op : The operation you're performing.  I have eight different operations configured for this function, listed below the example object.
		ItemId : Optionally the ID of a single list item you might want returned.  See *getListItem* for a direct way of doing this.
		top : Optionally, the total number of items to return.  If you pass nothing, this is arbitrarily set to 10,000 to (hopefully) return everything.
		params : An object containing, appropriately, parameters for the query.  All are optional, including the whole parameter object.  
			I've accounted for the following parameters: 
			{
				select : An array of fields (internal names) you'd like returned.
				filter : An OData compliant filter expression to apply to the query.
				expand : Used if any fields are lookups, people pickers, or managed metadata.  Any expanded field must also be included in the select array, pathed to the child field you're looking for.  For example: 'Category/Title'.
			} 
	}
```

**Available Operations**

|Operation|Description|
|---------|-----------|
|`All`|Returns everything about the list.|
|`Fields`|Returns a list of fields used on the list.|
|`ContentTypes`|REturns a list of content types configured on the list.|
|`Forms`|Any custom (read: InfoPath) forms used on the list.|
|`Views`|All the views configured on the list.|
|`WorkflowAssociations`|Any workflows associated with the list.  This will not include site/reusable workflows associated with content types that are configured on the list.|
|`Id`|Returns the list's GUID.|

If you need to pass additional parameters in, you will need to directly edit this function to include them.  The only operations that accept parameters are Fields and Items. If that needs to change down the road, I'll fix it.

####getListItem
Gets all data on a single list item.  It takes three parameters:
* `list` : The list's display name.
* `id` : The ID of the list item you want returned.
* `params` : An object, structured the same as the one used in *getListData*, but without the filter parameter.

####addListItem
Adds a new item to a list.  This operation takes a two parameters:
* `list` : The list's display name.
* `fields` : As you might expect, an object containing field/value pairs for the new item.  Obviously, this object should contain *at least* a field/value pair for any required fields on the list.  If you leave off required fields, this operation will likely silently fail.  Lookups, managed metadata, and people picker fields are stupidly complex and will probably require a level of special handling I likely haven't accounted for here.  I mean, you can still update items with those fields in them, you'll just have to structure your fields object according to how Microsoft expects that payload to look.

Note that your return value is the new list item, which you can use for some super nice UX things because you're decent to your users.

####updateListItem
Updates an existing list item.  Same parameters as *addListItem*, with one caveat: the fields object *must* contain the ID of the item you're updating, otherwise you'll just add a new item to the list.  I believe the return value is an HTTP code, like 200 for a successful update.  I don't really remember. Same basic rules apply, though.  Help your users out and let them know something successful happened.

####deleteListItem
This should be self explanatory.  Pass in the list's display name and an ID, both as strings.  If I remember correctly, you don't get anything at all back.  Eff you, Microsoft.

###FILE
The file endpoints are kind of strange and I'm still working out how they work, exactly.  There's a lot of potential, though, for some powerful stuff.

####getFilesFromFolder
Gets a list of files, given a document library and a folder.  The options object needs to be formatted as follows:
    {
		'library' : The display name of the library.
		'folders' : An array containing the folder path you want to grab files from.  The function will expect the array to be in the hierarchical order of the folders on the site.  For example, ['Folder A','Folder B'] will be assumed to be Folder A/Folder B.
		'files' : A boolean value indicating the function should get files (true) or folders (false)
	}

###FIELD
Most of these functions will deal with creating, removing, or changing fields on lists.  I'm unclear how this impacts site columns.

####getFieldChoices
This is used to get the available choices for a regular choice field (i.e., not a people picker, lookup, or managed metadata field).  It takes two parameters: the display name of the list and the GUID of the field in question.  I probably need to change this to work off of the display name of the field, but meh.

####addNewChoices
Adds a new choice to a choice field's list of choices, and it takes three parameters: the list name, the field guid, and an array for the choices.  The title of this function is kind of a misnomer, as what you're really doing is updating the choices list.  This is super important to remember, because if you don't want to screw up the existing list of choices, you'll need to create an array of the original list of choices (see *getFieldChoices*) and append your new choice to the end of that array.  I'm sure some day I (or some enterprising contributor) will leverage *getFieldChoices* to do this for you.  I'm lazy bruh.

###USER
This is a distinct endpoint from the user profile service endpoints.  These deal with the user information list (UIL) and groups and permissions at a site.

####getUsersFromGroup
Grabs all the users in the specified group.  Only takes a single parameter, the group's name.

####ensureUser
Given a `username`, including claims encoding (if that's what your organization gets into), this will add a user to the UIL and return to you that user's ID in the UIL.  Fantastically useful for making custom people pickers work.

###PROFILE 
These functions are used to manipulate or read user profile properties.  Most of the functions will use the same endpoint over and over again, I'm just formatting the output for some specific needs.  YMMV with these, as everyone's profile properties are called different things and mapped differently.  If you bring this to your organization, you'll probably want to see how I formatted the return objects and make sure they're sending you stuff you want.

####getMyOrgStructure
This one takes no parameter, and returns an array containing four objects: me, managers, peers, and reports.  Managers, peers, and reports are just an array of account IDs (LAN IDs in my company's case) that correspond to those groups of people relative to the current user.  Me contains the following information about the current user:
* Account ID
* Display name
* Job Title
* Picture URL
* MySite URL
* Email address

####getOrgStructure
This returns the exact same array of objects as *getMyOrgStructure*, but for the individual (using their account ID) specified as a parameter.

####getUserProperties
Takes a single parameter, the account ID, and returns an object with four properties: the user's display name, job title (I think...), email address, and an array of key/value pairs for every user profile property you might have banging around.

###WEB

####getSubsites

###SOCIAL

####getUnreadCount

###SEARCH

####searchUsers

####serachResultSource

####searchQuery

###UTILS

####QueryResults

####SearchResults

####convertRowsToObjects

####getType

####ExceptionMsg

