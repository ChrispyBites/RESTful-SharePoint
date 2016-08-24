# RESTful SharePoint
A library of jquery and javascript functions used to make life with the REST APIs a little easier.

This library attempts to abstract usage of SharePoint's REST API and will assist you in interacting with SharePoint objects outside the context of SP's UI (i.e., forms, etc.).  This library requires the jQuery library (tested with at least 1.11.2) and the underscore library (for the time being).	

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
* PROF: Work with the user profile service.	
* WEB: Work with websites.		
* SOCIAL: Work with the social service.
* SEARCH: Search REST API calls.

###LIST

####getLists
Get a list of all the lists at the target web site.  By default, this is going to give you just the title and guid for the lists.  Use getListData to get more information about a specific list.  This operation takes no parameters.

####getListData
Query a list for data. This function requires an options object formatted thusly:
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