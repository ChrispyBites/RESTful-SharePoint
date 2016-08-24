# RESTful SharePoint
A library of jquery and javascript functions used to make life with the REST APIs a little easier.

This library attempts to abstract usage of SharePoint's REST API and will assist you in interacting with SharePoint objects outside the context of SP's UI (i.e., forms, etc.).  This library requires the jQuery library (tested with at least 1.11.2) and the underscore library (for the time being).	

There are three main sections:
    MAIN - All of the primary functions that make everything else go.
    OPS	- The actual REST operations themselves.
    UTIL - Utility functions that format responses or do other calcs.

#MAIN
The functions in this section set up functionality for the rest of the library.  Currently we have:
	
	* The primary restObject definition: You'll start with this on the page by 
	creating an api object using the web address that has the resources you want 
	to mess around with:
	
		var MySite = new restObject('[siteURL]');
	
	From there, you can chain any of the functions below to the api object, like so:
		MySite.getListData('MyList','field1,field2,field3');
		MySite.getUsersFromGroup('MyGroup');
	
	* The primary GET function: sends a GET call through AJAX.  If you write new GET functions,
	you MUST reference this and provide an appropriately structured endpoint through relativeUrl.
	
	* The primary POST function: sends a POST call through AJAX.  If you write new POST functions,
	you MUST reference this and provide an appropriately structured endpoint and an HTTP request body object.
	Note that the originating POST operation must be placed on a SharePoint page with jQuery, otherwise the 
	digest grab won't work and the POST function will fail.