/********************************************************************************
*																				*
*	API MASTER LIBRARY															*
*	(c) Christopher Parker, TSYS, 2016											*	
*																				*
*	This library attempts to abstract usage of SharePoint's REST API and will	*
*	assist you in interacting with SharePoint objects outside the context of 	*
*	SP's UI (i.e., forms, etc.).  This library requires the jQuery library		*
*	(tested with at least 1.11.2) and the underscore library (for the time 		*
*	being).																		*
*																				*	
*	There are three main sections:												*
*	- #MAIN - All of the primary functions that make everything else go.		*
*	- #OPS	- The actual REST operations themselves.							*												*
*	- #UTIL - Utility functions that format responses or do other calcs.		*
*																				*
*********************************************************************************/

/*#MAIN

	The functions in this section set up functionality for the rest of the 
	library.  Currently we have:
	
	* The primary ApiHelper definition: You'll start with this on the page by 
	creating an api object using the web address that has the resources you want 
	to mess around with:
	
		var MySite = new ApiHelper('[siteURL]');
	
	From there, you can chain any of the functions below to the api object, like so:
		MySite.getListData('MyList','field1,field2,field3');
		MySite.getUsersFromGroup('MyGroup');
	
	* The primary GET function: sends a GET call through AJAX.  If you write new GET functions,
	you MUST reference this and provide an appropriately structured endpoint through relativeUrl.
	
	* The primary POST function: sends a POST call through AJAX.  If you write new POST functions,
	you MUST reference this and provide an appropriately structured endpoint and an HTTP request body object.
	Note that the originating POST operation must be placed on a SharePoint page with jQuery, otherwise the 
	digest grab won't work and the POST function will fail.
*/

window.ApiHelper = function(webUrl){
	this.webUrl = webUrl;
}

ApiHelper.prototype.execute = function(relativeUrl){
	var fullUrl = this.webUrl + "/_api" + relativeUrl;
	var executeOptions = {
		url: fullUrl,
		method: 'GET',
		dataType: "json",
		headers: {
			'accept':'application/json;odata=verbose',
			'content-type':'application/json;odata=verbose'
		}
	};	
	return $.ajax(executeOptions);
}

ApiHelper.prototype.post = function(relativeUrl,body,method){
	var fullUrl = this.webUrl + "/_api" + relativeUrl;
	var executeOptions = {
		url:fullUrl,
		method:"POST",
		data:JSON.stringify(body),
		dataType:'json',
		headers:{
			'X-RequestDigest':$("#__REQUESTDIGEST").val(),
			'accept':'application/json;odata=verbose',
			'content-type':'application/json;odata=verbose',
			"X-HTTP-Method": method,
            "If-Match": "*"
		}
	}
	return $.ajax(executeOptions);
}


ApiHelper.prototype.del = function(relativeUrl){
	var fullUrl = this.webUrl + "/_api" + relativeUrl;
	var executeOptions = {
		url:fullUrl,
		method:"POST",
		headers:{
			'X-RequestDigest':$("#__REQUESTDIGEST").val(),
			"X-HTTP-Method": "DELETE",
            "If-Match": "*"
		}
	}
	return $.ajax(executeOptions);
}
/********************************************************************************************************************/
/*	#OPS																											*/
/*		These are the actual operations.  This section is split into subsections for each API:						*/
/*			- #LIST: Work with lists and list items.																*/
/*			- #FILE: Work with files, folders, and libraries.														*/
/*			- #FIELD: Work with fields.																				*/
/*			- #USER: Work with site users, groups, and roles														*/
/*			- #PROF: Work with the user profile service.															*/
/*			- #WEB: Work with websites.																				*/
/*			- #SOCIAL: Work with the social service.																*/
/*			- #SEARCH: Search REST API calls.																		*/
/*																													*/
/********************************************************************************************************************/

/*#LIST*/

//	Get a list of all the lists at the target web site.
//		By default, this is going to give you just the title and guid for the lists.  Use getListData to get more
//		information about a specific list.  
//
//		This operation takes no parameters.
/*
ApiHelper.prototype.getLists = function(){
	var executeUrl = '/web/lists';
	
}

/*	

Query a list for data. This function requires an options object formatted thusly:
	- list: A string value for the actual list 
	- op: The operations you're performing.  Right now we can do eight operations:
		* 'All': Returns all information about the list itself.
		* 'Fields': List fields.
		* 'Items': List items.
		* 'ContentTypes': Content types configured for the list.
		* 'Forms': The forms the list uses.
		* 'Views': Views configured for the list.
		* 'WorkflowAssociations': Any workflows directly associated with the list.
		* 'Id': Returns the Guid. (Not filterable or selectable).
	- ItemId (optional): The ID of the specific item you want returned.
	- params (optional): An object containing the parameters 
		* select (optional): An array of fields you'd like the query to return.
		* filter (optional): An OData compliant filter expression to apply to the query.
		* expand (optional): Used if any fields you are returning are lookups, people pickers, or managed metadata.  
								Note: For every field you expand, you must also select that field and the appropriate child field, like this:
								'Category/Title'.  If you filter on this field, you will need to format the filter expression similarly.
	- top (optional): The total number of items to return.  If you pass nothing, the value is
	arbitrarily set to 10,000 to return everything.

If you need to pass additional parameters in, you will need to directly edit this function to include them.
The only operations that accept parameters are Fields and Items. If that needs to change down the road, I'll fix it.

*/

ApiHelper.prototype.getListData = function(options){
	//Set the base endpoint Url.
	var executeUrl = "/web/lists/getByTitle('"+options.list+"')";

	//The first two ops are super basic, so we'll collapse them into a single line.
	if(options.op == 'All'){return this.execute(executeUrl).then(function(data){if(data.d){return data.d;}else{throw "Something bad happened..."}});
	}else if(options.op == 'Id'){executeUrl+='/Id';return this.execute(executeUrl).then(function(data){if(data.d){return data.d.Id;}else{throw "Something bad happened..."}});
	}else if(options.op == 'Forms' || options.op == 'Views' || options.op == 'WorkflowAssociations'){
		executeUrl+=options.op;
		return this.execute(executeUrl).then(function(data){
			if(data.d && data.d.results){
				return new QueryResults(data.d.results);
			} else {
				throw "Something bad happened...";
			}
		});
	}else{
		if(options.ItemId){executeUrl+='/items('+options.ItemId+')';return this.execute(executeUrl).then(function(data){if(data.d){return data.d}})}
		else{executeUrl+='/items';}
		//If there's a parameter object, then add the dollar signs, turn it into a string, and add it to the URI.
		if(options.params){
			params = options.params;
			var key;
			for(key in params){
				//This if/else is necessary if you're batching a bunch of list calls at once. 
				//Without it, the parameters accrue dollar signs like an MCU film.
				//Let's hope that's still a relevant pop culture reference...			
				if(params.hasOwnProperty(key) && key.indexOf('$') === -1){params['$'+key] = params[key];delete params[key];}
				else{params[key] = params[key]}
			}

			queryString = decodeURIComponent($.param(params));
			executeUrl+='?'+queryString;
		}
		return this.execute(executeUrl).then(function(data){
			if(data.d && data.d.results){
				return new QueryResults(data.d.results);
			} else {
				throw "Something bad happened...";
			}
		});
	}
}


//Data shape is a little different for a single list item, which necessitates a slightly different return.
ApiHelper.prototype.getListItem = function(list,id,params){
	//Set the base endpoint Url.
	var executeUrl = "/web/lists/getByTitle('"+list+"')/Items("+id+")";

	//If there's a parameter object, then add the dollar signs, turn it into a string, and add it to the URI.
	if(params){
		var key;
		for(key in params){
			//This if/else is necessary if you're batching a bunch of list calls at once. 
			//Without it, the parameters accrue dollar signs like a Kanye album... we all know this one's the greenest meme there is.
			if(params.hasOwnProperty(key) && key.indexOf('$') === -1){params['$'+key] = params[key];delete params[key];}
			else{params[key] = params[key]}
		}
		queryString = decodeURIComponent($.param(params));
		executeUrl+='?'+queryString;
	}
	return this.execute(executeUrl).then(function(data){
		if(data.d){
			return data.d;
		} else {
			throw "Something bad happened...";
		}
	});
}

// Adds a list item to a list.  Parameters include a list name and a fields regular object.
// The fields object needs to include a property for at *least* each required field in the list.  
// Note that lookups, MM, and PP fields require special handling that I'm not super clear on, yet.
ApiHelper.prototype.addListItem = function(list,fields){
	var url = "/web/lists/getbytitle('"+list+"')/items";
	var type = getType(list);				//Builds the content type internal name and passes it to the HTTP request body.
	metadata = {							//Start of the HTTP request body object.
		'__metadata':{
			'type':type
		}
	};
	var body = _.extend(metadata,fields);	//Least complicated way to merge the metadata and fields objects.
	var method = 'POST';
	return this.post(url,body,method).then(function(response){
		if(response.d){
			return response.d;					//Do something with this on the page; I usually display some sort of "Thanks for your submission" text or alert or something.
		}else{
			throw false;
		}
	});
}

//Updates an existing item.  The Fields object must contain the ID of the item you're updating.  Otherwise,
//this operation is virtually the same as the addListItem operation.

ApiHelper.prototype.updateListItem = function(list,fields){
	var url = "/web/lists/getbytitle('"+list+"')/items("+fields.ID+")";
	var type = getType(list);				//Builds the content type internal name and passes it to the HTTP request body.
	metadata = {							//Start of the HTTP request body object.
		'__metadata':{
			'type':type
		}
	};
	var body = _.extend(metadata,fields);	//Least complicated way to merge the metadata and fields objects.
	var method = 'MERGE';
	return this.post(url,body,method);
}

ApiHelper.prototype.deleteListItem = function(list,id){
	var url = "/web/lists/getbytitle('"+list+"')/items("+id+")";
	return this.del(url);
}

/*#FILE*/

//Get a list of files, given a document library and a folder.  The options object needs to be formatted as follows:
// - options.library = The Title of the library (so, 'Shared Documents' not 'Shared%20Documents').
// - options.folders = An array containing the folder path you want to enumerate files for.  The function will expect
//   the array to be in the hierarchical order of the folders on the site.  i.e., ['Folder A','Folder B'] will be assumed to be Folder A/Folder B.
// - options.files = A boolean indicating whether the function should get files (true) or folders (false)

// Note, that this function is not intended to return contents from multiple folder locations.  You'll need to do that via promises and
// multiple calls.

ApiHelper.prototype.getFilesFromFolder = function(options){
	var files = [];
	var endpoint = '/web/getfolderbyserverrelativeurl';
	var path = '';
	var i;
	path += options.library + '/';
	for(i=0;i<options.folders.length;i++){
		if(i === options.folders.length - 1){
			path += options.folders[i];
		} else {
			path += options.folders[i] + '/';
		}
	}
	if(options.files){
		var executeUrl = endpoint + "('"+path+"')/files";
		return this.execute(executeUrl).then(function(data){
			if(data.d && data.d.results){

				for(i=0;i<data.d.results.length;i++){
					files.push({
						'Url':data.d.results[i].ServerRelativeUrl,
						'Title':data.d.results[i].Title,
						'Name':data.d.results[i].Name,
						'Version':data.d.results[i].UIVersionLabel,
						'Author':data.d.results[i].Author
					});
				}
				return files;
			} else {throw "Something bad happened...";}
		});
	} else {
		var executeUrl = endpoint + "('"+path+"')/folders";
		return this.execute(executeUrl).then(function(data){
			if(data.d && data.d.results){
				for(i=0;i<data.d.results.length;i++){
					files.push({
						'Name':data.d.results[i].Name,
						'Url':data.d.results[i].ServerRelativeUrl,
						'Items':data.d.results[i].ItemCount
					});
				}
				return files;
			}
		});
	}
}



/*#FIELD*/

//Used to return the available choices of a choice field.  You'll need to hand it the list name and guid of the field.
ApiHelper.prototype.getFieldChoices = function(list,fguid){
	var executeUrl = "/web/lists/getByTitle('"+list+"')/fields(guid'"+fguid+"')/Choices";
	return this.execute(executeUrl).then(function(data){
		if(data.d && data.d.Choices){
			return data.d.Choices.results;
		}else{
			throw "Something bad happened..."
		}
	});

}

ApiHelper.prototype.addNewChoices=function(list,fguid,choices){
	var executeUrl="/web/lists/getByTitle('"+list+"')/fields(guid'"+fguid+"')"
	var body = {'__metadata':{'type':'SP.FieldChoice'},'Choices':{ '__metadata':{ 'type': 'Collection(Edm.String)' }, 'results': choices }};
	var method = 'MERGE';
	return this.post(executeUrl,body,method);
}
/*#USER*/

//Query a group for user names.  This function only needs a group name.
ApiHelper.prototype.getUsersFromGroup = function(group){
	var executeUrl = "/web/sitegroups/getbyname('"+group+"')/Users";
	return this.execute(executeUrl).then(function(data){
		if(data.d && data.d.results){
			return new QueryResults(data.d.results);
		} else {
			throw "Something bad happened...";
		}
	});
}

//Can't believe I found this.  There's an API endpoint to perform the wonky ensureUser function that allows you to put people in people pickers!  Pass it the full name, including claims.
//You'll get an ID back that you need to use to post to a people picker field for writing people.
ApiHelper.prototype.ensureUser = function(username){
	var executeUrl="/web/ensureUser";
	var method='POST';
	body={
		"logonName":username
	};
	return this.post(executeUrl,body,method).then(function(data){
		if(data.d){
			return data.d.Id;
		} else {
			throw "Something bad happened...";
		}
	});	
}

/*#PROF*/

//Get the current user's organizational structure.  This takes no parameters, and returns an array of values for
//the user's manager, peers, reports, and their own display name.

ApiHelper.prototype.getMyOrgStructure = function(){
	var managers = [];
	var peers = [];
	var reports = [];
	var me = {};
	var executeUrl = '/sp.userprofiles.peoplemanager/getmyproperties';
	return this.execute(executeUrl).then(function(response){
		if(response.d){
			me = {
				'Account':response.d.AccountName,
				'Name':response.d.DisplayName,
				'Title':response.d.Title,
				'Picture':response.d.PictureUrl,
				'Mysite':response.d.PersonalUrl,
				'Email':response.d.Email
			}
			manResults = response.d.ExtendedManagers.results;
			for(var i=0;i < manResults.length;i++){
				managers.push(manResults[i]);
			}
			peerResults = response.d.Peers.results;
			for(var i=0;i<peerResults.length;i++){
				peers.push(peerResults[i]);
			}
			reportResults = response.d.DirectReports.results;
			for(var i=0;i<reportResults.length;i++){
				reports.push(reportResults[i]);
			}
			
			return {
				Me:me,
				Managers:managers,
				Peers:peers,
				Reports:reports
			}
		} else{
			throw "Something bad happened...";
		}
	});
}


//Just gets the organizational structure.  Use getUserProperties for a more expansive user profile property return.
//This and the following operation take the user's LAN ID as a parameter.
ApiHelper.prototype.getOrgStructure = function(user){
	var managers = [];
	var peers = [];
	var reports = [];
	var user = {};
	var executeUrl = "/sp.userprofiles.peoplemanager/getpropertiesfor(@v)?@v='"+user+"'";
	return this.execute(executeUrl).then(function(response){
		if(response.d){
			user = {
				'Account':response.d.AccountName,
				'Name':response.d.DisplayName,
				'Title':response.d.Title,
				'Picture':response.d.PictureUrl,
				'Mysite':response.d.PersonalUrl,
				'Email':response.d.Email
			}
			manResults = response.d.ExtendedManagers.results;
			for(var i=0;i < manResults.length;i++){
				managers.push(manResults[i]);
			}
			peerResults = response.d.Peers.results;
			for(var i=0;i<peerResults.length;i++){
				peers.push(peerResults[i]);
			}
			reportResults = response.d.DirectReports.results;
			for(var i=0;i<reportResults.length;i++){
				reports.push(reportResults[i]);
			}
			
			return {
				User:user,
				Managers:managers,
				Peers:peers,
				Reports:reports
			}
		} else{
			throw "Something bad happened...";
		}
	});
}

ApiHelper.prototype.getUserProperties = function(user){
	var executeUrl = "/sp.userprofiles.peoplemanager/getpropertiesfor(@v)?@v='"+user+"'";
	return this.execute(executeUrl).then(function(response){
		if(response.d && response.d.UserProfileProperties){
			var props = {};
			uProps = response.d.UserProfileProperties.results;
			for(var i=0;i<uProps.length;i++){
				props[uProps[i].Key] = uProps[i].Value;
			}
			
			var user = {
				'Name':response.d.DisplayName,
				'Title':response.d.Title,
				'Email':response.d.Email,
				'Props':props
			}
			return user;
		} else {
			throw "Something bad happened...";
		}
	});
}
/*#WEB*/

//Get the list of subsites for the site collection.  This function needs nothing at all.
//Returns an array sites where sites[0] is the title and sites[1] is the URL.
ApiHelper.prototype.getSubsites = function(){
	var sites = [];
	var executeUrl = '/web/webs';
	return this.execute(executeUrl).then(function(data){
		if(data.d && data.d.results){
			for(var i=0;i<data.d.results.length;i++){
				sites.push({
					'Title':data.d.results[i].Title,
					'Url':data.d.results[i].Url
				});
			}
			return sites;
		}
	});
}

/*#SOCIAL*/

//Access the social feed for a count of unread mentions.  Curious to see if it also gets unread... everythings.
ApiHelper.prototype.getUnreadCount = function(){
	var executeUrl = '/social.following/my/followed(types=8)';
	return this.execute(executeUrl).then(function(data){
		if(data.d){
			return data.d;
		} else {
			throw 'Something happened omg';
		}
	});
}

/*#SEARCH*/


//Uses the search API to specifically return users, specifically.  This function takes a javascript object
//as its parameter:
//	-- options.query: The actual query text (i.e., someone's last name).  If you're searching for everyone,
//		send "*" in this property.  Don't do this.  Please.
//	-- options.props: An array of values corresponding to the user profile properties you want back.
//		keep in mind that these have to be formated as they are in the data, so its highly recommended that you
//		run a few test queries in an application like PostMan before using this.  If you leave it off, the search will grab
//		everything.

ApiHelper.prototype.searchUsers = function(options){
	if(options.props){
		var executeUrl = "/search/query?querytext='"+options.query+"'&sourceid='b09a7990-05ea-4af9-81ef-edfab16c4e31'&selectproperties='"+options.props.toString()+"'&rowlimit=500";
	} else {
		var executeUrl = "/search/query?querytext='"+options.query+"'&sourceid='b09a7990-05ea-4af9-81ef-edfab16c4e31'&rowlimit=500";
	}
	return this.execute(executeUrl).then(function(data){
		if(data.d && data.d.query){
			return new SearchResults(data.d.query);
		} else {
			throw "Something bad happened...";
		}
	});
}

//Users the search API to return results from a specific result source ID.  This function takes an object for a paramter:
//	-- options.source: The source ID of the result source.  If you're not sure what this is... you probably shouldn't be using this.
//	-- options.rowlimit: A numerical value for the total number of rows you want back.  If you're not sure, set it to something high, like 500.
//  -- options.params: An optional object with properties for each parameter you want to pass in.  Search API is... kinda funky, so read up on what's available.  For now:
//		selectproperties: An array of values for each managed property you want to select.
//		querytext: A string of text to search against.  Not always reliable, especially if the result source is not properly configured.

ApiHelper.prototype.searchResultSource = function(options){
	if(options.params){
		var querytext=(options.params.querytext)?options.params.querytext:"*"
		var executeUrl = "/search/query?querytext='"+querytext+"'&sourceid='"+options.source+"'&selectproperties='"+options.params.selectproperties.toString()+"'&rowlimit="+options.rowlimit
	} else {
		var executeUrl = "/search/query?sourceid='"+options.source+"'&rowlimit="+options.rowlimit;
	}
	return this.execute(executeUrl).then(function(data){
		if(data.d && data.d.query){
			return new SearchResults(data.d.query);
		}else{
			throw "Something bad happened...";
		}
	});
}

ApiHelper.prototype.searchQuery = function(options){
	var i,querytext='';
	for(i=0;i<options.querytext.length;i++){
		querytext+=options.querytext[i]+' ';
	}
	if(options.params){
		var executeUrl = "/search/query?querytext='"+querytext+"'&selectproperties='"+options.params.selectproperties.toString()+"'&rowlimit="+options.rowlimit;
	}else{
		var executeUrl = "/search/query?querytext='"+querytext+"'&rowlimit="+options.rowlimit;
	}
	return this.execute(executeUrl).then(function(data){
		if(data.d && data.d.query){
			return new SearchResults(data.d.query);
		}else{
			throw "Something bad happened...";
		}
	});
}

/************************************************************************************************************/
/*																											*/
/*	#UTIL																									*/
/*		These functions are used by other functions to manipulate items, build objects, create				*/ 
/*		special variables or or format returns.  Additional utilities should be built here.					*/
/*		Where possible, we'd like to minify these things.													*/
/*																											*/
/************************************************************************************************************/

// QueryResults takes the returns from one of the GET functions and builds an array of objects, where each object is a single return.
function QueryResults(response){var items = [];for(i=0;i<response.length;i++){items.push(response[i]);}return items;}

//Search results takes the return from the search API GET and pulls out the total results along with the actual results and prepares them for formatting
function SearchResults(data){
	this.resultsCount = data.PrimaryQueryResult.RelevantResults.TotalRows;
	this.items = convertRowsToObjects(data.PrimaryQueryResult.RelevantResults.Table.Rows.results,this.resultsCount);
}

//As the name implies, takes the rows from the query return and turns them into javascript objects.
function convertRowsToObjects(itemRows){
	var items = [];
	for(var i=0;i<itemRows.length;i++){
		var row = itemRows[i], item = {};
		for(var j=0;j<row.Cells.results.length;j++){
			item[row.Cells.results[j].Key] = row.Cells.results[j].Value;
		}
		items.push(item);
	}
	return items;
}

//Builds the internal content type name.  Apparently they're all built the same.  I have a feeling that lists with
//site or list content types will need some special handling.
function getType(name){
	if(name=="Client Contact Database"){
		return "SP.Data.ClientContactListListItem";
	}else{
		return "SP.Data." + name.charAt(0).toUpperCase() + name.split(" ").join("").slice(1) + "ListItem";
	}
}

function ExceptionMsg(err){
	console.log(err); 
}

//Probably need to find a new home for this, but the following will wait until a specified div exists before running some specified function.
;(function ($, window) {

var intervals = {};
var removeListener = function(selector) {

    if (intervals[selector]) {

        window.clearInterval(intervals[selector]);
        intervals[selector] = null;
    }
};
var found = 'waitUntilExists.found';

/**
 * @function
 * @property {object} jQuery plugin which runs handler function once specified
 *           element is inserted into the DOM
 * @param {function|string} handler 
 *            A function to execute at the time when the element is inserted or 
 *            string "remove" to remove the listener from the given selector
 * @param {bool} shouldRunHandlerOnce 
 *            Optional: if true, handler is unbound after its first invocation
 * @example jQuery(selector).waitUntilExists(function);
 */

$.fn.waitUntilExists = function(handler, shouldRunHandlerOnce, isChild) {

    var selector = this.selector;
    var $this = $(selector);
    var $elements = $this.not(function() { return $(this).data(found); });

    if (handler === 'remove') {

        // Hijack and remove interval immediately if the code requests
        removeListener(selector);
    }
    else {

        // Run the handler on all found elements and mark as found
        $elements.each(handler).data(found, true);

        if (shouldRunHandlerOnce && $this.length) {

            // Element was found, implying the handler already ran for all 
            // matched elements
            removeListener(selector);
        }
        else if (!isChild) {

            // If this is a recurring search or if the target has not yet been 
            // found, create an interval to continue searching for the target
            intervals[selector] = window.setInterval(function () {

                $this.waitUntilExists(handler, shouldRunHandlerOnce, true);
            }, 500);
        }
    }

    return $this;
};

}(jQuery, window));
