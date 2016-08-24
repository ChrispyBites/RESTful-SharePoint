window.restObject = function(webUrl){
	this.webUrl = webUrl;
}

restObject.prototype.execute = function(relativeUrl){
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

restObject.prototype.post = function(relativeUrl,body,method){
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


restObject.prototype.del = function(relativeUrl){
    var fullUrl=this.webUrl +"/_api"+ relativeUrl;
	var executeOptions={url:fullUrl,method:"POST",headers:{'X-RequestDigest':$("#__REQUESTDIGEST").val(),"X-HTTP-Method": "DELETE","If-Match": "*"}}
	return $.ajax(executeOptions);
}

//A list of lists on the site.
restObject.prototype.getLists = function(){
	var executeUrl = '/web/lists';
	return this.execute(executeUrl).then(function(data){
		if(data.d){
			return data.d
		}else{
			//Someone help me figure out how to carry these errors out of here and back into whatever context you're calling this from.
			throw "Something bad happened..."
		}
	});
}

//Data about or within a specific list.
restObject.prototype.getListData = function(options){
	var executeUrl = "/web/lists/getByTitle('"+options.list+"')";
	if(options.op == 'All'){
		return this.execute(executeUrl).then(function(data){
			if(data.d){
				return data.d;
			}else{
				throw "Something bad happened..."
			}
		});
	}else if(options.op == 'Id'){
		executeUrl+='/Id';return this.execute(executeUrl).then(function(data){
			if(data.d){
				return data.d.Id;
			}else{
				throw "Something bad happened..."
			}
		});
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
		if(options.ItemId){
			executeUrl+='/items('+options.ItemId+')';
			return this.execute(executeUrl).then(function(data){
				if(data.d){
					return data.d;
				}
			});
		}else{
			executeUrl+='/items';
		}
		if(options.params){
			params=options.params;
			var key;
			for(key in params){
				//This if/else is necessary if you're batching a bunch of list calls at once. Without it, the parameters accrue dollar signs like an MCU film. Let's hope that's still a relevant pop culture reference by the time anyone sees this.			
				if(params.hasOwnProperty(key) && key.indexOf('$') === -1){
					params['$'+key] = params[key];
					delete params[key];
				}else{
					params[key] = params[key]
				}
			}
			queryString=decodeURIComponent($.param(params));
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

//Data on a specific list item.  Return shape is a little different; this necessitates a slightly different return.
restObject.prototype.getListItem=function(list,id,params){
	var executeUrl = "/web/lists/getByTitle('"+list+"')/Items("+id+")";
	if(params){
		var key;
		for(key in params){
			//This if/else is necessary if you're batching a bunch of list calls at once. Without it, the parameters accrue dollar signs like a Kanye album... we all know this one's the greenest meme there is.  I should probably turn this into a utility function.
			if(params.hasOwnProperty(key) && key.indexOf('$') === -1){
				params['$'+key] = params[key];
				delete params[key];
			}else{
				params[key] = params[key]
			}
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

// Adds an item to a list.
restObject.prototype.addListItem = function(list,fields){
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
			return response.d;				//Do something with this on the page; I usually display some sort of "Thanks for your submission" text or alert or something.
		}else{
			throw false;
		}
	});
}

//Updates an existing item.
restObject.prototype.updateListItem = function(list,fields){
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

restObject.prototype.deleteListItem = function(list,id){
	var url = "/web/lists/getbytitle('"+list+"')/items("+id+")";
	return this.del(url);
}

/*#FILE*/

//Get a list of files from a document library/folder.
restObject.prototype.getFilesFromFolder = function(options){
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
restObject.prototype.getFieldChoices = function(list,fguid){
	var executeUrl = "/web/lists/getByTitle('"+list+"')/fields(guid'"+fguid+"')/Choices";
	return this.execute(executeUrl).then(function(data){
		if(data.d && data.d.Choices){
			return data.d.Choices.results;
		}else{
			throw "Something bad happened..."
		}
	});

}

restObject.prototype.addNewChoices=function(list,fguid,choices){
	var executeUrl="/web/lists/getByTitle('"+list+"')/fields(guid'"+fguid+"')"
	var body = {'__metadata':{'type':'SP.FieldChoice'},'Choices':{ '__metadata':{ 'type': 'Collection(Edm.String)' }, 'results': choices }};
	var method = 'MERGE';
	return this.post(executeUrl,body,method);
}
/*#USER*/

//Query a group for user names.  This function only needs a group name.
restObject.prototype.getUsersFromGroup = function(group){
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
restObject.prototype.ensureUser = function(username){
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

restObject.prototype.getMyOrgStructure = function(){
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
restObject.prototype.getOrgStructure = function(user){
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

restObject.prototype.getUserProperties = function(user){
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
restObject.prototype.getSubsites = function(){
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
restObject.prototype.getUnreadCount = function(){
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

restObject.prototype.searchUsers = function(options){
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

restObject.prototype.searchResultSource = function(options){
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

restObject.prototype.searchQuery = function(options){
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
function SearchResults(data){this.resultsCount=data.PrimaryQueryResult.RelevantResults.TotalRows;this.items=convertRowsToObjects(data.PrimaryQueryResult.RelevantResults.Table.Rows.results,this.resultsCount)}

//As the name implies, takes the rows from the query return and turns them into javascript objects.
function convertRowsToObjects(itemRows){var items = [];itemRows.map(function(o){var row=o,item={};o.Cells.results.map(function(a){item[a.Key]=item[a.Value]});items.push(item)});return items;}

//Builds the internal content type name.  Apparently they're all built the same.  I have a feeling that lists with site or list content types will need some special handling.
function getType(name){return "SP.Data." + name.charAt(0).toUpperCase() + name.split(" ").join("").slice(1) + "ListItem";}

function ExceptionMsg(err){
	console.log(err); 
}