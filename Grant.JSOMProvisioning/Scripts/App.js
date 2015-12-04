'use strict';

var context = SP.ClientContext.get_current();
var user = context.get_web().get_currentUser();

var hostWebUrl = decodeURIComponent(getQueryStringParameter("SPHostUrl"));
var appWebUrl = decodeURIComponent(getQueryStringParameter("SPAppWebUrl"));

// This code runs when the DOM is ready and creates a context object which is needed to use the SharePoint object model
$(document).ready(function () {
    getUserName();

    var provisionManager = new Grant.JSOM.Provision.Manager(hostWebUrl, appWebUrl);
    var siteColumns = new Grant.JSOM.Store.SiteColumns();

    $('#btnProvision').click(function () {
        provisionManager.createSiteColumnText(siteColumns.SimpleTextColumn.Name,
            siteColumns.SimpleTextColumn.DisplayName, siteColumns.SimpleTextColumn.Description, siteColumns.SimpleTextColumn.Required, siteColumns.groupName);
    });

    $('#btnUnprovision').click(function () {
        provisionManager.deleteSiteColumn(siteColumns.SimpleTextColumn.DisplayName).then(function () {
            console.info("site column deleted");
        });
    });
});

// This function prepares, loads, and then executes a SharePoint query to get the current users information
function getUserName() {
    context.load(user);
    context.executeQueryAsync(onGetUserNameSuccess, onGetUserNameFail);
}

// This function is executed if the above call is successful
// It replaces the contents of the 'message' element with the user name
function onGetUserNameSuccess() {
    $('#message').text('Hello ' + user.get_title());
}

// This function is executed if the above call fails
function onGetUserNameFail(sender, args) {
    alert('Failed to get user name. Error:' + args.get_message());
}

// Function to retrieve a query string value.  
function getQueryStringParameter(paramToRetrieve) {
    var params = document.URL.split("?")[1].split("&");
    var strParams = "";

    for (var i = 0; i < params.length; i = i + 1) {
        var singleParam = params[i].split("=");
        if (singleParam[0] == paramToRetrieve)
            return singleParam[1];
    }
}
