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
    var siteContentTypes = new Grant.JSOM.Store.SiteContentTypes();

    $('#btnProvision').click(function () {
        provisionManager.createSiteColumn(provisionManager.createSiteColumnTextFieldXml(siteColumns.SimpleTextColumn.Name,
            siteColumns.SimpleTextColumn.DisplayName, siteColumns.SimpleTextColumn.Description,
            siteColumns.SimpleTextColumn.Required, siteColumns.groupName)).then(function () {
                console.log("Created site column: " + siteColumns.SimpleTextColumn.DisplayName);
                provisionManager.createSiteColumn(provisionManager.createSiteColumnNumberFieldXml(siteColumns.NumberColumn.Name,
                    siteColumns.NumberColumn.DisplayName, siteColumns.NumberColumn.Description, siteColumns.NumberColumn.Max,
                    siteColumns.NumberColumn.Min, siteColumns.NumberColumn.Decimals,
                    siteColumns.NumberColumn.Required, siteColumns.groupName)).then(function () {
                        console.log("Created site column: " + siteColumns.NumberColumn.DisplayName);
                        provisionManager.createSiteColumn(provisionManager.createSiteColumnUrlFieldXml(siteColumns.UrlColumn.Name,
                            siteColumns.UrlColumn.DisplayName, siteColumns.UrlColumn.Description,
                            siteColumns.UrlColumn.Required, siteColumns.groupName)).then(function () {
                                console.log("Created site column: " + siteColumns.UrlColumn.DisplayName);
                                provisionManager.createSiteColumn(provisionManager.createSiteColumnImageFieldXml(siteColumns.ImageColumn.Name,
                                    siteColumns.ImageColumn.DisplayName, siteColumns.ImageColumn.Description,
                                    siteColumns.ImageColumn.Required, siteColumns.groupName)).then(function () {
                                        console.log("Created site column: " + siteColumns.ImageColumn.DisplayName);
                                        provisionManager.createSiteColumn(provisionManager.createSiteColumnDropDownFieldXml(siteColumns.DropDownColumn.Name,
                                            siteColumns.DropDownColumn.DisplayName, siteColumns.DropDownColumn.Description, siteColumns.DropDownColumn.Choices,
                                            siteColumns.DropDownColumn.Required, siteColumns.groupName)).then(function () {
                                                console.log("Created site column: " + siteColumns.DropDownColumn.DisplayName);
                                                provisionManager.createSiteContentType(siteContentTypes.contentType.Id,
                                                    siteContentTypes.contentType.Name, siteContentTypes.groupName,
                                                    siteContentTypes.contentType.Description, siteContentTypes.contentType.Columns).then(function () {
                                                        console.log("Created content type: " + siteContentTypes.contentType.Name);
                                                        provisionManager.createCustomList("JSOM Custom List", siteContentTypes.contentType.Id).then(function () {
                                                            console.log("Created custom list");
                                                        });
                                                    });
                                            });
                                    });
                            });
                    });
            });
    });

    $('#btnUnprovision').click(function () {
        provisionManager.deleteCustomList("JSOM Custom List").then(function () {
            console.log("Deleted custom list");
            provisionManager.deleteSiteContentType(siteContentTypes.contentType.Id).then(function () {
                provisionManager.deleteSiteColumn(siteColumns.SimpleTextColumn.DisplayName).then(function () {
                    console.info("site column deleted: " + siteColumns.SimpleTextColumn.DisplayName);
                    provisionManager.deleteSiteColumn(siteColumns.NumberColumn.DisplayName).then(function () {
                        console.info("site column deleted: " + siteColumns.NumberColumn.DisplayName);
                        provisionManager.deleteSiteColumn(siteColumns.UrlColumn.DisplayName).then(function () {
                            console.info("site column deleted: " + siteColumns.UrlColumn.DisplayName);
                            provisionManager.deleteSiteColumn(siteColumns.ImageColumn.DisplayName).then(function () {
                                console.info("site column deleted: " + siteColumns.ImageColumn.DisplayName);
                                provisionManager.deleteSiteColumn(siteColumns.DropDownColumn.DisplayName).then(function () {
                                    console.info("site column deleted: " + siteColumns.DropDownColumn.DisplayName);
                                });
                            });
                        });
                    });
                });
            });
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
