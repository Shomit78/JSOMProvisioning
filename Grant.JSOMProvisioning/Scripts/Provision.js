"use strict"

var provisioning = namespace('Grant.JSOM.Provision');

provisioning.Manager = function (hostWebUrl, appWebUrl) {
    function getContext() {
        return new SP.ClientContext(appWebUrl);
    }
    function getAppContextSite(ctx) {
        var fct = new SP.ProxyWebRequestExecutorFactory(appWebUrl);
        ctx.set_webRequestExecutorFactory(fct);
        return new SP.AppContextSite(ctx, hostWebUrl);
    }
    function createSiteColumn(xmlFieldSchema) {
        var dfd = $.Deferred();

        var ctx = getContext();
        var appctx = getAppContextSite(ctx);

        var targetWeb = appctx.get_site().get_rootWeb();
        var fields = targetWeb.get_fields()
        fields.addFieldAsXml(xmlFieldSchema, false, SP.AddFieldOptions.addFieldCheckDisplayName);

        ctx.executeQueryAsync(function () {
            dfd.resolve();
        }, function (sender, args) {
            console.log("Site column creation failure: " + args.get_message());
            dfd.reject();
        });
        return dfd.promise();
    }
    var publicMembers = {
        createSiteColumnText: function (name, displayName, description, required, group) {
            var fieldSchema = '<Field Type="Text" Name="' + name + '" DisplayName="' +
                displayName + '" Description="' + description + '" Required="' + required + '" Group="' + group +
                '" SourceID="http://schemas.microsoft.com/sharepoint/v3" />';
            createSiteColumn(fieldSchema).then(function () {
                console.info("Text site column created: " + displayName);
            });
        },
        createSiteColumnNumber: function (name, displayName, description, max, min, decimals, required, group) {
            var fieldSchema = "";
            if ((max != null) || (min != null) || (decimals != null)) {
                fieldSchema += '<Field Type="Number" Name="' + name + '" DisplayName="' +
                    displayName + '" Description="' + description + '" Max="' + max + '" Min="' + min + '" Decimals="' + decimals + '" Required="' +
                    required + '" Group="' + group + '" SourceID="http://schemas.microsoft.com/sharepoint/v3" />';
            }
            else {
                fieldSchema += '<Field Type="Number" Name="' + name + '" DisplayName="' +
                    displayName + '" Description="' + description + '" Required="' +
                    required + '" Group="' + group + '" SourceID="http://schemas.microsoft.com/sharepoint/v3" />';
            }
            createSiteColumn(fieldSchema).then(function () {
                console.info("Number site column created: " + displayName);
            });
        },
        createSiteColumnUrl: function (name, displayName, description, required, group) {
            var fieldSchema = '<Field Type="URL" Format="Hyperlink" Name="' + name + '" DisplayName="' +
                displayName + '" Description="' + description + '" Required="' + required + '" Group="' + group +
                '" SourceID="http://schemas.microsoft.com/sharepoint/v3" />';
            createSiteColumn(fieldSchema).then(function () {
                console.info("Hyperlink site column created: " + displayName);
            });
        },
        createSiteColumnImage: function (name, displayName, description, required, group) {
            var fieldSchema = '<Field Type="URL" Format="Image" Name="' + name + '" DisplayName="' +
                displayName + '" Description="' + description + '" Required="' + required + '" Group="' + group +
                '" SourceID="http://schemas.microsoft.com/sharepoint/v3" />';
            createSiteColumn(fieldSchema).then(function () {
                console.info("Picture site column created: " + displayName);
            });
        },
        createSiteColumnDropDown: function (name, displayName, description, choices, required, group) {
            var fieldSchema = '<Field Type="Choice" Format="Dropdown" Name="' + name + '" DisplayName="' +
                displayName + '" Description="' + description + '" Required="' + required + '" Group="' + group +
                '" SourceID="http://schemas.microsoft.com/sharepoint/v3" ><CHOICES>';
            for (var i = 0; i < choices.length; i++) {
                console.log(choices[i]);
                fieldSchema += "<CHOICE>" + choices[i] + "</CHOICE>";
            }
            fieldSchema += "</CHOICES></Field>";
            createSiteColumn(fieldSchema).then(function () {
                console.info("Dropdown site column created: " + displayName);
            });
        },
        deleteSiteColumn: function (siteColumnDisplayName) {
            var dfd = $.Deferred();

            var ctx = getContext();
            var appctx = getAppContextSite(ctx);

            var targetWeb = appctx.get_site().get_rootWeb();
            var fields = targetWeb.get_fields()
            var field = fields.getByTitle(siteColumnDisplayName);
            field.deleteObject();

            ctx.executeQueryAsync(function () {
                dfd.resolve();
            }, function (sender, args) {
                console.log("Site column deletion failure: " + siteColumnDisplayName + " - " + args.get_message());
                dfd.reject();
            });
            return dfd.promise();
        }
    };
    return publicMembers;
}