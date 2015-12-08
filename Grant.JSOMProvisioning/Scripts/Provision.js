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
    function constructLCI(listTitle, listTemplateType) {
        var lci = new SP.ListCreationInformation();
        lci.set_title(listTitle);
        lci.set_templateType(listTemplateType);
        return lci;
    }
    function constructCustomLCI(listTitle) {
        var lci = constructLCI(listTitle, SP.ListTemplateType.genericList);
        return lci;
    }
    function constructCTCI(id, name, group, description) {
        var ctci = new SP.ContentTypeCreationInformation();
        ctci.set_description(description);
        ctci.set_group(group);
        ctci.set_id(id);
        ctci.set_name(name);
        return ctci;
    }
    function constructFLCI(targetField) {
        var flci = new SP.FieldLinkCreationInformation();
        flci.set_field(targetField);
        return flci;
    }
    var publicMembers = {
        createSiteColumn:  function(xmlFieldSchema) {
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
        },
        createSiteColumnTextFieldXml: function (name, displayName, description, required, group) {
            var fieldSchema = '<Field Type="Text" Name="' + name + '" DisplayName="' +
                displayName + '" Description="' + description + '" Required="' + required + '" Group="' + group +
                '" SourceID="http://schemas.microsoft.com/sharepoint/v3" />';
            return fieldSchema
        },
        createSiteColumnNumberFieldXml: function (name, displayName, description, max, min, decimals, required, group) {
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
            return fieldSchema;
        },
        createSiteColumnUrlFieldXml: function (name, displayName, description, required, group) {
            var fieldSchema = '<Field Type="URL" Format="Hyperlink" Name="' + name + '" DisplayName="' +
                displayName + '" Description="' + description + '" Required="' + required + '" Group="' + group +
                '" SourceID="http://schemas.microsoft.com/sharepoint/v3" />';
            return fieldSchema;
        },
        createSiteColumnImageFieldXml: function (name, displayName, description, required, group) {
            var fieldSchema = '<Field Type="URL" Format="Image" Name="' + name + '" DisplayName="' +
                displayName + '" Description="' + description + '" Required="' + required + '" Group="' + group +
                '" SourceID="http://schemas.microsoft.com/sharepoint/v3" />';
            return fieldSchema;
        },
        createSiteColumnDropDownFieldXml: function (name, displayName, description, choices, required, group) {
            var fieldSchema = '<Field Type="Choice" Format="Dropdown" Name="' + name + '" DisplayName="' +
                displayName + '" Description="' + description + '" Required="' + required + '" Group="' + group +
                '" SourceID="http://schemas.microsoft.com/sharepoint/v3" ><CHOICES>';
            for (var i = 0; i < choices.length; i++) {
                console.log(choices[i]);
                fieldSchema += "<CHOICE>" + choices[i] + "</CHOICE>";
            }
            fieldSchema += "</CHOICES></Field>";
            return fieldSchema;
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
        },
        //Create content type
        createSiteContentType: function (contentTypeId, contentTypeName, contentTypeGroup, contentTypeDescription, siteColumnNames) {
            var dfd = $.Deferred();

            var ctx = getContext();
            var appctx = getAppContextSite(ctx);

            var targetWeb = appctx.get_site().get_rootWeb();
            var fields = targetWeb.get_fields()
            var fieldLinks = {};
            ctx.load(fields);
            for (var i = 0; i < siteColumnNames.length; i++) {
                var field = fields.getByInternalNameOrTitle(siteColumnNames[i]);
                ctx.load(field);
                $(fieldLinks).push(field);
            }
            var ctci = constructCTCI(contentTypeId, contentTypeName)
            var newType = targetWeb.get_contentTypes().add(ctci);
            ctx.load(newType);

            ctx.executeQueryAsync(succeed, fail);
            function succeed(sender, args) {
                var fieldRefs = newType.get_fieldLinks();
                ctx.load(fieldRefs);
                ctx.executeQueryAsync(
                    function () {
                        for (var i = 0; i < fieldLinks.length; i++) {
                            var flci = constructFLCI(fieldLinks[i]);
                            newType.get_fieldLinks().add(flci);
                        }
                        newType.update();
                        ctx.executeQueryAsync(function () { dfd.resolve(); },
                            function (sender, args) {
                                console.log("Content type creation failure: " + args.get_message());
                                dfd.reject();
                            });
                    },
                    function (sender, args) {
                        console.log("Content type creation failure: " + args.get_message());
                        dfd.reject();
                    });
            }
            function fail(sender, args) {
                console.log("Content type creation failure: " + args.get_message());
                dfd.reject();
            }
            return dfd.promise();
        },
        //Delete content type
        deleteSiteContentType: function (contentTypeId) {
            var dfd = $.Deferred();

            var ctx = getContext();
            var appctx = getAppContextSite(ctx);

            var targetWeb = appctx.get_site().get_rootWeb();
            var webTypes = targetWeb.get_contentTypes();
            var targetType = webTypes.getById(contentTypeId)
            targetType.deleteObject();
            ctx.executeQueryAsync(succeed, fail);
            function succeed() { dfd.resolve(); }
            function fail(sender, args) {
                console.log("Content type deletion failure: " + args.get_message());
                dfd.reject();
            }
            return dfd.promise();
        }
        //Create list
        //Delete list
    };
    return publicMembers;
}