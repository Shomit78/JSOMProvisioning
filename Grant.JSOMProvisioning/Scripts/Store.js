"use strict"

var store = namespace('Grant.JSOM.Store');

store.SiteColumns = function () {
    this.groupName = "Grant JSOM Site Columns";
    this.SimpleTextColumn = {
        Name: "JSOMTextField",
        DisplayName: "JSOM Text Field",
        Description: "This field was created using JSOM",
        Required: "TRUE"
    };
}