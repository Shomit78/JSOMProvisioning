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
    this.NumberColumn = {
        Name: "JSOMNumberField",
        DisplayName: "JSOM Number Field",
        Description: "This field was created using JSOM",
        Required: "TRUE",
        Max: "100",
        Min: "1",
        Decimals: "0"
    };
    this.UrlColumn = {
        Name: "JSOMHyperlinkField",
        DisplayName: "JSOM Hyperlink Field",
        Description: "This field was created using JSOM",
        Required: "TRUE"
    };
    this.ImageColumn = {
        Name: "JSOMImageField",
        DisplayName: "JSOM Image Field",
        Description: "This field was created using JSOM",
        Required: "TRUE"
    };
    this.DropDownColumn = {
        Name: "JSOMDropDownField",
        DisplayName: "JSOM DropDown Field",
        Description: "This field was created using JSOM",
        Required: "TRUE",
        Choices: [ "Choice 1", "Choice 2", "Choice 3" ]
    };
}