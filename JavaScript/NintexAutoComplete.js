// NintexAutoComplete.js
// Replace a Nintex Forms drop-down list with an autocompleting textbox
// Fix for Nintex Forms version 1.11.4.0+ where the original script broke
//
// Blog Post: https://sharepointyankee.com/replacing-a-drop-down-list-in-nintex-forms-2010-with-an-autocompleting-textbox-fix-for-version-1-11-4-0-update
// Author: Geoff Varosky
// Website: https://sharepointyankee.com

NWF$(document).ready(function() {
    var textbox = NWF$("#" + mytext);
    mylist = mylist.replace("_hid", "");
    var dropDown1 = NWF$("#" + mylist);

    textbox.autocomplete({
        source: function(request, response) {
            var autocompleteVals = [];
            var dropDownOptions = "#" + mylist + " > option";
            NWF$(dropDownOptions).each(function() {
                if (NWF$(this).text() != "(None)" &&
                    NWF$(this).text().toLowerCase().indexOf(request.term.toLowerCase()) >= 0) {
                    autocompleteVals.push(NWF$(this).text());
                }
            });
            response(autocompleteVals);
        },
        minLength: 1,
        select: function(event, ui) {
            var fieldOption = NWF$("#" + dropDown1Id + " option").filter(function() {
                return NWF$(this).html() == ui.item.value;
            });
            NWF$(fieldOption).attr("selected", true);
            NWF$(dropDown1).change();
        }
    });
});
