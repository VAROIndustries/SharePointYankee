// HideSearchScopes.js
// Hide the search scopes drop-down in WSS v3/MOSS 2007
// Includes a getElementsByClassName polyfill for older browsers
//
// Blog Post: https://sharepointyankee.com/hiding-the-search-scopes-drop-down-in-wssv3moss-2007
// Author: Geoff Varosky
// Website: https://sharepointyankee.com

document.getElementsByClassName = function() {
    if (document.hasChildNodes && arguments[0]) {
        var data = new Array();
        for (a = 0; a < document.getElementsByTagName("*").length; a++) {
            if (document.getElementsByTagName("*")[a].className == arguments[0]) {
                data.push(document.getElementsByTagName("*")[a]);
            }
        }
        return data;
    }
}

document.getElementsByClassName("ms-sbscopes ms-sbcell")[0].style.display = "none";
