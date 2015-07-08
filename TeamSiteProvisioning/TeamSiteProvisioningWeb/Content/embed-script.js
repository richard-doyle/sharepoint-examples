var jQuery = "https://ajax.aspnetcdn.com/ajax/jQuery/jquery-2.0.2.min.js";

// Is MDS enabled?
if ("undefined" != typeof g_MinimalDownload && g_MinimalDownload && (window.location.pathname.toLowerCase()).endsWith("/_layouts/15/start.aspx") && "undefined" != typeof asyncDeltaManager) {
    // Register script for MDS if possible
    RegisterModuleInit("scenario1.js", JavaScript_Embed); //MDS registration
    JavaScript_Embed(); //non MDS run
} else {
    JavaScript_Embed();
}

function JavaScript_Embed() {

    loadScript(jQuery, function () {
        $(document).ready(function () {
            SP.SOD.executeOrDelayUntilScriptLoaded(function () { insertElem(); }, 'sp.js');
        });
    });
}

function insertElem() {
    var button = $('#suiteBarTop .o365cs-nav-appTitle.o365cs-topnavText.o365button');
    var newElem = button.clone();

    var span = newElem.find('span').html("Dashboard");
    newElem.attr('href', "http://www.google.com");
    button.before(newElem);
}

function loadScript(url, callback) {
    var head = document.getElementsByTagName("head")[0];
    var script = document.createElement("script");
    script.src = url;

    // Attach handlers for all browsers
    var done = false;
    script.onload = script.onreadystatechange = function () {
        if (!done && (!this.readyState
					|| this.readyState == "loaded"
					|| this.readyState == "complete")) {
            done = true;

            // Continue your code
            callback();

            // Handle memory leak in IE
            script.onload = script.onreadystatechange = null;
            head.removeChild(script);
        }
    };

    head.appendChild(script);
}