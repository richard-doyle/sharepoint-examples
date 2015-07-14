<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="AppPartContent.aspx.cs" Inherits="MembersAppPartWeb.AppPartContent" %>

<!DOCTYPE html>

<html>
    <body>
        <div id="content">
            <!-- Placeholders for properties -->
            Uri property: <span id="uriProp"></span><br />
        </div>

    <!-- Main JavaScript function, controls the rendering
         logic based on the custom property values -->
    <script lang="javascript">
        "use strict";

        var params = document.URL.split("?")[1].split("&");
        var uriProp;

        // Extracts the property values from the query string.
        for (var i = 0; i < params.length; i = i + 1) {
            var param = params[i].split("=");
            if (param[0] == "uriProp") {
                uriProp = decodeURIComponent(param[1]);
            }
        }

        // Use the URI in a call to our custom API which can query for the members.

        document.getElementById("uriProp").innerText = uriProp;
    </script>
    </body>
</html>
