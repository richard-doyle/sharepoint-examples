'use strict';

var context = SP.ClientContext.get_current();
var groups = context.get_web().get_siteGroups();
var membersGroup = groups.getByName("developer Members");
var members = membersGroup.get_users();

// This code runs when the DOM is ready and creates a context object which is needed to use the SharePoint object model
$(document).ready(function () {
    getMembers();
});

// This function prepares, loads, and then executes a SharePoint query to get the current users information
function getMembers() {
    context.load(members);
    context.executeQueryAsync(onGetMembersSuccess, onGetMembersFail);
}

function onGetMembersSuccess() {
    console.log(members);

    var userInfo = '';

    var userEnumerator = members.getEnumerator();
    while (userEnumerator.moveNext()) {
        var oUser = userEnumerator.get_current();
        console.log(oUser);
        userInfo += '<br/><br/>User: ' + oUser.get_title() +
            '<br/>ID: ' + oUser.get_id() +
            '<br/>Email: ' + oUser.get_email() +
            '<br/>Login Name: ' + oUser.get_loginName();
    }

    var messageElem = document.getElementById("message");
    message.innerHTML = userInfo;
}

// This function is executed if the above call fails
function onGetMembersFail(sender, args) {
    alert('Failed to get members. Error:' + args.get_message());
}
