/*
 * CONSTANTS
 */
let DELETEBUTTON = '<a class="delete-parent" onclick="deleteParent(this)">&times;</a>';
//Constants for section titles
let COVER_LETTER = "Cover Letter";
let EXECUTIVE_SUMMARY = "Executive Summary";
let PROFILE = "Profile of GGI";
let TEAM = "Team Introductory Text";
let ORG_CHART = "Organizational Chart";
let ROLES = "Roles of Team Members";



/*
 * INITIALIZE
 */
// This function is run when the app is ready to start interacting with the host application
// It ensures the DOM is ready before adding click handlers to buttons
Office.initialize = function (reason) {
    $(document).ready(function () {
        // hookup UI buttons
        clickHandler();
    });
}




/*
* WRITE FILES TO DOC
*/
//Add specified file to the end of the document
function writeContent(fileName) {
    var myOOXMLRequest = new XMLHttpRequest();
    var myXML;
    myOOXMLRequest.open('GET', fileName, false);
    myOOXMLRequest.send();
    if (myOOXMLRequest.status === 200) {
        myXML = myOOXMLRequest.responseText;
    }
    Word.run(function (ctx) {
        var body = ctx.document.body;
        body.insertOoxml(myXML, Word.InsertLocation.end);
        body.insertBreak("Next", Word.InsertLocation.end);

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return ctx.sync().then(function () {
            console.log('Added ' + filename + " to end of document.");
        });
    })
        .catch(function (error) {
            console.log('Error: ' + JSON.stringify(error));
            if (error instanceof OfficeExtension.Error) {
                console.log('Debug info: ' + JSON.stringify(error.debugInfo));
            }
        });
}


//Write all files from drag and drop list to document
function createDocument() {
    var strItems = [];

    $("#sortable").children().each(function (i) {
        var li = $(this);
        let str = li.context.innerText;
        strItems.push(str.substring(0, str.length - 1));
    });

    $.each(strItems, function (str) {
        switch (strItems[str]) {
            case COVER_LETTER:
                writeContent('../../documents/executive/CoverLetter.xml');
                break;
            case EXECUTIVE_SUMMARY:
                writeContent('../../documents/executive/ExecutiveSummary.xml');
                break;
            case PROFILE:
                writeContent('../../documents/qualifications/ProfileOfGGI.xml');
                break;
            case TEAM:
                writeContent('../../documents/qualifications/TeamIntroductoryText.xml');
                break;
            case ORG_CHART:
                writeContent('../../documents/qualifications/OrganizationalChart.xml');
                break;
            case ROLES:
                writeContent('../../documents/qualifications/Roles.xml');
                break;
            default:
                break;
        }
    });
}


//Add to drag and drop list
function addToList(fileTitle) {
    let li = $('<li class="ui-state-default">' + fileTitle + DELETEBUTTON + '</li>');
    $('#sortable').append(li);
}

//Remove from drag and drop list
function deleteParent(x) {
    $(x).parent().remove();
}




/*
* CONTENT CONTROLS
*/
//Go through document and make all place-holder text into a Content Control
function createContentControls() {
    Word.run(function (ctx) {
        // Queue a command to search the document for the string "Contoso".
        // Create a proxy search results collection object.
        var results = ctx.document.body.search("[CLIENT NAME]");

        // Queue a command to load all of the properties on the search results collection object.
        ctx.load(results);

        // Synchronize the document state by executing the queued commands,
        // and returning a promise to indicate task completion.        
        return ctx.sync().then(function () {

            // Once we have the results, we iterate through each result and set some properties on
            // each search result proxy object. Then we queue a command to wrap each search result
            // with a content control and set the tag and title property on the content control.
            for (var i = 0; i < results.items.length; i++) {
                var cc = results.items[i].insertContentControl();
                cc.tag = "client";  // This value is used in another part of this sample.
                cc.title = "Client Name";
            }
        })
            // Synchronize the document state by executing the queued commands.
            .then(ctx.sync)
            .then(function () {
                handleSuccess();
            })
            .catch(function (error) {
                handleError(error);
            })
    });
}

//Replace one type of Content Control with the appropriate text
function changeContentControl(control, newText) {
    Word.run(function (ctx) {
        var ccs = ctx.document.contentControls.getByTag(control);
        ctx.load(ccs, 'tag');
        return ctx.sync().then(function () {
            for (var i = 0; i < ccs.items.length; i++) {
                ccs.items[i].insertText(newText, "Replace");
            }
        });
    })
        .catch(function (error) {
            console.log('Error: ' + JSON.stringify(error));
            if (error instanceof OfficeExtension.Error) {
                console.log('Debug info: ' + JSON.stringify(error.debugInfo));
            }
        });
}

//Replace all Content Controls with the appropriate text
function changeAllContentControls() {
    //Ensure Content Controls are up to date
    createContentControls();

    //replace content controls with correct text;
    changeContentControl("client", $('#clientName').val());
}




/*
* NAVIGATION
*/

//Go to Add Components Tab
function addComponents() {
    $('#addComponentsTabLink').addClass('active');
    $('#addComponents').removeClass('display-none');
    $('#changeValuesTabLink').removeClass('active');
    $('#changeValues').addClass('display-none');

}

//Go to Change Values Tab
function changeValues() {
    $('#changeValuesTabLink').addClass('active');
    $('#changeValues').removeClass('display-none');
    $('#addComponentsTabLink').removeClass('active');
    $('#addComponents').addClass('display-none');

    //Ensure all place holder text is a content control
    createContentControls();
}


/*
 * CLICK HANDLER 
 */
//Specifiy What method each button should call
function clickHandler() {
    $('#addCoverLetter').click(function () { addToList(COVER_LETTER); });
    $('#addExecutiveSummary').click(function () { addToList(EXECUTIVE_SUMMARY); });

    $('#addProfile').click(function () { addToList(PROFILE); });
    $('#addTeamText').click(function () { addToList(TEAM); });
    $('#addOrgChart').click(function () { addToList(ORG_CHART); });
    $('#addRoles').click(function () { addToList(ROLES); });

    $('#createContentControls').click(createContentControls);
    $('#changeContentControl').click(function () { changeAllContentControls(); });

    $('#addComponentsTabLink').click(function () { addComponents(); });
    $('#changeValuesTabLink').click(function () { changeValues(); });

    $('#createDocument').click(function () { createDocument(); })
}




// *********************************************************
//
// Word-Add-in-Load-and-write-Open-XML, https://github.com/OfficeDev/Word-Add-in-Load-and-write-Open-XML
//
// Copyright (c) Microsoft Corporation
// All rights reserved.
//
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
//
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
//
// *********************************************************
