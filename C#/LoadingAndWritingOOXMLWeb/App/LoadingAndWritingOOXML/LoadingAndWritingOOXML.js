/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/
// This function is run when the app is ready to start interacting with the host application
// It ensures the DOM is ready before adding click handlers to buttons
Office.initialize = function (reason) {
    $(document).ready(function () {
        // If setSelectedDataAsync method is supported by the host application
        // the UI buttons are hooked up to call the method else the buttons are removed

        if (Office.context.document.setSelectedDataAsync) {

            clickHandler();

        }

        else {
            $('#setFText').remove();
            $('#setSText').remove();
            $('#setImage').remove();
            $('#setBox').remove();
            $('#setShape').remove();
            $('#setControl').remove();
            $('#setFTable').remove();
            $('#setSTable').remove();
            $('#setSmartArt').remove();
            $('#setChart').remove();
        }
    });
};


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

//Replace all Content Controls with the appropriate text
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

function addComponents() {
    $('#addComponentsButton').addClass('active');
    $('#addComponents').removeClass('display-none');
    $('#changeValuesButton').removeClass('active');
    $('#changeValues').addClass('display-none');

}

function changeValues() {
    $('#changeValuesButton').addClass('active');
    $('#changeValues').removeClass('display-none');
    $('#addComponentsButton').removeClass('active');
    $('#addComponents').addClass('display-none');

    //Ensure all place holder text is a content control
    createContentControls();
}


//Specifiy What method each button should call
function clickHandler() {
    $('#addCoverLetter').click(function () { writeContent('../../documents/executive/CoverLetter.xml'); });
    $('#addExecutiveSummary').click(function () { writeContent('../../documents/executive/ExecutiveSummary.xml'); });

    $('#addProfile').click(function () { writeContent('../../documents/qualifications/ProfileOfGGI.xml'); });
    $('#addTeamText').click(function () { writeContent('../../documents/qualifications/TeamIntroductoryText.xml'); });
    $('#addOrgChart').click(function () { writeContent('../../documents/qualifications/OrganizationalChart.xml'); });
    $('#addRoles').click(function () { writeContent('../../documents/qualifications/Roles.xml'); });

    $('#createContentControls').click(createContentControls);
    $('#changeContentControl').click(function () { changeContentControl("client", $('#clientName').val()); });

    $('#addComponentsButton').click(function () { addComponents(); });
    $('#changeValuesButton').click(function () { changeValues(); });
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
