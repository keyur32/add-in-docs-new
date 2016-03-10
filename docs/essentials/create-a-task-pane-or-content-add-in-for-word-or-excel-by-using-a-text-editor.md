
# Create a task pane or content add-in for Word or Excel by using a text editor


The simplest Office Add-in for Excel 2013 or Word 2013 consists of a manifest XML file that points to a webpage or website. This article shows you how to create a simple add-in that consists of only an XML manifest and an HTML file with a text editor.

To implement this add-in, you need to create the following:


- An HTML file that implements the UI of the add-in.
    
- An XML manifest file that defines the metadata required to display and run the add-in in Word or Excel.
    
- A CSS file to define a style sheet for the add-in.
    
- A project.js file that contains JavaScript programming logic that can use the JavaScript API for Office (Office.js) to perform data access operations against the content in the Word or Excel document.
    

## Create a Hello World Office Add-in


The UI of the add-in is provided by an HTML file that can optionally provide JavaScript programming logic. This first set of steps will show you how to create a Hello World add-in without programming logic. After you complete a Hello World add-in, we'll show you how to add some programming logic that interacts with the document or worksheet content.


### To create the files for a Hello World add-in


1. Create a folder on your local drive named HelloWorld (for example C:\HelloWorld). Save all of the files created in the following steps into this folder.
    
2. Create a file named HelloWorld.html that contains the following HTML code.
    
    ```HTML
    <!DOCTYPE html>
    <html>
        <head>
            <meta charset="UTF-8" />
            <meta http-equiv="X-UA-Compatible" content="IE=Edge"/>
            <link rel="stylesheet" type="text/css" href="program.css" />
        </head>
        <body>
            <p>Hello World!</p>
        </body>
    </html>
    
    ```


    This file provides the minimum set of HTML tags to display the UI of an add-in.
    
3. Create a file named program.css that contains the following CSS code.
    
    ```
     body
    {
        position:relative;
    }
    li :hover
    {  
        text-decoration: underline;
        cursor:pointer;
    }
    h1,h3,h4,p,a,li
    {
        font-family: "Segoe UI Light","Segoe UI",Tahoma,sans-serif;
        text-decoration-color:#4ec724;
    }
    
    ```


    This file provides the style sheet for the add-in.
    
4. Create an XML file named HelloWorld.xml that contains the following XML code.
    
     >**Important**  Replace the value in the  `<id>` tag with a GUID that you have generated yourself.

    ```XML
    <?xml version="1.0" encoding="utf-8"?>
    <OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
        xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
        xsi:type="TaskPaneApp">
      <Id>08afd7fe-1631-42f4-84f1-5ba51e242f98</Id>
      <Version>1.0</Version>
      <ProviderName>Microsoft</ProviderName>
      <DefaultLocale>EN-US</DefaultLocale>
      <DisplayName DefaultValue="Hello World add-in"/>
      <Description DefaultValue="My first app."/>
      <IconUrl DefaultValue=
        "http://officeimg.vo.msecnd.net/_layouts/images/general/office_logo.jpg"/>
    
      <Hosts>
        <Host Name="Document"/>
        <Host Name="Workbook"/>
      </Hosts>
    
      <DefaultSettings>
        <SourceLocation DefaultValue="\\MyShare\MyManifests\HelloWorld\HelloWorld.html"/>
      </DefaultSettings>
      <Permissions>ReadWriteDocument</Permissions>
    </OfficeApp>
    
    ```


    This file provides the manifest XML file for the add-in.
    
The next two procedures describe how to copy your files to a network share, and then specify that location as a trusted add-in catalog, so that you can test your add-in.


### To specify a trusted location for the manifest


1. Create a folder on a network share (for example, \\MyShare\MyManifests). Be sure that the  `<SourceLocation>` element of the HelloWorld.xml manifest file points to this location for the .html page of the add-in. Save all of the files you created into this folder.
    
     >**Note**  Alternatively, you can save only the HelloWorld.xml manifest file to this share, and then put the .html file on a web server. If you do that, be sure that the  `<SourceLocation>` element of the HelloWorld.xml manifest file points to the URL of the HelloWorld.html file on that server.
2. Open a new document in Excel or Word.
    
3. Choose the  **File** tab, and then choose **Options**.
    
4. Choose  **Trust Center**, and then choose the  **Trust Center Settings** button.
    
5. Choose  **Trusted Add-in Catalogs**.
    
6. In the  **Catalog Url** box, enter the path to the network share you created in Step 1, and then choose ** Add Catalog**.
    
7. Select the  **Show in Menu** check box, and then choose **OK**.
    
    A message is displayed to inform you that your settings will be applied the next time you start Office.
    
8. Close and restart Excel or Word.
    

### To test and run the Hello World add-in


1. On the  **Insert** tab, choose **My Add-ins**. 
    
2. In the  **Office Add-ins** dialog box, choose **Shared Folder**.
    
3. Choose  **Hello World add-in**, and then choose  **Start**.
    
    The add-in will open in a task pane to the right of the current document or worksheet.
    

## Add programming logic to the Hello World add-in


The next set of steps will show you how to add some basic programming logic to the Hello World add-in so that it can interact with the document or worksheet content.


### To add programming logic to the Hello World add-in


1. Open the HelloWorld.html file and add the  `<script>` tags inside the `<head>` tags of the file.
    
    ```HTML
    <!DOCTYPE html>
    <html>
        <head>
            <meta charset="UTF-8" />
            <meta http-equiv="X-UA-Compatible" content="IE=Edge"/>
            <link rel="stylesheet" type="text/css" href="program.css" />
            <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js"></script>
            <script src="http://ajax.aspnetcdn.com/ajax/jquery/jquery-1.9.0.min.js"></script>
            <script src="Program.js"></script>
        </head>
        <body>
            <p>Hello World!</p>
        </body>
    </html>
    ```


    This adds a reference to the Office.js library file that implements the JavaScript API for Office. It also adds a reference to Program.js, which is a file we'll create to contain the programming logic for the add-in.
    
     >**Note**  The  `src` attribute of the `<script>` tag references the JavaScript API for Office (office.js) that will be externally available.
2. Replace  `<p>Hello World!</p>` with the lines inside the `<body>` tags of the file.
    
    ```HTML
    <!DOCTYPE html>
    <html xmlns="http://www.w3.org/1999/xhtml">
        <head>
            <meta charset="UTF-8" />
            <meta http-equiv="X-UA-Compatible" content="IE=Edge"/>
            <link rel="stylesheet" type="text/css" href="program.css" />
            <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js"></script>
            <script src="Program.js"></script>
        </head>
        <body>
            <button onclick="writeData()"> Write Data </button></br>
            <button onclick="ReadData()"> Read Selected Data </button></br>
            Results: <div id="results"></div>
        </body>
    </html>
    
    ```


    This adds two buttons to the add-in UI and defines a  `div` to display results.
    
3. Create a file named Program.js that contains the following JavaScript code.
    
     >**Note**  [Office.initialize](http://msdn.microsoft.com/en-us/library/727adf79-a0b5-48d2-99c7-6642c2c334fc) must be initialized as a function at the beginning of the code file so that the [Office.context](http://msdn.microsoft.com/en-us/library/6c4b2c16-d4fb-4ecf-b72c-1e33b205daaf) property will be available when called from the functions that follow.

    ```js
    // The initialize function is required for all add-ins.
    Office.initialize = function (reason) {
        // Checks for the DOM to load using the jQuery ready function.
        $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Add any initialization logic to this function.
        });
    }
    var MyArray = [['Berlin'],['Munich'],['Duisburg']];
    
    function writeData() {
        Office.context.document.setSelectedDataAsync(MyArray, { coercionType: 'matrix' });
    }
    
    function ReadData() {
        Office.context.document.getSelectedDataAsync("matrix", function (result) {
            if (result.status === "succeeded"){
                printData(result.value);
            }
    
            else{
                printData(result.error.name + ":" + err.message);
            }
        });
    }
    
    function printData(data) {
        {
            var printOut = "";
    
            for (var x = 0 ; x < data.length; x++) {
                for (var y = 0; y < data[x].length; y++) {
                    printOut += data[x][y] + ",";
                }
            }
           document.getElementById("results").innerText = printOut;
        }
    }
    
    ```

4. Redeploy the add-in files as described in the "To specify a trusted location for the manifest" procedure.
    
5. Insert and test the add-in as described in the "To test and run the Hello World add-in" procedure.
    

## Additional resources



- [Task pane and content add-ins for Office 2013](../../docs/essentials/task-pane-and-content-add-ins.md)
    
- [Design guidelines for Office Add-ins](../../docs/design/add-in-design.md)
    
- [Office Add-ins development lifecycle](../../docs/design/add-in-development-lifecycle.md)
    
- [Publish your Office Add-in](../publish/publish.md)
    
- [Understanding the JavaScript API for Office](../../docs/overview/understanding-the-javascript-api-for-office.md)
    
- [Office Add-ins XML manifest](../../docs/overview/add-in-manifests.md)
    
- [Office Add-ins API and schema references](../../reference/reference.md)
    
