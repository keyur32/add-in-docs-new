
# Understanding the JavaScript API for Office


This article provides information about the JavaScript API for Office and how to use it. For reference information, see [JavaScript API for Office](http://msdn.microsoft.com/en-us/library/b27e70c3-d87d-4d27-85e0-103996273298%28Office.15%29.aspx). To run and edit some JavaScript API for Office code in your web browser with Excel Online, see the [API Tutorial for Office](http://msdn.microsoft.com/en-us/office/dn449240.aspx). For information about updating Visual Studio project files to the most current version of the JavaScript API for Office, see [Update the version of your JavaScript API for Office and manifest schema files](../../docs/overview/update-your-javascript-api-for-office-and-manifest-schema-version.md).

Explore the object model by add-in type or host: [1.1](http://msdn.microsoft.com/library/b27e70c3-d87d-4d27-85e0-103996273298%28Office.15%29.aspx)

## Referencing the JavaScript API for Office library in your add-in


The JavaScript API for Office library is implemented in the Office.js file and associated .js files that contain application-specific implementations, such as Excel-15.js and Outlook-15.js. [Reference the JavaScript API for Office library](../../docs/get-started/referencing-the-javascript-api-for-office-library-from-its-cdn.md) inside the `<head>` tag of the web page (such as an .html, .aspx, or .php file) that implements the UI of your add-in by using a `script` tag with its `src` attribute set to the following CDN URL:


```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js"/>
```

This will download and cache the JavaScript API for Office files the first time your add-in loads to make sure that it is using the most up-to-date implementation of Office.js and its associated files for the specified version.




## Initializing your add-in


 **Applies to:** All add-in types

The JavaScript API for Office provides the [Office](http://msdn.microsoft.com/library/c490b13d-ee52-4291-af5d-f4a5a11d3af0%28Office.15%29.aspx) object, which lets the developer implement a listener for the [initialize](http://msdn.microsoft.com/library/727adf79-a0b5-48d2-99c7-6642c2c334fc%28Office.15%29.aspx) event of an Office Add-in. When the API is loaded and ready for the add-in to start interacting with user's content, it triggers the **Office.initialize** event. You can use code in the **initialize** event handler to implement common add-in initialization scenarios, such as prompting the user to select some cells in Excel, and then inserting a chart initialized with those selected values. You can also use the initialize event handler to initialize other custom logic for your add-in, such as establishing bindings, prompting for default add-in settings values, and so on.

 **Important:** Even if your add-in has no initialization tasks to perform, you must include at least a minimal **Office.initialize** event handler function like the following example.




```js
Office.initialize = function () {
};
```

If you fail to include an  **Office.initialize** event handler, your add-in may raise an error when it starts. Also, if a user attempts to use your add-in with an Office Online web client, such as Excel Online, PowerPoint Online, or Outlook Web App, it will fail to run.

If your add-in includes more than one page, whenever it loads a new page that page must include or call an  **Office.initialize** event handler.

For more detail about the sequence of events when an add-in is initialized, see [Loading the DOM and runtime environment](../../docs/overview/loading-the-dom-and-runtime-environment.md).

For task pane and content add-ins (but not Outlook add-ins), the  _reason_ parameter of the **initialize** event listener function provides access to the [InitializationReason](http://msdn.microsoft.com/library/3a1fb60c-a6a7-4c73-b1d0-97096946382e%28Office.15%29.aspx) enumeration that specifies how the initialization occurred. For example, a task pane or content add-in can be initialized because the user inserted it from the Office client's ribbon UI, or because a document that already contains the add-in was opened.

You can use the value of the  **InitializationReason** enumeration to implement different logic for when the add-in is first inserted versus when it already exists in the document. The following example shows some simple logic you can add to the previous example to use the value of the _reason_ argument to display how the task pane or content add-in was initialized.




```js
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
    // Display initialization reason.
    if (reason == "inserted")
    write("The add-in was just inserted.");

    if (reason == "documentOpened")
    write("The add-in is already part of the document.");
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## Context Object


 **Applies to:** All add-in types

When an add-in is initialized, it has many different objects that it can interact with in the runtime environment. The add-in's runtime context is reflected in the API by the [Context](http://msdn.microsoft.com/library/662883d5-b86f-4bdc-99f0-9ee9129ed16c%28Office.15%29.aspx) object. The **Context** is the main object that provides access to the most important objects of the API, such as the [Document](http://msdn.microsoft.com/library/f8859516-cc1f-4b20-a8f3-cee37a983e70%28Office.15%29.aspx) and [Mailbox](https://dev.outlook.com/reference/add-ins/Office.context.mailbox.html) objects, which in turn provide access to document and mailbox content.

For example, in task pane or content add-ins, you can use the [document](http://msdn.microsoft.com/library/92351713-9ea0-43e1-b549-dad93b3208b2%28Office.15%29.aspx) property of the **Context** object to access the properties and methods of the **Document** object to interact with the content of Word documents, Excel worksheets, or Project schedules. Similarly, in Outlook add-ins, you can use the [mailbox](https://dev.outlook.com/reference/add-ins/Office.context.mailbox.html) property of the **Context** object to access the properties and methods of the **Mailbox** object to interact with the message, meeting request, or appointment content.

The  **Context** object also provides access to the [contentLanguage](http://msdn.microsoft.com/library/4fd063c2-0cd0-4b5b-8993-93d7ff8ce3bf%28Office.15%29.aspx) and [displayLanguage](http://msdn.microsoft.com/library/732ba34c-c99f-4c00-836d-4250eb7f0dac%28Office.15%29.aspx) properties that let you determine the locale (language) used in the document or item, or by the host application. And, the [roamingSettings](https://dev.outlook.com/reference/add-ins/Office.context.html) property that lets you access the members of the [RoamingSettings](https://dev.outlook.com/reference/add-ins/RoamingSettings.html) object.


## Document object


 **Applies to:** Content and task pane add-in types

To interact with document data in Excel, PowerPoint, and Word, the API provides the [Document](http://msdn.microsoft.com/library/f8859516-cc1f-4b20-a8f3-cee37a983e70%28Office.15%29.aspx) object. You can use **Document** object members to access data from the following ways:


- Read and write to active selections in the form of text, contiguous cells (matrices), or tables.
    
- Tabular data (matrices or tables).
    
- Bindings (created with the "add" methods of the  **Bindings** object).
    
- Custom XML parts (only for Word).
    
- Settings or add-in state persisted per add-in on the document.
    
You can also use the  **Document** object to interact with data in Project documents. The Project-specific functionality of the API is documented in the members [ProjectDocument](http://msdn.microsoft.com/library/1908af4f-93b9-4859-87e3-06942014fae1%28Office.15%29.aspx) abstract class. For more information about creating task pane add-ins for Project, see [Task pane add-ins for Project](../project/project-add-ins.md).

All these forms of data access start from an instance of the abstract  **Document** object.

You can access an instance of the  **Document** object when the task pane or content add-in is initialized by using the [document](http://msdn.microsoft.com/library/92351713-9ea0-43e1-b549-dad93b3208b2%28Office.15%29.aspx) property of the **Context** object. The **Document** object defines common data access functions shared across Word and Excel documents, and also provides access to the **CustomXmlParts** object for Word documents.

The  **Document** object supports four ways for developers to access document contents:


- Selection-based access
    
- Binding-based access
    
- Custom XML part-based access (Word only)
    
- Entire document-based access (PowerPoint and Word only)
    
To help you understand how selection- and binding-based data access methods work, we will first explain how the data-access APIs provide consistent data access across different Office applications.


### Consistent data access across Office applications

 **Applies to:** Content and task pane add-in types

To create extensions that seamlessly work across different Office documents, the JavaScript API for Office abstracts away the particularities of each Office application through common data types and the ability to coerce different document contents into three common data types.


#### Common data types

In both selection-based and binding-based data access, document contents are exposed through data types that are common across all the supported Office applications. In Office 2013, three main data types are supported:



|**Data type**|**Description**|**Host application support**|
|:-----|:-----|:-----|
|Text|Provides a string representation of the data in the selection or binding.|In Excel 2013, Project 2013, and PowerPoint 2013 only plain text is supported. In Word 2013, three text formats are supported: plain text, HTML, and Office Open XML (OOXML).When text is selected in a cell in Excel, selection-based methods read and write to the entire contents of the cell, even if only a portion of the text is selected in the cell. When text is selected in Word and PowerPoint, selection-based methods read and write only to the run of characters that are selected.Project 2013 and PowerPoint 2013 support only selection-based data access.|
|Matrix|Provides the data in the selection or binding as a two dimensional  **Array**, which in JavaScript is implemented as an array of arrays.For example, two rows of  **string** values in two columns would be ` [['a', 'b'], ['c', 'd']]`, and a single column of three rows would be  `[['a'], ['b'], ['c']]`.|Matrix data access is supported only in Excel 2013 and Word 2013.|
|Table|Provides the data in the selection or binding as a [TableData](http://msdn.microsoft.com/library/2183ea52-5a40-4048-b9a4-7cd66bb0ad5d%28Office.15%29.aspx) object. The **TableData** object exposes the data through the **headers** and **rows** properties.|Table data access is supported only in Excel 2013 and Word 2013.|

#### Data type coercion

The data access methods on the  **Document** and [Binding](http://msdn.microsoft.com/library/42882642-d22b-47d2-a8d3-3aa8c6a4435e%28Office.15%29.aspx) objects support specifying the desired data type using the _coercionType_ parameter of these methods, and corresponding [CoercionType](http://msdn.microsoft.com/library/735eaab6-5e31-4bc2-add5-9d378900a31b%28Office.15%29.aspx) enumeration values. Regardless of the actual shape of the binding, the different Office applications support the common data types by trying to coerce the data into the requested data type. For example, if a Word table or paragraph is selected, the developer can specify to read it as plain text, HTML, Office Open XML, or a table, and the API implementation handles the necessary transformations and data conversions.


 >**Tip**   **When should you use the matrix versus table coercionType for data access?** If you need your tabular data to grow dynamically when rows and columns are added, and you must work with table headers, you should use the table data type (by specifying the _coercionType_ parameter of a **Document** or **Binding** object data access method as `"table"` or **Office.CoercionType.Table**). Adding rows and columns within the data structure is supported in both table and matrix data, but appending rows and columns is supported only for table data. If you aren't planning on adding rows and columns, and your data doesn't require header functionality, then you should use the matrix data type (by specifying the  _coercionType_ parameter of the data access method as `"matrix"` or **Office.CoercionType.Matrix**), which provides a simpler model of interacting with the data.

If the data can't be coerced to the specified type, the [AsyncResult.status](http://msdn.microsoft.com/library/eec9c712-79eb-4365-88a1-6d77649727c1%28Office.15%29.aspx) property in the callback returns `"failed"`, and you can use the [AsyncResult.error](http://msdn.microsoft.com/library/51c46d36-972d-4d82-91aa-da99cbeb8d4f%28Office.15%29.aspx) property to access an [Error](http://msdn.microsoft.com/library/36d1d048-b888-4bb5-9321-d340bcbc86f4%28Office.15%29.aspx) object with information about why the method call failed.


## Working with selections using the Document object


The  **Document** object exposes methods that let you to read and write to the user's current selection in a "get and forget" fashion. To do that, the **Document** object provides the **getSelectedDataAsync** and **setSelectedDataAsync** methods.

For code examples that demonstrate how to perform tasks with selections, see [Read and write data to the active selection in a document or spreadsheet](../../docs/how-to/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md).


## Working with bindings using the Bindings and Binding objects


Binding-based data access enables content and task pane add-ins to consistently access a particular region of a document or spreadsheet through an identifier associated with a binding. The add-in first needs to establish the binding by calling one of the methods that associates a portion of the document with a unique identifier: [addFromPromptAsync](http://msdn.microsoft.com/library/9dc03608-b08b-4700-8be1-3c86ae236799%28Office.15%29.aspx), [addFromSelectionAsync](http://msdn.microsoft.com/library/edc99214-e63e-43f2-9392-97ead42fc155%28Office.15%29.aspx), or [addFromNamedItemAsync](http://msdn.microsoft.com/library/afbadac7-60c7-47cb-9477-6e9466ded44c%28Office.15%29.aspx). After the binding is established, the add-in can use the provided identifier to access the data contained in the associated region of the document or spreadsheet. Creating bindings provides the following value to your add-in:


- Permits access to common data structures across supported Office applications, such as: tables, ranges, or text (a contiguous run of characters).
    
- Enables read/write operations without requiring the user to make a selection.
    
- Establishes a relationship between the add-in and the data in the document. Bindings are persisted in the document, and can be accessed at a later time.
    
Establishing a binding also allows you to subscribe to data and selection change events that are scoped to that particular region of the document or spreadsheet. This means that the add-in is only notified of changes that happen within the bound region as opposed to general changes across the whole document or spreadsheet.

The [Bindings](http://msdn.microsoft.com/library/09979e31-3bfb-45be-adda-0f7cc2db1fe1%28Office.15%29.aspx) object exposes a [getAllAsync](http://msdn.microsoft.com/library/ef902b73-cc4c-4551-95de-d8a51eeba82f%28Office.15%29.aspx) method that gives access to the set of all bindings established on the document or spreadsheet. An individual binding can be accessed by its ID using either the [Bindings.getBindingByIdAsync](http://msdn.microsoft.com/library/2727c891-bc05-465c-9324-113fbfeb3fbb%28Office.15%29.aspx) or [Office.select](http://msdn.microsoft.com/library/23aeb136-da1f-4127-a798-99dc27bc4dae%28Office.15%29.aspx) methods. You can establish new bindings as well as remove existing ones by using one of the following methods of the **Bindings** object: [addFromSelectionAsync](http://msdn.microsoft.com/library/edc99214-e63e-43f2-9392-97ead42fc155%28Office.15%29.aspx), [addFromPromptAsync](http://msdn.microsoft.com/library/9dc03608-b08b-4700-8be1-3c86ae236799%28Office.15%29.aspx), [addFromNamedItemAsync](http://msdn.microsoft.com/library/afbadac7-60c7-47cb-9477-6e9466ded44c%28Office.15%29.aspx), or [releaseByIdAsync](http://msdn.microsoft.com/library/ad285984-8b44-435d-9b84-f0ade570c896%28Office.15%29.aspx).

There are three different types of bindings that you specify with the  _bindingType_ parameter when you create a binding with the **addFromSelectionAsync**, **addFromPromptAsync** or **addFromNamedItemAsync** methods:



|**Binding type**|**Description**|**Host application support**|
|:-----|:-----|:-----|
|Text binding|Binds to a region of the document that can be represented as text.|In Word, most contiguous selections are valid, while in Excel only single cell selections can be the target of a text binding. In Excel, only plain text is supported. In Word, three formats are supported: plain text, HTML, and Open XML for Office.|
|Matrix binding|Binds to a fixed region of a document that contains tabular data without headers.Data in a matrix binding is written or read as a two dimensional  **Array**, which in JavaScript is implemented as an array of arrays. For example, two rows of  **string** values in two columns can be written or read as ` [['a', 'b'], ['c', 'd']]`, and a single column of three rows can be written or read as  `[['a'], ['b'], ['c']]`.|In Excel, any contiguous selection of cells can be used to establish a matrix binding. In Word, only tables support matrix binding.|
|Table binding|Binds to a region of a document that contains a table with headers.Data in a table binding is written or read as a [TableData](http://msdn.microsoft.com/library/2183ea52-5a40-4048-b9a4-7cd66bb0ad5d%28Office.15%29.aspx) object. The **TableData** object exposes the data through the **headers** and **rows** properties.|Any Excel or Word table can be the basis for a table binding. After you establish a table binding, each new row or column a user adds to the table is automatically included in the binding. |
After a binding is created by using one of the three "add" methods of the  **Bindings** object, you can work with the binding's data and properties by using the methods of the corresponding object: [MatrixBinding](http://msdn.microsoft.com/library/35e8568e-9129-4c00-b30f-d8c3b2555f1e%28Office.15%29.aspx), [TableBinding](http://msdn.microsoft.com/library/1508795b-1c70-456c-b3bf-666d40cf8f50%28Office.15%29.aspx), or [TextBinding](http://msdn.microsoft.com/library/6b71b21d-f64d-425c-99d9-c62b2a9969be%28Office.15%29.aspx). All three of these objects inherit the [getDataAsync](http://msdn.microsoft.com/library/5372ffd8-579d-4fcb-9e5b-e9a2128f3201%28Office.15%29.aspx) and [setDataAsync](http://msdn.microsoft.com/library/6a59bb6d-40b6-4a95-9b98-d70d4616de09%28Office.15%29.aspx) methods of the **Binding** object that enable to you interact with the bound data.

For code examples that demonstrate how to perform tasks with bindings, see [Bind to regions in a document or spreadsheet](../../docs/how-to/bind-to-regions-in-a-document-or-spreadsheet.md).


## Working with custom XML parts using the CustomXmlParts and CustomXmlPart objects


 **Applies to:** Task pane add-ins for Word

The [CustomXmlParts](http://msdn.microsoft.com/library/ba40cd4c-29bb-4f31-875d-6f1382fd1ee8%28Office.15%29.aspx) and [CustomXmlPart](http://msdn.microsoft.com/library/83f0e668-8236-4f2f-a20f-b173a9e3f65f%28Office.15%29.aspx) objects of the API provide access to custom XML parts in Word documents, which enable XML-driven manipulation of the contents of the document. For demonstrations of working with the **CustomXmlParts** and **CustomXmlPart** objects, see the [Word-Add-in-Work-with-custom-XML-parts](https://github.com/OfficeDev/Word-Add-in-Work-with-custom-XML-parts) code sample.


## Working with the entire document using the getFileAsync method


 **Applies to:** Task pane add-ins for Word and PowerPoint

The [Document.getFileAsync](http://msdn.microsoft.com/library/78047418-89c4-4c7d-9427-4735b8559518%28Office.15%29.aspx) method and members of the [File](http://msdn.microsoft.com/library/04923ddf-8efa-459f-aed5-d8c06385ca50%28Office.15%29.aspx) and [Slice](http://msdn.microsoft.com/library/011b5647-639b-4b06-8625-ba9de01bed4b%28Office.15%29.aspx) objects to provide functionality for getting entire Word and PowerPoint document files in slices (chunks) of up to 4 MB at a time. For more information, see [How to: Get all file content from a document in an add-in](../../docs/how-to/get-the-whole-document-from-an-add-in-for-powerpoint-or-word.md).


## Mailbox object


 **Applies to:** Outlook add-ins

Outlook add-ins primarily use a subset of the API exposed through the [Mailbox](https://dev.outlook.com/reference/add-ins/Office.context.mailbox.html) object. To access the objects and members specifically for use in Outlook add-ins, such as the [Item](https://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html) object, you use the [mailbox](https://dev.outlook.com/reference/add-ins/Office.context.mailbox.html) property of the **Context** object to access the **Mailbox** object, as shown in the following line of code.




```js
// Access the Item object.
var item = Office.context.mailbox.item;

```

Additionally, Outlook add-ins can use the following objects:


-  **Office** object: for initialization.
    
-  **Context** object: for access to content and display language properties.
    
-  **RoamingSettings** object: for saving Outlook add-in-specific custom settings to the user's mailbox where the add-in is installed.
    
For information about using JavaScript in Outlook add-ins, see [Outlook add-ins](../outlook/outlook-add-ins.md) and [Overview of Outlook add-ins architecture and features](../outlook/overview.md).


## API support matrix


This table summarizes the API and features supported across add-in types (content, task pane, and Outlook) and the Office applications that can host them when you specify the [Office host applications your add-in supports](http://msdn.microsoft.com/library/cff9fbdf-a530-4f6e-91ca-81bcacd90dcd%28Office.15%29.aspx) using the [1.1 add-in manifest schema and features supported by v1.1 JavaScript API for Office](../../docs/overview/update-your-javascript-api-for-office-and-manifest-schema-version.md).


|||||||||
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
||**Host name**|Database|Workbook|Mailbox|Presentation|Document|Project|
||**Supported** **Host applications**|Access web apps|ExcelExcel Online|OutlookOutlook Web AppOWA for Devices|PowerPointPowerPoint Online|Word|Project|
|**Supported add-in types**|Content|Y|Y||Y|||
||Task pane||Y||Y|Y|Y|
||Outlook|||Y||||
|**Supported API features**|Read/Write Text||Y||Y|Y|Y (Read only)|
||Read/Write Matrix||Y|||Y||
||Read/Write Table||Y|||Y||
||Read/Write HTML|||||Y||
||Read/WriteOffice Open XML|||||Y||
||Read task, resource, view, and field properties||||||Y|
||Selection changed events||Y|||Y||
||Get whole document||||Y|Y||
||Bindingsand binding events|Y (Only full and partialtable bindings)|Y|||Y||
||Read/WriteCustom Xml Parts|||||Y||
||Persist add-in state data(settings)|Y (Per host add-in)|Y (Per document)|Y (Per mailbox)|Y (Per document)|Y (Per document)||
||Settings changed events|Y|Y||Y|Y||
||Get active view modeand view changed events||||Y|||
||Navigate to locationsin the document||Y||Y|Y||
||Activate contextuallyusing rules and RegEx|||Y||||
||Read Item properties|||Y||||
||Read User profile|||Y||||
||Get attachments|||Y||||
||Get User identity token|||Y||||
||Call Exchange Web Services|||Y||||
