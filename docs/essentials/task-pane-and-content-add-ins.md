
# Task pane and content add-ins for Office 2013



## JavaScript API support for content and task pane add-ins


This section briefly describes the subset of the [JavaScript API for Office](http://msdn.microsoft.com/library/b27e70c3-d87d-4d27-85e0-103996273298%28Office.15%29.aspx) that you can call from content and task pane add-ins. See [Understanding the JavaScript API for Office](../../docs/overview/understanding-the-javascript-api-for-office.md) for an overview of the features of the API. For samples, see [Office Add-ins code samples](../samples/code-samples.md).

You can categorize the primary objects and methods supported by content and task pane add-ins as follows:


1.  **Common objects shared with other Office Add-ins**
    
    These objects include [Office](http://msdn.microsoft.com/en-us/library/c490b13d-ee52-4291-af5d-f4a5a11d3af0%28Office.15%29.aspx), [Context](http://msdn.microsoft.com/library/662883d5-b86f-4bdc-99f0-9ee9129ed16c%28Office.15%29.aspx), and [AsyncResult](http://msdn.microsoft.com/library/540c114f-0398-425c-baf3-7363f2f6bc47%28Office.15%29.aspx). The  **Office** object is the root object of the JavaScript API for Office. The **Context** object represents the add-in's runtime environment. Both **Office** and **Context** are the fundamental objects for any Office Add-in. The **AsyncResult** object represents the results of an asynchronous operation, such as the data returned to the **getSelectedDataAsync** method, which reads what a user has selected in a document.
    
2.  **The Document object**
    
    The majority of the API available to content and task pane add-ins is exposed through the methods, properties, and events of the [Document](http://msdn.microsoft.com/en-us/library/f8859516-cc1f-4b20-a8f3-cee37a983e70%28Office.15%29.aspx) object. Using this subset of the API, your content or task pane add-in can perform the tasks described later in this topic.
    
    A content or task pane add-in can use the [Office.context.document](http://msdn.microsoft.com/library/92351713-9ea0-43e1-b549-dad93b3208b2%28Office.15%29.aspx) property to access the **Document** object, and through it, can access the key members of the API for working with data in documents, such as the [Bindings](http://msdn.microsoft.com/library/09979e31-3bfb-45be-adda-0f7cc2db1fe1%28Office.15%29.aspx) and [CustomXmlParts](http://msdn.microsoft.com/library/ba40cd4c-29bb-4f31-875d-6f1382fd1ee8%28Office.15%29.aspx) objects, and the [getSelectedDataAsync](http://msdn.microsoft.com/library/f85ad02c-64f0-4b73-87f6-7f521b3afd69%28Office.15%29.aspx), [setSelectedDataAsync](http://msdn.microsoft.com/library/998f38dc-83bd-4659-a759-4758c632a6ef%28Office.15%29.aspx), and [getFileAsync](http://msdn.microsoft.com/library/78047418-89c4-4c7d-9427-4735b8559518%28Office.15%29.aspx) methods. The **Document** object also provides the [mode](http://msdn.microsoft.com/library/551369c3-315b-428f-8b7e-08987f6b0e00%28Office.15%29.aspx) property for determining whether a document is read-only or in edit mode, the [url](http://msdn.microsoft.com/library/480ac3c6-370e-4505-aba3-1d0dce9fb3dc%28Office.15%29.aspx) property to get the URL of the current document, and access to the [Settings](http://msdn.microsoft.com/library/ad733387-a58c-4514-8fc2-53e64fad468d%28Office.15%29.aspx) object. The **Document** object also supports adding event handlers for the [SelectionChanged](http://msdn.microsoft.com/library/4cbc527c-a1d5-4fb0-b6db-28cc40c5d5e2%28Office.15%29.aspx) event, so you can detect when a user changes his or her selection in the document.
    
    A content or task pane add-in can access the  **Document** object only after the DOM and runtime environment has been loaded, typically in the event handler for the [Office.initialize](http://msdn.microsoft.com/library/727adf79-a0b5-48d2-99c7-6642c2c334fc%28Office.15%29.aspx) event. For information about the flow of events when an add-in is initialized, and how to check that the DOM and runtime and loaded successfully, see [Loading the DOM and runtime environment](../../docs/overview/loading-the-dom-and-runtime-environment.md).
    
3.  **Objects for working with specific features**
    
    To work with specific features of the API, your content or task pane add-in can work with the following objects and methods:
    
      - Use the methods of the [Bindings](http://msdn.microsoft.com/library/09979e31-3bfb-45be-adda-0f7cc2db1fe1%28Office.15%29.aspx) object to create or get bindings, and then work with their data by using the methods and properties of the [Binding](http://msdn.microsoft.com/library/42882642-d22b-47d2-a8d3-3aa8c6a4435e%28Office.15%29.aspx) object.
    
  - Use the [CustomXmlParts](http://msdn.microsoft.com/library/ba40cd4c-29bb-4f31-875d-6f1382fd1ee8%28Office.15%29.aspx), [CustomXmlPart](http://msdn.microsoft.com/library/83f0e668-8236-4f2f-a20f-b173a9e3f65f%28Office.15%29.aspx) and associated objects to create and manipulate custom XML parts in Word documents.
    
  - Use the [File](http://msdn.microsoft.com/library/04923ddf-8efa-459f-aed5-d8c06385ca50%28Office.15%29.aspx) and [Slice](http://msdn.microsoft.com/library/011b5647-639b-4b06-8625-ba9de01bed4b%28Office.15%29.aspx) objects to create a copy of the entire document, break it into chunks or "slices", and then read or transmit the data in those slices.
    
  - Use the [Settings](http://msdn.microsoft.com/library/ad733387-a58c-4514-8fc2-53e64fad468d%28Office.15%29.aspx) object to save custom data, such as user preferences, and add-in state.
    

 >**Important**  Some of the API members aren't supported across all Office applications that can host content and task pane add-ins. To determine which members are supported, see any of the following:

For a summary of the JavaScript API for Office support available across Office host applications, see the [API support matrix](../../docs/overview/understanding-the-javascript-api-for-office.md#APIOverview_APISupportMatrix) in [Understanding the JavaScript API for Office](../../docs/overview/understanding-the-javascript-api-for-office.md).


## Reading and writing to an active selection


You can read or write to the user's current selection in a document, spreadsheet, or presentation. Depending on the host application for your add-in, you can specify the type of data structure to read or write as a parameter in the [getSelectedDataAsync](http://msdn.microsoft.com/library/f85ad02c-64f0-4b73-87f6-7f521b3afd69%28Office.15%29.aspx) and [setSelectedDataAsync](http://msdn.microsoft.com/library/998f38dc-83bd-4659-a759-4758c632a6ef%28Office.15%29.aspx) methods of the [Document](http://msdn.microsoft.com/library/f8859516-cc1f-4b20-a8f3-cee37a983e70%28Office.15%29.aspx) object. For example, you can specify any type of data (text, HTML, tabular data, or Office Open XML) for Word, text and tabular data for Excel, and text for PowerPoint and Project. You can also create event handlers to detect changes to the user's selection. The following example gets data from the selection as text using the **getSelectedDataAsync** method.


```js
Office.context.document.getSelectedDataAsync(
    Office.CoercionType.Text, function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
            write('Action failed. Error: ' + asyncResult.error.message);
        }
        else {
            write('Selected data: ' + asyncResult.value);
        }
    });

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}

```

For more details and examples, see [Read and write data to the active selection in a document or spreadsheet](../../docs/how-to/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md).


## Binding to a region in a document or spreadsheet


As described in the previous section, you can use the  **getSelectedDataAsync** and **setSelectedDataAsync** methods to read or write to the user's _current_ selection in a document, spreadsheet, or presentation. However, if you would like to access the same region in a document across sessions of running your add-in without requiring the user to make a selection, you should first bind to that region. You can also subscribe to data and selection change events for that bound region.

You can add a binding by using [addFromNamedItemAsync](http://msdn.microsoft.com/library/afbadac7-60c7-47cb-9477-6e9466ded44c%28Office.15%29.aspx), [addFromPromptAsync](http://msdn.microsoft.com/library/9dc03608-b08b-4700-8be1-3c86ae236799%28Office.15%29.aspx), or [addFromSelectionAsync](http://msdn.microsoft.com/library/edc99214-e63e-43f2-9392-97ead42fc155%28Office.15%29.aspx) methods of the [Bindings](http://msdn.microsoft.com/library/09979e31-3bfb-45be-adda-0f7cc2db1fe1%28Office.15%29.aspx) object. These methods return an identifier that you can use to access data in the binding, or to subscribe to its data change or selection change events.

The following is an example that adds a binding to the currently selected text in a document, by using the  **Bindings.addFromSelectionAsync** method.




```js
Office.context.document.bindings.addFromSelectionAsync(
    Office.BindingType.Text, { id: 'myBinding' }, function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } else {
        write('Added new binding with type: ' +
            asyncResult.value.type + ' and id: ' + asyncResult.value.id);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

For more details and examples, see [Bind to regions in a document or spreadsheet](../../docs/how-to/bind-to-regions-in-a-document-or-spreadsheet.md).


## Getting entire documents


If your task pane add-in runs in PowerPoint or Word, you can use the [Document.getFileAsync](http://msdn.microsoft.com/en-us/library/78047418-89c4-4c7d-9427-4735b8559518%28Office.15%29.aspx), [File.getSliceAsync](http://msdn.microsoft.com/en-us/library/5a8a5cc2-e883-42cd-92ab-d63e10c4c707%28Office.15%29.aspx), and [File.closeAsync](http://msdn.microsoft.com/en-us/library/1ad5cebf-6feb-43ff-8b19-97d91132ab2b%28Office.15%29.aspx) methods to get an entire presentation or document.

When you call  **Document.getFileAsync**, you get a copy of the document in a [File](http://msdn.microsoft.com/en-us/library/04923ddf-8efa-459f-aed5-d8c06385ca50%28Office.15%29.aspx) object. The **File** object provides access to the document in "chunks" represented as [Slice](http://msdn.microsoft.com/en-us/library/011b5647-639b-4b06-8625-ba9de01bed4b%28Office.15%29.aspx) objects. When you call **getFileAsync**, you can specify the file type (text or compressed Open Office XML format), and size of the slices (up to 4MB). To access the contents of the  **File** object, you then call **File.getSliceAsync** which returns the raw data in the [Slice.data](http://msdn.microsoft.com/en-us/library/95a68949-6009-49ae-a531-2df77687b85d%28Office.15%29.aspx) property. If you specified compressed format, you will get the file data as a byte array. If you are transmitting the file to a web service, you can transform the compressed raw data to a base64-encoded string before submission. Finally, when you are finished getting slices of the file, use the **File.closeAsync** method to close the document.

For more details, see how to [get the whole document from an add-in for PowerPoint or Word](http://msdn.microsoft.com/en-us/library/47a4ab14-0f1e-4cc8-8814-fa7e97362360%28Office.15%29.aspx). 


## Reading and writing custom XML parts of a Word document


Using the Open Office XML file format and content controls, you can add custom XML parts to a Word document and bind elements in the XML parts to content controls in that document. When you open the document, Word reads and automatically populates bound content controls with data from the custom XML parts. Users can also write data into the content controls, and when the user saves the document, the data in the controls will be saved to the bound XML parts. Task pane add-ins for Word, can use the [Document.customXmlParts](http://msdn.microsoft.com/en-us/library/b72c08bc-b49c-497c-9521-26ccce148bda%28Office.15%29.aspx) property, [CustomXmlParts](http://msdn.microsoft.com/en-us/library/ba40cd4c-29bb-4f31-875d-6f1382fd1ee8%28Office.15%29.aspx), [CustomXmlPart](http://msdn.microsoft.com/en-us/library/83f0e668-8236-4f2f-a20f-b173a9e3f65f%28Office.15%29.aspx), and [CustomXmlNode](http://msdn.microsoft.com/en-us/library/dc1518de-47fa-4108-aab7-04a022724b04%28Office.15%29.aspx) objects to read and write data dynamically to the document.

Custom XML parts may be associated with namespaces. To get data from custom XML parts in a namespace, use the [CustomXmlParts.getByNamespaceAsync](http://msdn.microsoft.com/en-us/library/9902f555-5c20-45d6-9a8c-ae6bf013dfaf%28Office.15%29.aspx) method.

You can also use the [CustomXmlParts.getByIdAsync](http://msdn.microsoft.com/en-us/library/31a21b58-426e-4bbe-acdf-885b32ce50ab%28Office.15%29.aspx) method to access custom XML parts by their GUIDs. After getting a custom XML part, use the [CustomXmlPart.getXmlAsync](http://msdn.microsoft.com/en-us/library/6606365a-9244-49b5-9393-fe2186091af7%28Office.15%29.aspx) method to get the XML data.

To add a new custom XML part to a document, use the  **Document.customXmlParts** property to get the custom XML parts that are in the document, and call the [CustomXmlParts.addAsync](http://msdn.microsoft.com/en-us/library/2816397c-b86a-4f52-8b13-036f527f4bb7%28Office.15%29.aspx) method.

For detailed information about how to work with custom XML parts with a task pane add-in, see [Creating Better Add-ins for Word with Office Open XML](http://msdn.microsoft.com/en-us/library/c5bad651-a42f-4e57-bc60-c9b27eb2383b%28Office.15%29.aspx).


## Persisting add-in settings


Often you need to save custom data for your add-in, such as a user's preferences or the add-in's state, and access that data the next time the add-in is opened. You can use common web programming techniques to save that data, such as browser cookies or HTML 5 web storage. Alternatively, if your add-in runs in Excel, PowerPoint, or Word, you can use the methods of the [Document.Settings](http://msdn.microsoft.com/en-us/library/ad733387-a58c-4514-8fc2-53e64fad468d%28Office.15%29.aspx) object. Data created with the **Settings** object is stored in the spreadsheet, presentation, or document that the add-in was inserted into and saved with. This data is available to only the add-in that created it.

To avoid roundtrips to the server where the document is stored, data created with the  **Settings** object is managed in memory at run time. Previously saved settings data is loaded into memory when the add-in is initialized, and changes to that data are only saved back to the document when you call the [Settings.saveAsync](http://msdn.microsoft.com/en-us/library/7147c221-937c-477c-98a6-f59d6200c27b%28Office.15%29.aspx) method. Internally, the data is stored in a serialized JSON object as name/value pairs. You use the [get](http://msdn.microsoft.com/en-us/library/aeac06dd-994e-4235-b208-1bd117395296%28Office.15%29.aspx), [set](http://msdn.microsoft.com/en-us/library/4e2c9758-953e-41e8-aca6-d8daf764a584%28Office.15%29.aspx), and [remove](http://msdn.microsoft.com/en-us/library/a92446bf-de65-45bd-8412-36ea8e77c5a2%28Office.15%29.aspx) methods of the **Settings** object, to read, write, and delete items from the in-memory copy of the data. The following line of code shows how to create a setting named `themeColor` and set its value to 'green'.




```
Office.context.document.settings.set('themeColor', 'green');
```

Because settings data created or deleted with the  **set** and **remove** methods is acting on an in-memory copy of the data, you must call **saveAsync** to persist changes to settings data into the document your add-in is working with.

For more details about working with custom data using the methods of the  **Settings** object, see [Persisting add-in state and settings](../../docs/how-to/persisting-add-in-state-and-settings.md).


## Reading properties of a project document


If your task pane add-in runs in Project, your add-in can read data from some of the project fields, resource, and task fields in the active project. To do that, you use the methods and events of the [ProjectDocument](http://msdn.microsoft.com/en-us/library/1908af4f-93b9-4859-87e3-06942014fae1%28Office.15%29.aspx) object, which extends the **Document** object to provide additional Project-specific functionality.

For examples of reading Project data, see [Create your first task pane add-in for Project 2013 by using a text editor](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md).


## Permissions model and governance


Your add-in uses the  **Permissions** element in its manifest to request permission to access the level of functionality it requires from the JavaScript API for Office. For example, if your add-in requires read/write access to the document, its manifest must specify `ReadWriteDocument` as the text value in its **Permissions** element. Because permissions exist to protect a user's privacy and security, as a best practice you should request the minimum level of permissions it needs for its features. The following example shows how to request the **ReadDocument** permission in a task pane's manifest.


```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.0"
 xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
 xsi:type="TaskPaneApp">
???????<!-- Other manifest elements omitted. -->
  <Permissions>ReadDocument</Permissions>
???????
</OfficeApp>

```

For more information, see [Requesting permissions for API use in content and task pane add-ins](../../docs/how-to/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md).


## Additional resources



- [JavaScript API for Office](http://msdn.microsoft.com/library/b27e70c3-d87d-4d27-85e0-103996273298%28Office.15%29.aspx)
    
- [Schema reference for Office Add-ins manifests](http://msdn.microsoft.com/en-us/library/7e0cadc3-f613-8eb9-57ef-9032cbb97f92.aspx)
    
- [Troubleshoot user errors with Office Add-ins](../testing/testing-and-troubleshooting.md)
    
