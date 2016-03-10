
# Bind to regions in a document or spreadsheet
This article describes how to create bindings to regions of documents and spreadsheets, and then read and write data to those bindings. It also describes how to create and remove event handlers for changes to data or the user's selection in a binding. 



## Binding to regions in a document or spreadsheet

Binding-based data access enables content and task pane add-ins to consistently access a particular region of a document or spreadsheet through an identifier. The add-in first needs to establish the binding by calling one of the methods that associates a portion of the document with a unique identifier: [addFromPromptAsync](http://msdn.microsoft.com/en-us/library/9dc03608-b08b-4700-8be1-3c86ae236799%28Office.15%29.aspx), [addFromSelectionAsync](http://msdn.microsoft.com/en-us/library/edc99214-e63e-43f2-9392-97ead42fc155%28Office.15%29.aspx), or [addFromNamedItemAsync](http://msdn.microsoft.com/en-us/library/afbadac7-60c7-47cb-9477-6e9466ded44c%28Office.15%29.aspx). After the binding is established, the add-in can use the provided identifier to access the data contained in the associated region of the document or spreadsheet. Creating bindings provides the following value to your add-in:


- Permits access to common data structures across supported Office applications, such as: tables, ranges, or text (a contiguous run of characters).
    
- Enables read/write operations without requiring the user to make a selection.
    
- Establishes a relationship between the add-in and the data in the document. Bindings are persisted in the document, and can be accessed at a later time.
    
Establishing a binding also allows you to subscribe to data and selection change events that are scoped to that particular region of the document or spreadsheet. This means that the add-in is only notified of changes that happen within the bound region as opposed to general changes across the whole document or spreadsheet.

The [Bindings](http://msdn.microsoft.com/en-us/library/09979e31-3bfb-45be-adda-0f7cc2db1fe1%28Office.15%29.aspx) object exposes a [getAllAsync](http://msdn.microsoft.com/en-us/library/ef902b73-cc4c-4551-95de-d8a51eeba82f%28Office.15%29.aspx) method that gives access to the set of all bindings established on the document or spreadsheet. An individual binding can be accessed by its ID using either the [Bindings.getBindingByIdAsync](http://msdn.microsoft.com/en-us/library/2727c891-bc05-465c-9324-113fbfeb3fbb%28Office.15%29.aspx) or [Office.select](http://msdn.microsoft.com/en-us/library/23aeb136-da1f-4127-a798-99dc27bc4dae%28Office.15%29.aspx) methods. You can establish new bindings as well as remove existing ones by using one of the following methods of the **Bindings** object: [addFromSelectionAsync](http://msdn.microsoft.com/en-us/library/edc99214-e63e-43f2-9392-97ead42fc155%28Office.15%29.aspx), [addFromPromptAsync](http://msdn.microsoft.com/en-us/library/9dc03608-b08b-4700-8be1-3c86ae236799%28Office.15%29.aspx), [addFromNamedItemAsync](http://msdn.microsoft.com/en-us/library/afbadac7-60c7-47cb-9477-6e9466ded44c%28Office.15%29.aspx), or [releaseByIdAsync](http://msdn.microsoft.com/en-us/library/ad285984-8b44-435d-9b84-f0ade570c896%28Office.15%29.aspx).

There are three different types of bindings that you specify with the  _bindingType_ parameter when you create a binding with the **addFromSelectionAsync**, **addFromPromptAsync** or **addFromNamedItemAsync** methods:



|**Binding type**|**Description**|**Host application support**|
|:-----|:-----|:-----|
|Text binding|Binds to a region of the document that can be represented as text.|In Word, most contiguous selections are valid, while in Excel only single cell selections can be the target of a text binding. In Excel, only plain text is supported. In Word, three formats are supported: plain text, HTML, and Open XML for Office.|
|Matrix binding|Binds to a fixed region of a document that contains tabular data without headers.Data in a matrix binding is written or read as a two dimensional  **Array**, which in JavaScript is implemented as an array of arrays. For example, two rows of  **string** values in two columns can be written or read as ` [['a', 'b'], ['c', 'd']]`, and a single column of three rows can be written or read as  `[['a'], ['b'], ['c']]`.|In Excel, any contiguous selection of cells can be used to establish a matrix binding. In Word, only tables support matrix binding.|
|Table binding|Binds to a region of a document that contains a table with headers.Data in a table binding is written or read as a [TableData](http://msdn.microsoft.com/en-us/library/2183ea52-5a40-4048-b9a4-7cd66bb0ad5d%28Office.15%29.aspx) object. The **TableData** object exposes the data through the **headers** and **rows** properties.|Any Excel or Word table can be the basis for a table binding. After you establish a table binding, each new row or column a user adds to the table is automatically included in the binding. |
After a binding is created by using one of the three "addFrom" methods of the  **Bindings** object, you can work with the binding's data and properties by using the methods of the corresponding object: [MatrixBinding](http://msdn.microsoft.com/en-us/library/35e8568e-9129-4c00-b30f-d8c3b2555f1e%28Office.15%29.aspx), [TableBinding](http://msdn.microsoft.com/en-us/library/1508795b-1c70-456c-b3bf-666d40cf8f50%28Office.15%29.aspx), or [TextBinding](http://msdn.microsoft.com/en-us/library/6b71b21d-f64d-425c-99d9-c62b2a9969be%28Office.15%29.aspx). All three of these objects inherit the [getDataAsync](http://msdn.microsoft.com/en-us/library/5372ffd8-579d-4fcb-9e5b-e9a2128f3201%28Office.15%29.aspx) and [setDataAsync](http://msdn.microsoft.com/en-us/library/6a59bb6d-40b6-4a95-9b98-d70d4616de09%28Office.15%29.aspx) methods of the **Binding** object that enable you to interact with the bound data.


 >**Tip**   **When should you use matrix versus table bindings?** **Note:** When the tabular data you are working with contains a total row, you must use a matrix binding if your add-in's script needs to access values in the total row or detect that the user's selection is in the total row. If you establish a table binding for tabular data that contains a total row, the [TableBinding.rowCount](http://msdn.microsoft.com/en-us/library/128d0e04-8ec7-4f52-a31d-421d225c39fa%28Office.15%29.aspx) property and the [rowCount](http://msdn.microsoft.com/en-us/library/110d45f7-40b7-4005-b080-ef748cbf337c%28Office.15%29.aspx) and [startRow](http://msdn.microsoft.com/en-us/library/3cc1c014-b18d-4e7b-9ec0-5500b43c4016%28Office.15%29.aspx) properties of the **BindingSelectionChangedEventArgs** object in event handers won't reflect the total row in their values. To work around this limitation, you must use establish a matrix binding to work with the total row.


### Add a binding to the user's current selection


The following example shows how to add a text binding called  `myBinding` to the current selection in a document by using the [Bindings.addFromSelectionAsync](http://msdn.microsoft.com/en-us/library/edc99214-e63e-43f2-9392-97ead42fc155%28Office.15%29.aspx) method.


```js
Office.context.document.bindings.addFromSelectionAsync(Office.BindingType.Text, { id: 'myBinding' }, function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } else {
        write('Added new binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

In this example, the specified binding type is text. This means that a [TextBinding](http://msdn.microsoft.com/en-us/library/6b71b21d-f64d-425c-99d9-c62b2a9969be%28Office.15%29.aspx) will be created for the selection. Different binding types expose different data and operations. [Office.BindingType](http://msdn.microsoft.com/en-us/library/e5591c38-806a-423d-b9d1-3041c726d524%28Office.15%29.aspx) is an enumeration of available binding type values.

The second optional parameter is an object that specifies the ID of the new binding being created. If an ID is not specified, one is generated automatically.

The anonymous function that is passed into the function as the final  _callback_ parameter is executed when the creation of the binding is complete. The function is called with a single parameter, _asyncResult_, which provides access to an [AsyncResult](http://msdn.microsoft.com/en-us/library/540c114f-0398-425c-baf3-7363f2f6bc47%28Office.15%29.aspx) object that provides the status of the call. The [AsyncResult.value](http://msdn.microsoft.com/en-us/library/453a4b43-0fdc-4ea9-967a-c033fab31507%28Office.15%29.aspx) property contains a reference to a [Binding](http://msdn.microsoft.com/en-us/library/42882642-d22b-47d2-a8d3-3aa8c6a4435e%28Office.15%29.aspx) object of the type that is specified for the newly created binding. You can use this **Binding** object to get and set data.


### Add a binding from a prompt


The following example shows how to add a text binding called  `myBinding` by using the [Bindings.addFromPromptAsync](http://msdn.microsoft.com/en-us/library/9dc03608-b08b-4700-8be1-3c86ae236799%28Office.15%29.aspx) method, which is only supported in Excel 2013 and Excel Online. This method lets the user specify the range for the binding by using the application's built-in range selection prompt.


```js
function bindFromPrompt() {
    Office.context.document.bindings.addFromPromptAsync(Office.BindingType.Text, { id: 'myBinding' }, function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
            write('Action failed. Error: ' + asyncResult.error.message);
        } else {
            write('Added new binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
        }
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

In this example, the specified binding type is text. This means that a [TextBinding](http://msdn.microsoft.com/en-us/library/6b71b21d-f64d-425c-99d9-c62b2a9969be%28Office.15%29.aspx) will be created for the selection that the user specifies in the prompt.

The second parameter is an object that contains the ID of the new binding being created. If an ID is not specified, one is generated automatically.

The anonymous function passed into the function as the third  _callback_ parameter is executed when the creation of the binding is complete. When the callback function executes, the [AsyncResult](http://msdn.microsoft.com/en-us/library/540c114f-0398-425c-baf3-7363f2f6bc47%28Office.15%29.aspx) object contains the status of the call and the newly created binding.

Figure 1 shows the built-in range selection prompt in Excel.


**Figure 1. Excel Select Data UI**

![Excel Select Data UI](../../images/AgaveAPIOverview_ExcelSelectionUI.png)


### Add a binding to a named item


The following example shows how to add a binding to the existing  `myRange` named item as a "matrix" binding by using the [Bindings.addFromNamedItemAsync](http://msdn.microsoft.com/en-us/library/afbadac7-60c7-47cb-9477-6e9466ded44c%28Office.15%29.aspx) method, and assigns the binding's **id** as "myMatrix".


```js
function bindNamedItem() {
    Office.context.document.bindings.addFromNamedItemAsync("myRange", "matrix", {id:'myMatrix'}, function (result) {
        if (result.status == 'succeeded'){
            write('Added new binding with type: ' + result.value.type + ' and id: ' + result.value.id);
            }
        else
            write('Error: ' + result.error.message);
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}

```

 **For Excel**, the  _itemName_ parameter of the **addFromNamedItemAsync** method can refer to an existing named range, a range specified with the A1 reference style ("A1:A3"), or a table. By default, adding a table in Excel assigns the name "Table1" for the first table you add, "Table2" for the second table you add, and so on. To assign a meaningful name for a table in the Excel UI, use the **Table Name** property on the **Table Tools | Design** tab of the ribbon.


 >**Note**  In Excel 2013, when specifying a table as a named item, you must fully qualify the name to include the worksheet name in the name of the table in this format:  `"Sheet1!Table1"`

The following example creates a binding in Excel to the first three cells in column A ( `"A1:A3"`), assigns the  **id** `"MyCities"`, and then writes three city names to that binding.




```js
 function bindingFromA1Range() {
    Office.context.document.bindings.addFromNamedItemAsync("A1:A3", "matrix", {id: "MyCities" },
        function (asyncResult) {
            if (asyncResult.status == "failed") {
                write('Error: ' + asyncResult.error.message);
            }
            else {
                // Write data to the new binding.
                Office.select("bindings#MyCities").setDataAsync([['Berlin'], ['Munich'], ['Duisburg']], { coercionType: "matrix" },
                    function (asyncResult) {
                        if (asyncResult.status == "failed") {
                            write('Error: ' + asyncResult.error.message);
                        }
                    });
            }
        });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

 **For Word**, the  _itemName_ parameter of the **addFromNamedItemAsync** method refers to the **Title** property of a **Rich Text** content control. (You can't bind to content controls other than the **Rich Text** content control.)

By default, a content control has no  **Title** value assigned. To assign a meaningful name in the Word UI, after inserting a **Rich Text** content control from the **Controls** group on the **Developer** tab of the ribbon, use the **Properties** command in the **Controls** group to display the **Content Control Properties** dialog box. Then set the **Title** property of the content control to the name you want to reference from your code.

The following example creates a text binding in Word to a rich text content control named  `"FirstName"`, assigns the  **id** `"firstName"`, and then displays that information.




```js
function bindContentControl() {
    Office.context.document.bindings.addFromNamedItemAsync('FirstName', 
        Office.BindingType.Text, {id:'firstName'},
        function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                write('Control bound. Binding.id: '
                    + result.value.id + ' Binding.type: ' + result.value.type);
            } else {
                write('Error:', result.error.message);
            }
    });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


### Get all bindings


The following example shows how to get all bindings in a document by using the [Bindings.getAllAsync](http://msdn.microsoft.com/en-us/library/ef902b73-cc4c-4551-95de-d8a51eeba82f%28Office.15%29.aspx) method.


```js
Office.context.document.bindings.getAllAsync(function (asyncResult) {
    var bindingString = '';
    for (var i in asyncResult.value) {
        bindingString += asyncResult.value[i].id + '\n';
    }
    write('Existing bindings: ' + bindingString);
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

The anonymous function that is passed into the function as the  _callback_ parameter is executed when the operation is complete. The function is called with a single parameter, _asyncResult_, which contains an  **array** of the bindings in the document. The array is iterated to build a string that contains the IDs of the bindings. The string is then displayed in a message box.


### Get a binding by ID using the getByIdAsync method of the Bindings object


The following example shows how to use the [Bindings.getByIdAsync](http://msdn.microsoft.com/en-us/library/2727c891-bc05-465c-9324-113fbfeb3fbb%28Office.15%29.aspx) method to get a binding in a document by specifying its ID. This example assumes that a binding named `'myBinding'` was added to the document using one of the methods described earlier in this topic.


```js
Office.context.document.bindings.getByIdAsync('myBinding', function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } 
    else {
        write('Retrieved binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

In the example, the first  _id_ parameter is the ID of the binding to retrieve.

The anonymous function that is passed into the function as the second  _callback_ parameter is executed when the operation is completed. The function is called with a single parameter, _asyncResult_, which contains the status of the call and the binding with the ID "myBinding".


### Get a binding by ID using the select method of the Office object


The following example shows how to use the [Office.select](http://msdn.microsoft.com/en-us/library/23aeb136-da1f-4127-a798-99dc27bc4dae%28Office.15%29.aspx) method to get a **Binding** object promise in a document by specifying its ID in a selector string. It then calls the [Binding.getDataAsync](http://msdn.microsoft.com/en-us/library/5372ffd8-579d-4fcb-9e5b-e9a2128f3201%28Office.15%29.aspx) method to get data from the specified binding. This example assumes that a binding named `'myBinding'` was added to the document using one of the methods described earlier in this topic.


```js
Office.select("bindings#myBinding", function onError(){}).getDataAsync(function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } else {
        write(asyncResult.value);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


 >**Note**  If the  **select** method promise successfully returns a **Binding** object, that object exposes only the following four methods of the [Binding](http://msdn.microsoft.com/en-us/library/42882642-d22b-47d2-a8d3-3aa8c6a4435e%28Office.15%29.aspx) object: [getDataAsync](http://msdn.microsoft.com/en-us/library/5372ffd8-579d-4fcb-9e5b-e9a2128f3201%28Office.15%29.aspx), [setDataAsync](http://msdn.microsoft.com/en-us/library/6a59bb6d-40b6-4a95-9b98-d70d4616de09%28Office.15%29.aspx), [addHandlerAsync](http://msdn.microsoft.com/en-us/library/b9c2f4ea-726c-4b48-a3fb-89beda337a17%28Office.15%29.aspx), and [removeHandlerAsync](http://msdn.microsoft.com/en-us/library/5ae3a860-1fc4-46ce-858e-98545c3e2d77%28Office.15%29.aspx). If the promise cannot return a  **Binding** object, the _onError_ callback can be used to access an [asyncResult.error](http://msdn.microsoft.com/en-us/library/51c46d36-972d-4d82-91aa-da99cbeb8d4f%28Office.15%29.aspx) object to get more information.If you need to call a member of the  **Binding** object other than the four methods exposed by the **Binding** object promise returned by the **select** method, instead use the [getByIdAsync](http://msdn.microsoft.com/en-us/library/2727c891-bc05-465c-9324-113fbfeb3fbb%28Office.15%29.aspx) method by using the [Document.bindings](http://msdn.microsoft.com/en-us/library/6512eabc-a177-42da-bc52-99665817515f%28Office.15%29.aspx) property and [Bindings.getByIdAsync](http://msdn.microsoft.com/en-us/library/2727c891-bc05-465c-9324-113fbfeb3fbb%28Office.15%29.aspx) method to retrieve the **Binding** object.


### Release a binding by ID


The following example shows how use the [Bindings.releaseByIdAsync](http://msdn.microsoft.com/en-us/library/ad285984-8b44-435d-9b84-f0ade570c896%28Office.15%29.aspx) method to release a binding in a document by specifying its ID.


```js
Office.context.document.bindings.releaseByIdAsync('myBinding', function (asyncResult) {
    write('Released myBinding!');
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

In the example, the first  _id_ parameter is the ID of the binding to release.

The anonymous function that is passed into the function as the second parameter is a callback that is executed when the operation is complete. The function is called with a single parameter,  _asyncResult_, which contains the status of the call.


### Read data from a binding


The following example shows how to use the [Binding.getDataAsync](http://msdn.microsoft.com/en-us/library/5372ffd8-579d-4fcb-9e5b-e9a2128f3201%28Office.15%29.aspx) method to get data from an existing binding.


```js
myBinding.getDataAsync(function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } else {
        write(asyncResult.value);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

 `myBinding` is a variable that contains an existing text binding in the document. Alternatively, you could use the [Office.select method](../../docs/how-to/bind-to-regions-in-a-document-or-spreadsheet.md#BindRegions_Select) to access the binding by its ID, and start your call to the **getDataAsync** method, like this: `Office.select("bindings#myBindingID").getDataAsync`.

The anonymous function that is passed into the function is a callback that is executed when the operation is complete. The [AsyncResult.value](http://msdn.microsoft.com/en-us/library/453a4b43-0fdc-4ea9-967a-c033fab31507%28Office.15%29.aspx) property contains the data within `myBinding`. The type of the value depends on the binding type. The binding in this example is a text binding. Therefore, the value will contain a string. For additional examples of working with matrix and table bindings, see the [Binding.getDataAsync](http://msdn.microsoft.com/en-us/library/5372ffd8-579d-4fcb-9e5b-e9a2128f3201%28Office.15%29.aspx) method topic.


### Write data to a binding


The following example shows how to use the [Binding.setDataAsync](http://msdn.microsoft.com/en-us/library/6a59bb6d-40b6-4a95-9b98-d70d4616de09%28Office.15%29.aspx) method to set data in an existing binding.


```js
myBinding.setDataAsync('Hello World!', function (asyncResult) { });
```

 `myBinding` is a variable that contains an existing text binding in the document.

In the example, the first parameter is the value to set on  `myBinding`. Because this is a text binding, the value is a  **string**. Different binding types accept different types of data.

The anonymous function that is passed into the function is a callback that is executed when the operation is complete. The function is called with a single parameter,  _asyncResult_, which contains the status of the result.

 **Note:** Starting with the release of the Excel 2013 SP1 and the corresponding build of Excel Online, you can now [set formatting when writing and updating data in bound tables](../../docs/how-to/format-tables-in-add-ins-for-excel.md).


### Detect changes to data or the selection in a binding


The following example shows how to attach an event handler to the [DataChanged](http://msdn.microsoft.com/en-us/library/7b9ed4bf-3ce5-44eb-8548-2b081afd868d%28Office.15%29.aspx) event of a binding with an id of "MyBinding".


```js
function addHandler() {
Office.select("bindings#MyBinding").addHandlerAsync(
    Office.EventType.BindingDataChanged, dataChanged);
}
function dataChanged(eventArgs) {
    write('Bound data changed in binding: ' + eventArgs.binding.id);
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

 `myBinding` is a variable that contains an existing text binding in the document.

The first  _eventType_ parameter of the [binding.addHandlerAsync](http://msdn.microsoft.com/en-us/library/b9c2f4ea-726c-4b48-a3fb-89beda337a17%28Office.15%29.aspx) method specifies the name of the event to subscribe to. [Office.EventType](http://msdn.microsoft.com/en-us/library/82c79659-52da-48b0-92a9-831226eb9a7f%28Office.15%29.aspx) is an enumeration of available event type values. **Office.EventType.BindingDataChanged** evaluates to the string `"bindingDataChanged"`.

The  `dataChanged` function that is passed into the function as the second _handler_ parameter is an event handler that is executed when the data in the binding is changed. The function is called with a single parameter, _eventArgs_, which contains a reference to the binding. This binding can be used to retrieve the updated data.

Similarly, you can detect when a user changes selection in a binding by attaching an event handler to the [SelectionChanged](http://msdn.microsoft.com/en-us/library/5bcbb5e2-f8e6-48ee-bde0-60d12d43ff5f%28Office.15%29.aspx) event of a binding. To do that, specify the _eventType_ parameter of the **binding.addHandlerAsync** method as **Office.EventType.BindingSelectionChanged** or `"bindingSelectionChanged"`.

You can add multiple event handlers for a given event by calling the  **addHandlerAsync** method again and passing in an additional event handler function for the _handler_ parameter. This will work correctly as long as the name of each event handler function is unique.


### Remove an event handler


To remove an event handler for an event, call the [Binding.removeHandlerAsync](http://msdn.microsoft.com/en-us/library/5ae3a860-1fc4-46ce-858e-98545c3e2d77%28Office.15%29.aspx) method passing in the event type as the first _eventType_ parameter, and the name of the event handler function to remove as the second _handler_ parameter. For example, the following function will remove the `dataChanged` event handler function added in the previous section's example.


```
function removeEventHandlerFromBinding() {
    Office.select("bindings#MyBinding").removeHandlerAsync(
        Office.EventType.BindingDataChanged, {handler:dataChanged});
}
```


 >**Important**  If the optional  _handler_ parameter is omitted when the **removeHandlerAsync** method is called, all event handlers for the specified _eventType_ will be removed.


## Additional resources



- [Task pane and content add-ins for Office 2013](../../docs/essentials/task-pane-and-content-add-ins.md)
    
- [Understanding the JavaScript API for Office](../../docs/overview/understanding-the-javascript-api-for-office.md)
    
- [Asynchronous programming in Office Add-ins](../../docs/how-to/asynchronous-programming-in-office-add-ins.md)
    
- [Read and write data to the active selection in a document or spreadsheet](../../docs/how-to/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)
    
