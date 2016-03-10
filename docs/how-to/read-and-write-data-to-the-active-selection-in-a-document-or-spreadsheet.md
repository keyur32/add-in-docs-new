
# Read and write data to the active selection in a document or spreadsheet
This article describes how to read and write to the user's selection in a document or spreadsheet. It also describes how to create event handlers for changes in user's selection.



## Overview


The [Document](http://msdn.microsoft.com/en-us/library/f8859516-cc1f-4b20-a8f3-cee37a983e70%28Office.15%29.aspx) object exposes methods that let you to read and write to the user's current selection in a document or spreadsheet. To do that, the **Document** object provides the **getSelectedDataAsync** and **setSelectedDataAsync** methods. This topic also describes how to read, write, and create event handlers to detect changes to the user's selection.

The  **getSelectedDataAsync** method only works against the user's current selection. If you need to persist the selection in the document, so that the same selection is available to read and write to across sessions of running your add-in, you must add abinding using the [Bindings.addFromSelectionAsync](http://msdn.microsoft.com/en-us/library/edc99214-e63e-43f2-9392-97ead42fc155.aspx) method (or create a binding with one of the other "addFrom" methods of the [Bindings](http://msdn.microsoft.com/en-us/library/09979e31-3bfb-45be-adda-0f7cc2db1fe1.aspx) object). For information about creating a binding to a region of a document, and then reading and writing to a binding, see [Bind to regions in a document or spreadsheet](../../docs/how-to/bind-to-regions-in-a-document-or-spreadsheet.md).


### [Read selected data] (#ReadWriteDocumentData_Read)


The following example shows how to get data from a selection in a document by using the [getSelectedDataAsync](http://msdn.microsoft.com/en-us/library/f85ad02c-64f0-4b73-87f6-7f521b3afd69%28Office.15%29.aspx) method.


```js
Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function (asyncResult) {
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

In this example, the first  _coercionType_ parameter is specified as **Office.CoercionType.Text** (you can also specify this parameter by using the literal string `"text"`). This means that the [value](http://msdn.microsoft.com/en-us/library/453a4b43-0fdc-4ea9-967a-c033fab31507%28Office.15%29.aspx) property of the [AsyncResult](http://msdn.microsoft.com/en-us/library/540c114f-0398-425c-baf3-7363f2f6bc47%28Office.15%29.aspx) object that is available from the _asyncResult_ parameter in the callback function will return a **string** that contains the selected text in the document. Specifying different coercion types will result in different values. [Office.CoercionType](http://msdn.microsoft.com/en-us/library/735eaab6-5e31-4bc2-add5-9d378900a31b%28Office.15%29.aspx) is an enumeration of available coercion type values. **Office.CoercionType.Text** evaluates to the string "text".


 >**Tip**   **When should you use the matrix versus table coercionType for data access?** If you need your selected tabular data to grow dynamically when rows and columns are added, and you must work with table headers, you should use the table data type (by specifying the _coercionType_ parameter of the **getSelectedDataAsync** method as `"table"` or **Office.CoercionType.Table**). Adding rows and columns within the data structure is supported in both table and matrix data, but appending rows and columns is supported only for table data. If you are you aren't planning on adding rows and columns, and your data doesn't require header functionality, then you should use the matrix data type (by specifying the  _coercionType_ parameter of **getSelecteDataAsync** method as `"matrix"` or **Office.CoercionType.Matrix**), which provides a simpler model of interacting with the data.

The anonymous function that is passed into the function as the second  _callback_ parameter is executed when the **getSelectedDataAsync** operation is completed. The function is called with a single parameter, _asyncResult_, which contains the result and the status of the call. If the call fails, the [error](http://msdn.microsoft.com/en-us/library/51c46d36-972d-4d82-91aa-da99cbeb8d4f%28Office.15%29.aspx) property of the **AsyncResult** object provides access to the [Error](http://msdn.microsoft.com/en-us/library/36d1d048-b888-4bb5-9321-d340bcbc86f4%28Office.15%29.aspx) object. You can check the value of the [Error.name](http://msdn.microsoft.com/en-us/library/b76aaafd-bb34-4853-b29d-67adb1111b37%28Office.15%29.aspx) and [Error.message](http://msdn.microsoft.com/en-us/library/594e168e-4fdf-4e80-ba7e-4856a4a8ea5f%28Office.15%29.aspx) properties to determine why the set operation failed. Otherwise, the selected text in the document is displayed.

The [AsyncResult.status](http://msdn.microsoft.com/en-us/library/eec9c712-79eb-4365-88a1-6d77649727c1%28Office.15%29.aspx) property is used in the **if** statement to test whether the call succeeded. [Office.AsyncResultStatus](http://msdn.microsoft.com/en-us/library/e2652105-03e8-4771-a985-e66c661fe3ea%28Office.15%29.aspx) is an enumeration of available **AsyncResult.status** property values. **Office.AsyncResultStatus.Failed** evaluates to the string "failed" (and, again, can also be specified as that literal string).


### Write data to the selection


The following example shows how to set the selection to show "Hello World!".


```js
Office.context.document.setSelectedDataAsync("Hello World!", function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write(asyncResult.error.message);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

Passing in different object types for the  _data_ parameter will have different results. The result depends on what is currently selected in the document, which application is hosting your add-in, and whether the data passed in can be coerced to the current selection.

The anonymous function passed into the [setSelectedDataAsync](http://msdn.microsoft.com/en-us/library/998f38dc-83bd-4659-a759-4758c632a6ef%28Office.15%29.aspx) method as the _callback_ parameter is executed when the asynchronous call is completed. When you write data to the selection by using the **setSelectedDataAsync** method, the _asyncResult_ parameter of the callback provides access only to the status of the call, and to the [Error](http://msdn.microsoft.com/en-us/library/36d1d048-b888-4bb5-9321-d340bcbc86f4%28Office.15%29.aspx) object if the call fails.

 **Note:** Starting with the release of the Excel 2013 SP1 and the corresponding build of Excel Online, you can now [set formatting when writing a table to the current selection](../../docs/how-to/format-tables-in-add-ins-for-excel.md).


### Detect changes in the selection


The following example shows how to detect changes in the selection by using the [Document.addHandlerAsync](http://msdn.microsoft.com/en-us/library/8b2ec6c4-0983-4f5e-abd9-16f15b4fc87b%28Office.15%29.aspx) method to add an event handler for the [SelectionChanged](http://msdn.microsoft.com/en-us/library/4cbc527c-a1d5-4fb0-b6db-28cc40c5d5e2%28Office.15%29.aspx) event on the document.


```
Office.context.document.addHandlerAsync("documentSelectionChanged", myHandler, function(result){} 
);

// Event handler function.
function myHandler(eventArgs){
write('Document Selection Changed');
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

The first  _eventType_ parameter specifies the name of the event to subscribe to. Passing the string `"documentSelectionChanged"` for this parameter is equivalent to passing the **Office.EventType.DocumentSelectionChanged** event type of the [Office.EventType](http://msdn.microsoft.com/en-us/library/82c79659-52da-48b0-92a9-831226eb9a7f%28Office.15%29.aspx) enumeration.

The  `myHander()` function that is passed into the function as the second _handler_ parameter is an event handler that is executed when the selection is changed on the document. The function is called with a single parameter, _eventArgs_, which will contain a reference to a [DocumentSelectionChangedEventArgs](http://msdn.microsoft.com/en-us/library/283f2d97-2595-444b-86a6-286efd77f638%28Office.15%29.aspx) object when the asynchronous operation completes. You can use the [DocumentSelectionChangedEventArgs.document](http://msdn.microsoft.com/en-us/library/12974085-c146-45ff-aede-70e247d8426f%28Office.15%29.aspx) property to access the document that raised the event.


 >**Note**  You can add multiple event handlers for a given event by calling the  **addHandlerAsync** method again and passing in an additional event handler function for the _handler_ parameter. This will work correctly as long as the name of each event handler function is unique.


### Stop detecting changes in the selection


The following example shows how to stop listening to the [Document.SelectionChanged](http://msdn.microsoft.com/en-us/library/4cbc527c-a1d5-4fb0-b6db-28cc40c5d5e2%28Office.15%29.aspx) event by calling the [document.removeHandlerAsync](http://msdn.microsoft.com/en-us/library/47e0b00f-e301-4f21-836d-aeac783c42e0%28Office.15%29.aspx) method.


```
Office.context.document.removeHandlerAsync("documentSelectionChanged", {handler:myHandler}, function(result){});
```

The  `myHandler` function name that is passed as the second _handler_ parameter specifies the event handler that will be removed from the **SelectionChanged** event.


 >**Important**  If the optional  _handler_ parameter is omitted when the **removeHandlerAsync** method is called, all event handlers for the specified _eventType_ will be removed.


## Additional resources



- [Task pane and content add-ins for Office 2013](../../docs/essentials/task-pane-and-content-add-ins.md)
    
- [Read data from a binding](../../docs/how-to/bind-to-regions-in-a-document-or-spreadsheet.md#BindRegions_Read)
    
- [Write data to a binding](../../docs/how-to/bind-to-regions-in-a-document-or-spreadsheet.md#BindRegions_Write)
    
