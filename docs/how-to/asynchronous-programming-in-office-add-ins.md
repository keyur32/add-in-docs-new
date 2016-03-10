
# Asynchronous programming in Office Add-ins
Develop Office Add-ins using the nested callbacks and promises patterns asynchronous programming patterns supported by the JavaScript API for Office. 


 **Why does the Office Add-ins API use asynchronous programming?** Because JavaScript is a single-threaded language, if script invokes a long-running synchronous process, all subsequent script execution will be blocked until that process completes. Because certain operations against Office web clients (but rich clients as well) could block execution if they are run synchronously, most of the methods in the JavaScript API for Office are designed to execute asynchronously. This makes sure that Office Add-ins are responsive and highly performing. It also frequently requires you to write callback functions when working with these asynchronous methods.

The names of all asynchronous methods in the API end with "Async", such as the  [Document.getSelectedDataAsync](http://msdn.microsoft.com/en-us/library/f85ad02c-64f0-4b73-87f6-7f521b3afd69%28Office.15%29.aspx), [Binding.getDataAsync](http://msdn.microsoft.com/en-us/library/5372ffd8-579d-4fcb-9e5b-e9a2128f3201%28Office.15%29.aspx), or [Item.loadCustomPropertiesAsync](https://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html) methods. When an "Async" method is called, it executes immediately and any subsequent script execution can continue. The optional callback function you pass to an "Async" method executes as soon as the data or requested operation is ready. This generally occurs promptly, but there can be a slight delay before it returns.

The following diagram shows the flow of execution for a call to an "Async" method that reads the data the user selected in a document open in the server-based Word Online or Excel Online. At the point when the "Async" call is made, the JavaScript execution thread is free to perform any additional client-side processing. (Although none are shown in the diagram.) When the "Async" method returns, the callback resumes execution on the thread, and the add-in can the access data, do something with it, and display the result. The same asynchronous execution pattern holds when working with the Office rich client host applications, such as Word 2013 or Excel 2013.

**Figure 1. Asynchronous programing execution flow**


![Asynchronous programming thread execution flow](../../images/off15appAsyncProgFig01.png)

Support for this asynchronous design in both rich and web clients is part of the "write once-run cross-platform" design goals of the Office Add-ins development model. For example, you can create a content or task pane add-in with a single code base that will run in both Excel 2013 and Excel Online.

## Writing the callback function for an "Async" method


The callback function you pass as the  _callback_ argument to an "Async" method must declare a single parameter that the add-in runtime will use to provide access to an [AsyncResult](http://msdn.microsoft.com/en-us/library/540c114f-0398-425c-baf3-7363f2f6bc47%28Office.15%29.aspx) object when the callback function executes. You can write:


- An anonymous function that must be written and passed directly in line with the call to the "Async" method as the  _callback_ parameter of the "Async" method.
    
- A named function, passing the name of that function as the  _callback_ parameter of an "Async" method.
    
An anonymous function is useful if you are only going to use its code once - because it has no name, you can't reference it in another part of your code. A named function is useful if you want to reuse the callback function for more than one "Async" method.


### Writing an anonymous callback function

The following anonymous callback function declares a single parameter named  `result` that retrieves data from the [AsyncResult.value](http://msdn.microsoft.com/en-us/library/453a4b43-0fdc-4ea9-967a-c033fab31507%28Office.15%29.aspx) property when the callback returns.


```
function (result) {
        write('Selected data: ' + result.value);
}
```

The following example shows how to pass this anonymous callback function in line in the context of a full "Async" method call to the  **Document.getSelectedDataAsync** method.


- The first  _coercionType_ argument, `Office.CoercionType.Text`, specifies to return the selected data as a string of text.
    
- The second  _callback_ argument is the anonymous function passed in-line to the method. When the function executes, it uses the _result_ parameter to access the **value** property of the **AsyncResult** object to display the data selected by the user in the document.
    



```js
Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, 
    function (result) {
        write('Selected data: ' + result.value);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

You can also use the parameter of your callback function to access other properties of the  **AsyncResult** object. Use the [AsyncResult.status](http://msdn.microsoft.com/en-us/library/eec9c712-79eb-4365-88a1-6d77649727c1%28Office.15%29.aspx) property to determine if the call succeeded or failed. If your call fails you can use the [AsyncResult.error](http://msdn.microsoft.com/en-us/library/51c46d36-972d-4d82-91aa-da99cbeb8d4f%28Office.15%29.aspx) property to access an [Error](http://msdn.microsoft.com/en-us/library/36d1d048-b888-4bb5-9321-d340bcbc86f4%28Office.15%29.aspx) object for error information.

For more information about using the  **getSelectedDataAsync** method, see [Read and write data to the active selection in a document or spreadsheet](../../docs/how-to/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md). 


### Writing a named callback function

Alternatively, you can write a named function and pass its name to the  _callback_ parameter of an "Async" method. For example, the previous example can be rewritten to pass a function named `writeDataCallback` as the _callback_ parameter like this.


```
Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, 
    writeDataCallback);

// Callback to write the selected data to the add-in UI.
function writeDataCallback(result) {
    write('Selected data: ' + result.value);
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## Differences in what's returned to the AsyncResult.value property


The  **asyncContext**,  **status**, and  **error** properties of the **AsyncResult** object return the same kinds of information to the callback function passed to all "Async" methods. However, what's returned to the **AsyncResult.value** property varies depending on the functionality of the "Async" method.

For example, the  **addHandlerAsync** methods (of the [Binding](http://msdn.microsoft.com/en-us/library/42882642-d22b-47d2-a8d3-3aa8c6a4435e%28Office.15%29.aspx), [CustomXmlPart](http://msdn.microsoft.com/en-us/library/83f0e668-8236-4f2f-a20f-b173a9e3f65f%28Office.15%29.aspx), [Document](http://msdn.microsoft.com/en-us/library/f8859516-cc1f-4b20-a8f3-cee37a983e70%28Office.15%29.aspx), [RoamingSettings](https://dev.outlook.com/reference/add-ins/RoamingSettings.html), and [Settings](http://msdn.microsoft.com/en-us/library/ad733387-a58c-4514-8fc2-53e64fad468d%28Office.15%29.aspx) objects) are used to add event handler functions to the items represented by these objects. You can access the **AsyncResult.value** property from the callback function you pass to any of the **addHandlerAsync** methods, but since no data or object is being accessed when you add an event handler, the **value** property always returns **undefined** if you attempt to access it.

On the other hand, if you call the  **Document.getSelectedDataAsync** method, it returns the data the user selected in the document to the **AsyncResult.value** property in the callback. Or, if you call the [Bindings.getAllAsync](http://msdn.microsoft.com/en-us/library/ef902b73-cc4c-4551-95de-d8a51eeba82f%28Office.15%29.aspx) method, it returns an array of all of the **Binding** objects in the document. And, if you call the [Bindings.getByIdAsync](http://msdn.microsoft.com/en-us/library/2727c891-bc05-465c-9324-113fbfeb3fbb%28Office.15%29.aspx) method, it returns a single **Binding** object.

For a description of what's returned to the  **AsyncResult.value** property for an "Async" method, see the "Callback value" section of that method's reference topic. For a summary of all of the objects that provide "Async" methods, see the table at the bottom of the [AsyncResult](http://msdn.microsoft.com/en-us/library/540c114f-0398-425c-baf3-7363f2f6bc47%28Office.15%29.aspx) object topic.


## Asynchronous programming patterns


The JavaScript API for Office supports two kinds of asynchronous programming patterns:


- Using nested callbacks
    
- Using the promises pattern
    
Asynchronous programming with callback functions frequently requires you to nest the returned result of one callback within two or more callbacks. If you need to do so, you can use nested callbacks from all "Async" methods of the API.

Using nested callbacks is a programming pattern familiar to most JavaScript developers, but code with deeply nested callbacks can be difficult to read and understand. As an alternative to nested callbacks, the JavaScript API for Office also supports an implementation of the promises pattern. However, in the current version of the JavaScript API for Office, the promises pattern only works with code for [bindings in Excel spreadsheets and Word documents](http://msdn.microsoft.com/en-us/library/5bf788db-d788-4d91-bcb6-fc3913b40012%28Office.15%29.aspx).


### Asynchronous programming using nested callback functions


Frequently, you need to perform two or more asynchronous operations to complete a task. To accomplish that, you can nest one "Async" call inside another. 

The following code example nests two asynchronous calls. 


- First, the [Bindings.getByIdAsync](http://msdn.microsoft.com/en-us/library/2727c891-bc05-465c-9324-113fbfeb3fbb%28Office.15%29.aspx) method is called to access a binding in the document named "MyBinding". The **AsyncResult** object returned to the `result` parameter of that callback provides access to the specified binding object from the **AsyncResult.value** property.
    
- Then, the binding object accessed from the first  `result` parameter is used to call the [Binding.getDataAsync](http://msdn.microsoft.com/en-us/library/5372ffd8-579d-4fcb-9e5b-e9a2128f3201%28Office.15%29.aspx) method.
    
- Finally, the  `result2` parameter of the callback passed to the **Binding.getDataAsync** method is used to display the data in the binding.
    



```js
function readData() {
    Office.context.document.bindings.getByIdAsync("MyBinding", function (result) {
        result.value.getDataAsync({ coercionType: 'text' }, function (result2) {
            write(result2.value);
        });
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

This basic nested callback pattern can be used for all asynchronous methods in the JavaScript API for Office.

The following sections show how to use either anonymous or named functions for nested callbacks in asynchronous methods.


#### Using anonymous functions for nested callbacks

In the following example, two anonymous functions are declared inline and passed into the  **getByIdAsync** and **getDataAsync** methods as nested callbacks. Because the functions are simple and inline, the intent of the implementation is immediately clear.


```js
Office.context.document.bindings.getByIdAsync('myBinding', function (bindingResult) {
    bindingResult.value.getDataAsync(function (getResult) {
        if (getResult.status == Office.AsyncResultStatus.Failed) {
            write('Action failed. Error: ' + asyncResult.error.message);
        } else {
            write('Data has been read successfully.');
        }
    });
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


#### Using named functions for nested callbacks

In complex implementations, it may be helpful to use named functions to make your code easier to read, maintain, and reuse. In the following example, the two anonymous functions from the example in the previous section have been rewritten as functions named  `deleteAllData` and `showResult`. These named functions are then passed into the  **getByIdAsync** and **deleteAllDataValuesAsync** methods as callbacks by name.


```js
Office.context.document.bindings.getByIdAsync('myBinding', deleteAllData);

function deleteAllData(asyncResult) {
    asyncResult.value.deleteAllDataValuesAsync(showResult);
}

function showResult(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } else {
        write('Data has been deleted successfully.');
    }
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


### Asynchronous programming using the promises pattern to access data in bindings


Instead of passing a callback function and waiting for the function to return before execution continues, the promises programming pattern immediately returns apromise object that represents its intended result. However, unlike true synchronous programming, under the covers the fulfillment of the promised result is actually deferred until the Office Add-ins runtime environment can complete the request. An _onError_ handler is provided to cover situations when the request can't be fulfilled.

The JavaScript API for Office provides the [Office.select](http://msdn.microsoft.com/en-us/library/23aeb136-da1f-4127-a798-99dc27bc4dae%28Office.15%29.aspx) method to support the promises pattern for working with existing binding objects. The promise object returned to the **Office.select** method supports only the four methods that you can access directly from the [Binding](http://msdn.microsoft.com/en-us/library/42882642-d22b-47d2-a8d3-3aa8c6a4435e%28Office.15%29.aspx) object: [getDataAsync](http://msdn.microsoft.com/en-us/library/5372ffd8-579d-4fcb-9e5b-e9a2128f3201%28Office.15%29.aspx), [setDataAsync](http://msdn.microsoft.com/en-us/library/6a59bb6d-40b6-4a95-9b98-d70d4616de09%28Office.15%29.aspx), [addHandlerAsync](http://msdn.microsoft.com/en-us/library/b9c2f4ea-726c-4b48-a3fb-89beda337a17%28Office.15%29.aspx), and [removeHandlerAsync](http://msdn.microsoft.com/en-us/library/5ae3a860-1fc4-46ce-858e-98545c3e2d77%28Office.15%29.aspx).

The promises pattern for working with bindings takes this form:

 **Office.select(**_selectorExpression_,  _onError_**).**_BindingObjectAsyncMethod_

The  _selectorExpression_ parameter takes the form `"bindings#bindingId"`, where  _bindingId_ is the name ( **id**) of a binding that you created previously in the document or spreadsheet (using one of the "addFrom" methods of the  **Bindings** collection: **addFromNamedItemAsync**,  **addFromPromptAsync**, or  **addFromSelectionAsync**). For example, the selector expression  `bindings#cities` specifies that you want to access the binding with an **id** of 'cities'.

The  _onError_ parameter is an error handling function which takes a single parameter of type **AsyncResult** that can be used to access an **Error** object, if the **select** method fails to access the specified binding. The following example shows a basic error handler function that can be passed to the _onError_ parameter.




```js
function onError(result){
    var err = result.error;
    write(err.name + ": " + err.message);
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

Replace the  _BindingObjectAsyncMethod_ placeholder with a call to any of the four **Binding** object methods supported by the promise object: **getDataAsync**,  **setDataAsync**,  **addHandlerAsync**, or  **removeHandlerAsync**. Calls to these methods don't support additional promises. You must call them using the [nested callback function pattern](http://msdn.microsoft.com/en-us/library/7fe6bb42-3178-4d96-85f5-af5caea7b950%28Office.15%29.aspx#AsyncProgramming_NestedCallbacks).

After a  **Binding** object promise is fulfilled, it can be reused in the chained method call as if it were a binding (the add-in runtime won't asynchronously retry fulfilling the promise). If the **Binding** object promise can't be fulfilled, the add-in runtime will try again to access the binding object the next time one of its asynchronous methods is invoked.

The following code example uses the  **select** method to retrieve a binding with the **id** " `cities`" from the  **Bindings** collection, and then calls the [addHandlerAsync](http://msdn.microsoft.com/en-us/library/b9c2f4ea-726c-4b48-a3fb-89beda337a17%28Office.15%29.aspx) method to add an event handler for the [dataChanged](http://msdn.microsoft.com/en-us/library/7b9ed4bf-3ce5-44eb-8548-2b081afd868d%28Office.15%29.aspx) event of the binding.




```js
function addBindingDataChangedEventHandler() {
    Office.select("bindings#cities", function onError(){/* error handling code */}).addHandlerAsync(Office.EventType.BindingDataChanged,
    function (eventArgs) {
        doSomethingWithBinding(eventArgs.binding);
    });
}

```


 >**Important**  The  **Binding** object promise returned by the **Office.select** method provides access to only the four methods of the **Binding** object. If you need to access any of the other members of the **Binding** object, instead you must use the **Document.bindings** property and **Bindings.getByIdAsync** or **Bindings.getAllAsync** methods to retrieve the **Binding** object. For example, if you need to access any of the **Binding** object's properties (the **document**,  **id**, or  **type** properties), or need to access the properties of the [MatrixBinding](http://msdn.microsoft.com/en-us/library/35e8568e-9129-4c00-b30f-d8c3b2555f1e%28Office.15%29.aspx) or [TableBinding](http://msdn.microsoft.com/en-us/library/1508795b-1c70-456c-b3bf-666d40cf8f50%28Office.15%29.aspx) objects, you must use the **getByIdAsync** or **getAllAsync** methods to retrieve a **Binding** object.


## Passing optional parameters to asynchronous methods


The common syntax for all "Async" methods follows this pattern:

 _AsyncMethod_ `(` _RequiredParameters_ `, [` _OptionalParameters_ `],` _CallbackFunction_ `);`

All asynchronous methods support optional parameters, which are passed in as a JavaScript Object Notation (JSON) object that contains one or more optional parameters. The JSON object containing the optional parameters is an unordered collection of key-value pairs with the ":" character separating the key and the value. Each pair in the object is comma-separated, and the entire set of pairs is enclosed in braces. The key is the parameter name, and value is the value to pass for that parameter.

You can create the JSON object that contains optional parameters inline, or by creating an  `options` object and passing that in as the _options_ parameter.


### Passing optional parameters inline

For example, the syntax for calling the [Document.setSelectedDataAsync](http://msdn.microsoft.com/en-us/library/f85ad02c-64f0-4b73-87f6-7f521b3afd69%28Office.15%29.aspx) method with optional parameters inline looks like this:

 `Office.context.document.setSelectedDataAsync(` _data_ `, {coercionType:` _coercionType_ `, asyncContext:` _asyncContext_ `},` _callback_ `);`

In this form of the calling syntax, the two optional parameters,  _coercionType_ and _asyncContext_, are defined as a JSON object inline enclosed in braces.

The following example shows how to call to the  **Document.setSelectedDataAsync** method by specifying optional parameters inline.




```
Office.context.document.setSelectedDataAsync(
    "<html><body>hello world</body></html>",
    {coercionType: "html", asyncContext: 42},
    function(asyncResult) {
        write(asyncResult.status + " " + asyncResult.asyncContext);
    }
)

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


 >**Note**  You can specify optional parameters in any order in the JSON object as long as their names are specified correctly.


### Passing optional parameters in an options object

Alternatively, you can create an object named  `options` that specifies the optional parameters separately from the method call, and then pass the `options` object as the _options_ argument.

The following example shows one way of creating the  `options` object, where `parameter1`,  `value1`, and so on, are placeholders for the actual parameter names and values.




```
var options = {
    parameter1: value1,
    parameter2: value2,
    ...
    parameterN: valueN
};

```

Which looks like the following example when used to specify the [ValueFormat](http://msdn.microsoft.com/en-us/library/75e4a0f9-e0c6-4c8b-ac87-95b824356a4e%28Office.15%29.aspx) and [FilterType](http://msdn.microsoft.com/en-us/library/1d182c44-526d-4f7e-9557-78534f845e5b%28Office.15%29.aspx) parameters.




```
var options = {
    valueFormat: "unformatted",
    filterType: "all"
};
```

Here's another way of creating the  `options` object.




```
var options = {};
options[parameter1] = value1;
options[parameter2] = value2;
...
options[parameterN] = valueN;
```

Which looks like the following example when used to specify the  **ValueFormat** and **FilterType** parameters.:




```
var options = {};
options["ValueFormat"] = "unformatted";
options["FilterType"] = "all";
```


 >**Note**  When using either method of creating the  `options` object, you can specify optional parameters in any order as long as their names are specified correctly.

The following example shows how to call to the  **Document.setSelectedDataAsync** method by specifying optional parameters in an `options` object.




```
var options = {
   coercionType: "html",
   asyncContext: 42
};

document.setSelectedDataAsync(
    "<html><body>hello world</body></html>",
    options,
    function(asyncResult) {
        write(asyncResult.status + " " + asyncResult.asyncContext);
    }
)

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


 >**Note**  
 
In both optional parameter examples, the  _callback_ parameter is specified as the last parameter (following the inline optional parameters, or following the _options_ argument object). Alternatively, you can specify the _callback_ parameter inside either the inline JSON object, or in the `options` object. However, you can pass the _callback_ parameter in only one location: either in the _options_ object (inline or created externally), or as the last parameter, but not both.


## Additional resources



- [Understanding the JavaScript API for Office](../../docs/overview/understanding-the-javascript-api-for-office.md)
    
- [JavaScript API for Office](http://msdn.microsoft.com/library/b27e70c3-d87d-4d27-85e0-103996273298%28Office.15%29.aspx)
    
