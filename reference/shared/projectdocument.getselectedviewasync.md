

# ProjectDocument.getSelectedViewAsync method
Asynchronously gets the type and name of the active view in the document.

|||
|:-----|:-----|
|**Hosts:**|Project|
|**Available in [Requirement set](http://msdn.microsoft.com/library/6b6702f2-b0a5-46ab-a356-8dda897ca8ae%28Office.15%29.aspx)**|Selection|
|**Added in**|1.0|

```js
Office.context.document.getSelectedViewAsync([options,] [callback]);
```


## Parameters



|**Name**|**Type**|**Description**|**Support notes**|
|:-----|:-----|:-----|:-----|
| _options_|**object**|Specifies any of the following [optional parameters](http://msdn.microsoft.com/library/7fe6bb42-3178-4d96-85f5-af5caea7b950%28Office.15%29.aspx#AsyncProgramming_OptionalParameters).||
| _asyncContext_|**array**,  **boolean**,  **null**,  **number**,  **object**, **string**, or  **undefined**|A user-defined item of any type that is returned in the  **AsyncResult** object without being altered.||
| _callback_|**object**|A function that is invoked when the callback returns, whose only parameter is of type  **AsyncResult**.||

## Callback Value

When the  _callback_ function executes, it receives an [AsyncResult](../../reference/shared/asyncresult.md) object that you can access from the parameter in the callback function.

For the  **getSelectedViewAsync** method, the returned [AsyncResult](../../reference/shared/asyncresult.md) object contains the following properties.


****


|**Name**|**Description**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|The data passed in the optional  _asyncContext_ parameter, if the parameter was used.|
|[error](../../reference/shared/asyncresult.error.md)|Information about the error, if the  **status** property equals **failed**.|
|[status](../../reference/shared/asyncresult.status.md)|The  **succeeded** or **failed** status of the asynchronous call.|
|[value](../../reference/shared/asyncresult.value.md)|Contains the following properties:<br/><br/><div>* **viewName** - The name of the view, as a [ProjectViewTypes](../../reference/shared/projectviewtypes-enumeration.md) constant.<br/>* **viewType** - The type of view, as the integer value of a [ProjectViewTypes](../../reference/shared/projectviewtypes-enumeration.md) constant.</div>|

## Example

The following code example calls adds a [ViewSelectionChanged](../../reference/shared/projectdocument.viewselectionchanged.event.md) event handler that calls **getSelectedViewAsync** to get the name and type of the active view in the document.

The example assumes your add-in has a reference to the jQuery library and that the following page control is defined in the content div in the page body.




```HTML
<span id="message"></span>
```




```js
(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            // After the DOM is loaded, add-in-specific code can run.
            Office.context.document.addHandlerAsync(
                Office.EventType.ViewSelectionChanged,
                getActiveView);
            getActiveView();
        });
    };

    // Get the active view's name and type.
    function getActiveView() {
        Office.context.document.getSelectedViewAsync(
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    onError(result.error);
                }
                else {
                    var output = String.format(
                        'View name: {0}<br/>View type: {1}',
                        result.value.viewName, viewType);
                    $('#message').html(output);
                }
            }
        );
    }

    function onError(error) {
        $('#message').html(error.name + ' ' + error.code + ': ' + error.message);
    }
})();
```


## Support details


A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](http://msdn.microsoft.com/library/67340567-bb9a-498c-96d3-3f52f28c16bc%28Office.15%29.aspx).


||**Office for Windows desktop**|**Office Online (in browser)**|
|:-----|:-----|:-----|
|**Project**|Y||

|||
|:-----|:-----|
|**Available in requirement sets**|Selection|
|**Minimum permission level**|[ReadDocument](http://msdn.microsoft.com/library/da2efadc-4ebf-45fe-be39-397ac1eb1dbd%28Office.15%29.aspx)|
|**Add-in types**|Task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history



****


|**Version**|**Changes**|
|:-----|:-----|
|1.0|Introduced|

## See also



#### Other resources


[ProjectViewTypes enumeration](../../reference/shared/projectviewtypes-enumeration.md)
[AsyncResult object](../../reference/shared/asyncresult.md)
[ViewSelectionChanged event](../../reference/shared/projectdocument.viewselectionchanged.event.md)
[ProjectDocument object](../../reference/shared/projectdocument.projectdocument.md)
