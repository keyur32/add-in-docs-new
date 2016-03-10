
# ProjectDocument.setTaskFieldAsync method (JavaScript API for Office v1.1)
Asynchronously sets the value of the specified field for the specified task.
 **Important:** This API works only in Project 2016 on Windows desktop.

|||
|:-----|:-----|
|**Hosts:**|Project|
|**Available in [Requirement set](http://msdn.microsoft.com/library/6b6702f2-b0a5-46ab-a356-8dda897ca8ae%28Office.15%29.aspx)**|Selection|
|**Added in**|1.1|

```
Office.context.document.setTaskFieldAsync(taskId, fieldId, fieldValue[, options][, callback]);
```


## Parameters


-  _taskId_The GUID of the task. Required.
    
-  _fieldId_The ID of the target field, as a [ProjectTaskFields](../../reference/shared/projecttaskfields-enumeration.md) constant or its corresponding integer value. Required.
    
-  _fieldValue_The value for the target field, as a  **string**,  **number**,  **boolean**, or  **object**. Required.
    
-  _options_The following [optional parameter](http://msdn.microsoft.com/library/7fe6bb42-3178-4d96-85f5-af5caea7b950%28Office.15%29.aspx#AsyncProgramming_OptionalParameters):
    
||
|:-----|
|<dl class="authored" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:MSHelp="http://msdn.microsoft.com/mshelp" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><dt><span class="parameter" sdata="paramReference">asyncContext</span></dt><dd><p>Type: <span class="keyword">array</span>, <span class="keyword">boolean</span>, <span class="keyword">null</span>, <span class="keyword">number</span>, <b>object</b> , <span class="keyword">string</span>, or <span class="keyword">undefined</span></p><p>A user-defined item of any type that is returned in the <a href="540c114f-0398-425c-baf3-7363f2f6bc47.htm">AsyncResult</a> object without being altered. Optional.</p><p>For example, you can pass the <span class="parameter" sdata="paramReference">asyncContext</span> argument by using the format <span class="code">{asyncContext: 'Some text'}</span> or <span class="code">{asyncContext: <object>}</span>.</p></dd></dl>|
-  _callback_ Type:  **function**
    
    A function that is invoked when the method call returns, where the only parameter is of type [AsyncResult](../../reference/shared/asyncresult.md). Optional.
    

## Callback Value

When the  _callback_ function executes, it receives an [AsyncResult](../../reference/shared/asyncresult.md) object that you can access from the parameter in the callback function.

For the  **setTaskFieldAsync** method, the returned [AsyncResult](../../reference/shared/asyncresult.md) object contains following properties.


****


|**Name**|**Description**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|The data passed in the optional  _asyncContext_ parameter, if the parameter was used.|
|[error](../../reference/shared/asyncresult.error.md)|Information about the error, if the  **status** property equals **failed**.|
|[status](../../reference/shared/asyncresult.status.md)|The  **succeeded** or **failed** status of the asynchronous call.|
|[value](../../reference/shared/asyncresult.value.md)|This method does not return a value.|

## Remarks

First call the [getSelectedTaskAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) or [getTaskByIndexAsync](../../reference/shared/projectdocument.settaskfieldasync.md) method to get the task GUID, and then pass the GUID as the _taskId_ argument to **setTaskFieldAsync**. Only a single field for a single task can be updated in each asynchronous call.


## Example

The following code example calls [getSelectedTaskAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) to get the GUID of the task that's currently selected in a task view. Then it sets two task field values by calling **setTaskFieldAsync** recursively.

The [getSelectedTaskAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) method used in the example requires that a task view (for example, Task Usage) is the active view and that a task is selected. See the [addHandlerAsync](../../reference/shared/projectdocument.addhandlerasync.md) method for an example that activates a button based on the active view type.

The example assumes your add-in has a reference to the jQuery library and that the following page controls are defined in the content div in the page body.




```HTML
<input id="set-info" type="button" value="Set info" /><br />
<span id="message"></span>
```




```js
(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            
            // After the DOM is loaded, add-in-specific code can run.
            app.initialize();
            $('#set-info').click(setTaskInfo);
        });
    };

    // Get the GUID of the task, and then get the task fields.
    function setTaskInfo() {
        getTaskGuid().then(
            function (data) {
                setTaskFields(data);
            }
        );
    }

    // Get the GUID of the selected task.
    function getTaskGuid() {
        var defer = $.Deferred();
        Office.context.document.getSelectedTaskAsync(
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    onError(result.error);
                }
                else {
                    defer.resolve(result.value);
                }
            }
        );
        return defer.promise();
    }

    // Set the specified fields for the selected task.
    function setTaskFields(taskGuid) {
        var targetFields = [Office.ProjectTaskFields.Active, Office.ProjectTaskFields.Notes];
        var fieldValues = [true, 'Notes for the task.'];

        // Set the field value. If the call is successful, set the next field.
        for (var i = 0; i < targetFields.length; i++) {
            Office.context.document.setTaskFieldAsync(
                taskGuid,
                targetFields[i],
                fieldValues[i],
                function (result) {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        i++;
                    }
                    else {
                        onError(result.error);
                    }
                }
            );
        }
        $('#message').html('Field values set');
    }

    function onError(error) {
        app.showNotification(error.name + ' ' + error.code + ': ' + error.message);
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
|**Available in requirement sets**||
|**Minimum permission level**|[WriteDocument](http://msdn.microsoft.com/library/da2efadc-4ebf-45fe-be39-397ac1eb1dbd%28Office.15%29.aspx)|
|**Add-in types**|Task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history



****


|**Version**|**Changes**|
|:-----|:-----|
|1.1|Introduced|

## See also



#### Other resources


[getSelectedTaskAsync method](../../reference/shared/projectdocument.getselectedresourceasync.md)
[getTaskByIndexAsync](../../reference/shared/projectdocument.settaskfieldasync.md)
[AsyncResult object](../../reference/shared/asyncresult.md)
[ProjectTaskFields enumeration](../../reference/shared/projecttaskfields-enumeration.md)
[ProjectDocument object](../../reference/shared/projectdocument.projectdocument.md)
