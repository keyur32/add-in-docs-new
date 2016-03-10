
# Document.getActiveViewAsync method
 Returns the state of the current view of the presentation (edit or read).

|||
|:-----|:-----|
|**Hosts:** Excel, PowerPoint, Word|**Add-in types:** Content, Task pane|
|**Available in [Requirement set](http://msdn.microsoft.com/library/6b6702f2-b0a5-46ab-a356-8dda897ca8ae%28Office.15%29.aspx)**|ActiveView|
|**Added in ActiveView**|1.1|

```
Office.context.document.getActiveViewAsync([,options], callback);
```


## Parameters



|**Name**|**Type**|**Description**|**Support notes**|
|:-----|:-----|:-----|:-----|
| _options_|**object**|Specifies any of the following [optional parameters](http://msdn.microsoft.com/library/7fe6bb42-3178-4d96-85f5-af5caea7b950%28Office.15%29.aspx#AsyncProgramming_OptionalParameters)||
| _asyncContext_|**array**,  **boolean**,  **null**,  **number**,  **object**, **string**, or  **undefined**|A user-defined item of any type that is returned in the  **AsyncResult** object without being altered.||
| _callback_|**object**|A function that is invoked when the callback returns, whose only parameter is of type  **AsyncResult**.||

## Callback Value

When the function you passed to the  _callback_ parameter executes, it receives an [AsyncResult](../../reference/shared/asyncresult.md) object that you can access from the callback function's only parameter.

In the callback function passed to the  **getActiveViewAsync** method, the [AsyncResult.value](../../reference/shared/asyncresult.value.md) property returns the state of the presentation's current view. The value returned can be either `edit` or `read`.  `edit` corresponds to any of the views in which you can edit slides, such as **Normal** or **Outline View**.  `read` corresponds to either **Slide Show** or **Reading View**.


## Remarks

Can trigger an event when the view changes.


## Example

To get the view of the current presentation, you need to write a callback function that returns that value. The following example shows how to:


-  **Pass an anonymous callback function** that returns the view type to the _callback_ parameter of the **getActiveViewAsync** method.
    
-  **Display the value** on the add-in's page.
    

```js
function getFileView() {
    // Get whether the current view is edit or read.
    Office.context.document.getActiveViewAsync(function (asyncResult) {
        if (asyncResult.status == "failed") {
            showMessage("Action failed with error: " + asyncResult.error.message);
        }
        else {
            showMessage(asyncResult.value);
        }
    });
}
```




## Support details


A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](http://msdn.microsoft.com/library/67340567-bb9a-498c-96d3-3f52f28c16bc%28Office.15%29.aspx).


**Supported hosts, by platform**


||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|||Y|
|**PowerPoint**|Y|Y|Y|
|**Word**|||Y|

|||
|:-----|:-----|
|**Available in requirement sets**|ActiveView|
|**Added in ActiveView**|1.1|
|**Minimum permission level**|[Restricted](http://msdn.microsoft.com/library/da2efadc-4ebf-45fe-be39-397ac1eb1dbd%28Office.15%29.aspx)|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history





****


|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Excel, PowerPoint, and Word in Office for iPad.|
|1.1|Introduced.|
