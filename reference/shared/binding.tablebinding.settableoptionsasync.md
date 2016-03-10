
# TableBinding.setTableOptionsAsync method
Updates table formatting options on the bound table.

|||
|:-----|:-----|
|**Hosts:**|Excel|
|**Available in [Requirement set](http://msdn.microsoft.com/library/6b6702f2-b0a5-46ab-a356-8dda897ca8ae%28Office.15%29.aspx)**|Not in a set|
|**Added in**|1.1|

```
bindingObj.setTableOptionsAsync(tableOptions [,options] , callback);
```


## Parameters



|**Name**|**Type**|**Description**|**Support notes**|
|:-----|:-----|:-----|:-----|
| _tableOptions_|**object**|An object literal containing a list of property name-value pairs that define the table options to apply. Required.||
| _options_|**object**|Specifies any of the following [optional parameters](http://msdn.microsoft.com/library/7fe6bb42-3178-4d96-85f5-af5caea7b950%28Office.15%29.aspx#AsyncProgramming_OptionalParameters)||
| _asyncContext_|**array**,  **boolean**,  **null**,  **number**,  **object**, **string**, or  **undefined**|A user-defined item of any type that is returned in the  **AsyncResult** object without being altered.||
| _callback_|**object**|A function that is invoked when the callback returns, whose only parameter is of type  **AsyncResult**.||

## Callback Value

When the function you passed to the  _callback_ parameter executes, it receives an [AsyncResult](../../reference/shared/asyncresult.md) object that you can access from the callback function's only parameter.

In the callback function passed to the  **goToByIdAsync** method, you can use the properties of the **AsyncResult** object to return the following information.



|**Property**|**Use to...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Always returns  **undefined** because there is no data or object to retrieve when setting table options.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Determine the success or failure of the operation.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Access an [Error](../../reference/shared/error.md) object that provides error information if the operation failed.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Access your user-defined  **object** or value, if you passed one as the _asyncContext_ parameter.|

## Example

The following example shows how to:


-  **Create an object literal** that specifies the [table formatting options](http://msdn.microsoft.com/library/46b05707-b350-41be-b6b8-311799c71a33%28Office.15%29.aspx) to update on the bound table.
    
-  **Call setTableOptions** on a previously bound table (with an **id** of `myBinding`) passing the object with formatting settin as the  _tableOptions_ parameter.
    

```js
function updateTableFormatting(){
    var tableOptions = {bandedRows: true, filterButton: false, style: "TableStyleMedium3"}; 

    Office.select("bindings#myBinding").setTableOptionsAsync(tableOptions, function(asyncResult){});
}
```




## Support details


A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](http://msdn.microsoft.com/library/67340567-bb9a-498c-96d3-3f52f28c16bc%28Office.15%29.aspx).


**Supported hosts, by platform**


||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|

|||
|:-----|:-----|
|**Available in requirement sets**|Not in a set.|
|**Minimum permission level**|[WriteDocument](http://msdn.microsoft.com/library/da2efadc-4ebf-45fe-be39-397ac1eb1dbd%28Office.15%29.aspx)|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history




|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Excel in Office for iPad.|
|1.1|Introduced|
