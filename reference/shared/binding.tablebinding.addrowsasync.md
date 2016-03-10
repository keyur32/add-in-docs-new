
# TableBinding.addRowsAsync method
Adds rows and values to a table.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, Word|
|**Available in [Requirement set](http://msdn.microsoft.com/library/6b6702f2-b0a5-46ab-a356-8dda897ca8ae%28Office.15%29.aspx)**|TableBindings|
|**Last changed in**|1.1|

```
bindingObj.addRowsAsync(rows, [,options], callback);
```


## Parameters

_rows_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Type: **Array**

&nbsp;&nbsp;&nbsp;&nbsp;An array of arrays that contains one or more rows of data to add to the table. Required.
    
_options_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Type: **object**

&nbsp;&nbsp;&nbsp;&nbsp;Specifies the following [optional parameters](http://msdn.microsoft.com/library/7fe6bb42-3178-4d96-85f5-af5caea7b950%28Office.15%29.aspx#AsyncProgramming_OptionalParameters).
    
||
|:-----|
|<dl class="authored" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:MSHelp="http://msdn.microsoft.com/mshelp" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><dt><span class="parameter" sdata="paramReference">asyncContext</span></dt><dd><p>Type: <span class="keyword">array</span>, <span class="keyword">Boolean</span>, <span class="keyword">null</span>, <span class="keyword">number</span>, <b>object</b> , <span class="keyword">string</span>, or <span class="keyword">undefined</span></p><p>A user-defined item of any type that is returned in the <b>AsyncResult</b>  object without being altered. Optional.</p></dd></dl>|
-  _callback_   - Type:  **object**
    
    A function that is invoked when the callback returns, whose only parameter is of type [AsyncResult](../../reference/shared/asyncresult.md). Optional.
    


|**Name**|**Type**|**Description**|**Support notes**|
|:-----|:-----|:-----|:-----|
| _rows_|**array**|An array of arrays that contains one or more rows of data to add to the table. Required.||
| _options_|**object**|Specifies any of the following [optional parameters](http://msdn.microsoft.com/library/7fe6bb42-3178-4d96-85f5-af5caea7b950%28Office.15%29.aspx#AsyncProgramming_OptionalParameters).||
| _asyncContext_|**array**,  **boolean**,  **null**,  **number**,  **object**, **string**, or  **undefined**|A user-defined item of any type that is returned in the  **AsyncResult** object without being altered.||
| _callback_|**object**|A function that is invoked when the callback returns, whose only parameter is of type  **AsyncResult**.||

## Callback Value

When the function you passed to the  _callback_ parameter executes, it receives an [AsyncResult](../../reference/shared/asyncresult.md) object that you can access from the callback function's only parameter.

In the callback function passed to the  **addRowsAsync** method, you can use the properties of the **AsyncResult** object to return the following information.



|**Property**|**Use to...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Always returns  **undefined** because there is no object or data to retrieve.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Determine the success or failure of the operation.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Access an [Error](../../reference/shared/error.md) object that provides error information if the operation failed.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Access your user-defined  **object** or value, if you passed one as the _asyncContext_ parameter.|

## Remarks

The success or failure of an  **addRowsAsync** operation is atomic. That is, the entire add rows operation must succeed, or it will be completely rolled back (and the **AsyncResult.status** property returned to the callback will report failure):


- Each row in the array you pass as the  _data_ argument must have the same number of columns as the table being updated. If not, the entire operation will fail.
    
- Each row and cell in the array must successfully add that row and cell to the table in the newly added row(s). If any row or cell fails to be set for any reason, the entire operation will fail.
    
 **Additional remarks for Excel Online**

The total number of cells in the value passed to the  _rows_ parameter can't exceed 20,000 in a single call to this method.


## Example




```js
function addRowsToTable() {
    Office.context.document.bindings.getByIdAsync("myBinding", function (asyncResult) {
        var binding = asyncResult.value;
        binding.addRowsAsync([["6", "k"], ["7", "j"]]);
    });
}

```




## Support details


A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](http://msdn.microsoft.com/library/67340567-bb9a-498c-96d3-3f52f28c16bc%28Office.15%29.aspx).


**Supported hosts, by platform**


||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**Available in requirement sets**|TableBindings|
|**Minimum permission level**|[ReadWriteDocument](http://msdn.microsoft.com/library/da2efadc-4ebf-45fe-be39-397ac1eb1dbd%28Office.15%29.aspx)|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history




|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Excel and Word in Office for iPad|
|1.1|Added support for writing table data in add-ins for Access.|
|1.0|Introduced|
