
# SelectionMode enumeration
Specifies whether to select (highlight) the location to navigate to (when using the [Document.goToByIdAsync](../../reference/shared/document.gotobyidasync.md) method).

|||
|:-----|:-----|
|**Introduced in Office.js version**|1.1|

|||
|:-----|:-----|
|**Hosts:**|Excel, PowerPoint, Word|
|**Added in**|1.1|



```
Office.SelectionMode
```


## Members


**Values**


|**Enumeration**|**Value**|**Description**|
|:-----|:-----|:-----|
|Office.SelectionMode.Selected|"selected"|The location will be selected (highlighted).|
|Office.SelectionMode.None|"none"|The cursor is moved the beginning of the location.|

## Support details


A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](http://msdn.microsoft.com/library/67340567-bb9a-498c-96d3-3f52f28c16bc%28Office.15%29.aspx).


**Supported hosts, by platform**


||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**PowerPoint**|Y|||
|**Word**|Y||Y|

|||
|:-----|:-----|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history



****


|**Version**|**Changes**|
|:-----|:-----|
|1.1|Introduced|
