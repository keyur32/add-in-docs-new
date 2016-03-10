
# Document.ActiveViewChanged event
Occurs when the user changes the current view of the document.

|||
|:-----|:-----|
|**Hosts:**|PowerPoint|
|**Introduced in**|1.1|

```
Office.EventType.ActiveViewChanged
```


## Remarks

To add an event handler for the  **ActiveViewChanged** event of a document, use the [addHandlerAsync](../../reference/shared/document.addhandlerasync.md) method of the **Document** object. The event handler receives an argument of type [ActiveViewChangedEventArgs](../../reference/shared/document.activeviewchangedeventargs.md).


## Support details


A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](http://msdn.microsoft.com/library/67340567-bb9a-498c-96d3-3f52f28c16bc%28Office.15%29.aspx).


**Supported hosts, by platform**


||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**PowerPoint**|Y|Y|Y|

|||
|:-----|:-----|
|**Introduced in**|1.1|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|
