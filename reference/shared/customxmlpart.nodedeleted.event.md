
# CustomXmlPart.nodeDeleted event
Occurs when a node is deleted.

|||
|:-----|:-----|
|**Hosts:**|Word|
|**Available in [Requirement set](http://msdn.microsoft.com/library/6b6702f2-b0a5-46ab-a356-8dda897ca8ae%28Office.15%29.aspx)**|CustomXmlParts|
|**Last changed in**|1.1|

```
Office.EventType.NodeDeleted
```


## Remarks

To add an event handler for the  **nodeDeleted** event, use the [addHandlerAsync](../../reference/shared/customxmlpart.addhandlerasync.md) method of the **CustomXmlPart** object.


## Example




```js
function addNodeDeletedEvent() {
    Office.context.document.customXmlParts.getByIdAsync("{3BC85265-09D6-4205-B665-8EB239A8B9A1}", function (result) {
        var xmlPart = result.value;
        xmlPart.addHandlerAsync(Office.EventType.NodeDeleted, function (eventArgs) {
            write("A node has been deleted.");
        });
    });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## Support details


A capital Y in the following matrix indicates that this event is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this event.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](http://msdn.microsoft.com/library/67340567-bb9a-498c-96d3-3f52f28c16bc%28Office.15%29.aspx).

||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||||
|**Excel**||||
|**Word**|Y||Y|

|||
|:-----|:-----|
|**Available in requirement sets**|CustomXmlParts|
|**Minimum permission level**|[ReadWriteDocument](http://msdn.microsoft.com/library/da2efadc-4ebf-45fe-be39-397ac1eb1dbd%28Office.15%29.aspx)|
|**Add-in types**|Task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history



****


|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Word in Office for iPad.|
|1.0|Introduced|
