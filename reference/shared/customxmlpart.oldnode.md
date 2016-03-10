
# NodeDeletedEventArgs.oldNode property
Gets the node which was just deleted from the  **CustomXmlPart** object.

|||
|:-----|:-----|
|**Hosts:**|Word|
|**Available in [Requirement set](http://msdn.microsoft.com/library/6b6702f2-b0a5-46ab-a356-8dda897ca8ae%28Office.15%29.aspx)**|CustomXmlParts|
|**Last changed in**|1.1|

```
var myNode = eventArgsObj.oldNode;
```


## Return Value

A [CustomXmlNode](../../reference/shared/customxmlnode.customxmlnode.md) that represents the node that was just deleted.


## Remarks

Note that this node may have children, if a subtree is being removed from the document. Also, this node will be a "disconnected" node in that you can query down from the node, but you cannot query up the tree - the node appears to exist alone.


## Support details


A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](http://msdn.microsoft.com/library/67340567-bb9a-498c-96d3-3f52f28c16bc%28Office.15%29.aspx).

||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**Available in requirement sets**|CustomXmlParts|
|**Minimum permission level**|[ReadWriteDocument](http://msdn.microsoft.com/library/da2efadc-4ebf-45fe-be39-397ac1eb1dbd%28Office.15%29.aspx)|
|**Add-in types**|Task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history




|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Word in Office for iPad.|
|1.0|Introduced|
