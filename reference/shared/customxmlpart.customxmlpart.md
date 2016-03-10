
# CustomXmlPart object
Represents a single  **CustomXMLPart** in a [CustomXMLParts](../../reference/shared/customxmlparts.customxmlparts.md) collection.

|||
|:-----|:-----|
|**Hosts:**|Word|
|**Available in [Requirement set](http://msdn.microsoft.com/library/6b6702f2-b0a5-46ab-a356-8dda897ca8ae%28Office.15%29.aspx)**|CustomXmlParts|
|**Last changed in**|1.1|

```
Office.context.document.customXmlParts.getByIdAsync(id);
```


## Members


**Properties**


|**Name**|**Description**|
|:-----|:-----|
|[builtIn](../../reference/shared/customxmlpart.builtin.md)|Get a value that indicates whether the CustomXMLPart is built-in.|
|[id](../../reference/shared/customxmlpart.id.md)|Gets the GUID of the CustomXMLPart|
|[namespaceManager](../../reference/shared/customxmlpart.namespacemanager.md)|Gets the set of namespace prefix mappings (CustomXMLPrefixMappings) used against the current CustomXMLPart.|

**Methods**


|**Name**|**Description**|
|:-----|:-----|
|[addHandlerAsync](../../reference/shared/customxmlpart.addhandlerasync.md)|Asynchronously adds an event handler for a  **CustomXmlPart** object event.|
|[deleteAsync](../../reference/shared/customxmlpart.deleteasync.md)|Asynchronously deletes this custom XML part from the collection.|
|[getNodesAsync](../../reference/shared/customxmlpart.getnodesasync.md)|Asynchronously gets any CustomXmlNodes in this custom XML part which match the specified XPath.|
|[getXmlAsync](../../reference/shared/customxmlpart.getxmlasync.md)|Asynchronously gets the XML inside this custom XML part.|
|[removeHandlerAsync](../../reference/shared/customxmlpart.removehandlerasync.md)|Removes an event handler for a  **CustomXmlPart** object event.|

**Events**


|**Name**|**Description**|
|:-----|:-----|
|[nodeDeleted](../../reference/shared/customxmlpart.nodedeleted.event.md)|Occurs when a node is deleted.|
|[nodeInserted](../../reference/shared/customxmlpart.nodeinserted.event.md)|Occurs when a node is inserted.|
|[nodeReplaced](../../reference/shared/customxmlpart.nodereplaced.event.md)|Occurs when a node is replaced.|

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



****


|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Word in Office for iPad.|
|1.0|Introduced|
