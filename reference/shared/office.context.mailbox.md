
# Context.mailbox property
Gets the  **mailbox** object that provides access to API members specifically for Outlook add-ins.

|||
|:-----|:-----|
|**Hosts:**|Outlook|
|**Available in [Requirement set](http://msdn.microsoft.com/library/6b6702f2-b0a5-46ab-a356-8dda897ca8ae%28Office.15%29.aspx)**|Mailbox|
|**Last changed in**|1.0|

```
var outlookOm = Office.context.mailbox;
```


## Return value

The [mailbox](http://msdn.microsoft.com/library/a3880d3b-8a09-4cf9-9274-f2682cb3b769%28Office.15%29.aspx) object.


## Example

The following line of code access the [item](http://msdn.microsoft.com/library/ad288df1-3ca2-474c-bea4-c51f46e6fc43%28Office.15%29.aspx) object of the JavaScript API for Office.


```
// Access the Item object.
var item = Office.context.mailbox.item;

```




## Support details


A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](http://msdn.microsoft.com/library/67340567-bb9a-498c-96d3-3f52f28c16bc%28Office.15%29.aspx).


||**Office for Windows desktop**|**Office Online (in browser)**|**Outlook for Mac**|
|:-----|:-----|:-----|:-----|
|**Outlook**|Y|Y|Y|

|||
|:-----|:-----|
|**Available in requirement sets**|Mailbox|
|**Minimum permission level**|[Restricted](http://msdn.microsoft.com/library/da2efadc-4ebf-45fe-be39-397ac1eb1dbd%28Office.15%29.aspx)|
|**Add-in types**|Outlook|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history



****


|**Version**|**Changes**|
|:-----|:-----|
|1.0|Introduced|
