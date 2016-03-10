

# event.completed
The callback that the add-in invokes to let Outlook know that the operation is done.

****

|||
|:-----|:-----|
|**Hosts:**Outlook|**Add-in type:** Outlook|
|**Available in [requirement sets](http://msdn.microsoft.com/library/6b6702f2-b0a5-46ab-a356-8dda897ca8ae%28Office.15%29.aspx)**|Mailbox|
|**Last changed in Mailbox**|1.3|
|**Applicable Outlook modes**|Read and Compose|



```js
event.completed();
```


## Parameters

None


## Support details


A capital Y in the following table indicates that this property is supported in the corresponding Outlook host application. An empty cell indicates that the Outlook host application doesn't support this property.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](http://msdn.microsoft.com/library/67340567-bb9a-498c-96d3-3f52f28c16bc%28Office.15%29.aspx).

 **Important:** Add-in commands and the APIs associated with them currently work only in Outlook in [Office 2016 Preview](https://products.office.com/en-us/office-2016-preview) on Windows desktop.


**Supported hosts, by platform**


| |**Office for Windows desktop**|**Office Online(in browser)**|**OWA for Devices**|
|:-----|:-----|:-----|:-----|
|**Outlook**|Y|||

|||
|:-----|:-----|
|**Available in requirement sets**|Mailbox|
|**Minimum permission level**|[ReadWriteItem](http://msdn.microsoft.com/library/5bca69f2-b287-4e19-8f0f-78d896b2a3d3%28Office.15%29.aspx)|
|**Add-in types**|Outlook|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history




|**Version**|**Changes**|
|:-----|:-----|
|1.3|Introduced|
