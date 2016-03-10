
# officeTheme.controlBackgroundColor property
Gets the Office theme control background color.

 **Important:** This API currently works only in Excel, Outlook, PowerPoint, and Word in [Office 2016 Preview](https://products.office.com/en-us/office-2016-preview) on Windows desktop.



|||
|:-----|:-----|
|**Hosts:**|Excel, Outlook, PowerPoint, Word|
|**Available in [Requirement set](http://msdn.microsoft.com/library/6b6702f2-b0a5-46ab-a356-8dda897ca8ae%28Office.15%29.aspx)**|Not in a set|
|**Added in**|1.3|

```
var controlBackgroundColor = Office.context.officeTheme.controlBackgroundColor;
```


## Return value

A hex color triplet.


## Remarks

The colors returned correspond to the values of the Office theme selected by the user with  **File** > **Office Account** > **Office Theme** UI, which is applied across all Office host applications.


## Support details


A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](http://msdn.microsoft.com/library/67340567-bb9a-498c-96d3-3f52f28c16bc%28Office.15%29.aspx).


**Supported hosts, by platform**


||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|**OWA for Devices**|
|:-----|:-----|:-----|:-----|:-----|
|**Excel**|Y||||
|**Outlook**|Y||||
|**PowerPoint**|Y||||
|**Word**|Y||||

|||
|:-----|:-----|
|**Minimum permission level**|[Restricted](http://msdn.microsoft.com/library/da2efadc-4ebf-45fe-be39-397ac1eb1dbd%28Office.15%29.aspx)|
|**Add-in types**|Content, task pane, Outlook|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history



****


|**Version**|**Changes**|
|:-----|:-----|
|1.3|Introduced|
