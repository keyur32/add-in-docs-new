

# Settings.set method
Sets or creates the specified setting.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, PowerPoint, Word|
|**Available in [Requirement set](http://msdn.microsoft.com/library/6b6702f2-b0a5-46ab-a356-8dda897ca8ae%28Office.15%29.aspx)**|Settings|
|**Last changed in**|1.1|

[![Try out this call in the interactive API Tutorial for Excel](../../images/819b84bf-151c-4a12-80c3-d6f8d7c03251.png)](http://officeapitutorial.azurewebsites.net/Redirect.html?scenario=Persist+Settings&amp;amp;task=setSettings" target="_blank")


```js
Office.context.document.settings.set(name, value);
```


## Parameters



_name_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Type:  **string**

&nbsp;&nbsp;&nbsp;&nbsp;The case-sensitive name of the setting to set or create.

    
_value_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Type:  **string**,  **number**,  **boolean**,  **null**,  **object** or **array**

&nbsp;&nbsp;&nbsp;&nbsp;Specifies the value to be stored.
    

## Remarks

The  **set** method creates a new setting of the specified name if it does not already exist, or sets an existing setting of the specified name in the in-memory copy of the settings property bag. After you call the [Settings.saveAsync](../../reference/shared/settings.saveasync.md) method, the value is stored in the document as the serialized JSON representation of its data type. A maximum of 2MB is available for the settings of each add-in.


>**Important**
---
Be aware that the  **Settings.set** method affects only the in-memory copy of the settings property bag. To make sure that additions or changes to settings will be available to your add-in the next time the document is opened, at some point after calling the **Settings.set** method and before the add-in is closed, you must call the **Settings.saveAsync** method to persist settings in the document.


## Example




```js
function setMySetting() {
    Office.context.document.settings.set('mySetting', 'mySetting value');
}

```




## Support details


A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](http://msdn.microsoft.com/library/67340567-bb9a-498c-96d3-3f52f28c16bc%28Office.15%29.aspx).



||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**PowerPoint**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**Available in requirement sets**|Settings|
|**Minimum permission level**|[Restricted](http://msdn.microsoft.com/library/da2efadc-4ebf-45fe-be39-397ac1eb1dbd%28Office.15%29.aspx)|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history




|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for PowerPoint Online.|
|1.1|Added support for Excel, PowerPoint, and Word in Office for iPad.|
|1.1|Added support for custom settings in content add-ins for Access.|
|1.0|Introduced|
