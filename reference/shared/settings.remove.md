

# Settings.remove method
Removes the specified setting.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, PowerPoint, Word|
|**Available in [Requirement set](http://msdn.microsoft.com/library/6b6702f2-b0a5-46ab-a356-8dda897ca8ae%28Office.15%29.aspx)**|Settings|
|**Last changed in**|1.1|

```js
Office.context.document.settings.remove(name);
```


## Parameters



_name_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Type:  **string**

&nbsp;&nbsp;&nbsp;&nbsp;The case-sensitive name of the setting to remove.

    



## Remarks

 **null** is a valid value for a setting. Therefore, assigning **null** to the setting will not remove it from the settings property bag.


>**Important**
---
Be aware that the  **Settings.remove** method affects only the in-memory copy of the settings property bag. To persist the removal of the specified setting in the document, at some point after calling the **Settings.remove** method and before the add-in is closed, you must call the [Settings.saveAsync](../../reference/shared/settings.saveasync.md) method.


## Example




```js
function removeMySetting() {
    Office.context.document.settings.remove('mySetting');
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
|**Word**|Y||Y|

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
|1.1|Added support to create custom settings in content add-ins for Access.|
|1.0|Introduced|
