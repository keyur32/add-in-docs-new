
# Table enumeration
Specifies enumerated values for the  `cells:` property in the _cellFormat_ parameter of [table formatting methods](http://msdn.microsoft.com/library/46b05707-b350-41be-b6b8-311799c71a33%28Office.15%29.aspx).

|||
|:-----|:-----|
|**Hosts:**|Excel|
|**Added**|1.1|

```
Office.Table
```

[![Try out this call in the interactive API Tutorial for Excel](../../images/819b84bf-151c-4a12-80c3-d6f8d7c03251.png)](https://officeapitutorial.azurewebsites.net/TryItOut.html)

## Members


**Values**


|**Enumeration**|**Value**|**Description**|
|:-----|:-----|:-----|
|Office.Table.All|"all"|The entire table, including column headers, data, and totals (if any).|
|Office.Table.Data|"data"|Only the data (no headers and totals).|
|Office.Table.Headers|"headers"|Only the header row.|

## Support details


A capital Y in following matrix indicates supported enumerations in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this enumeration.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](http://msdn.microsoft.com/library/67340567-bb9a-498c-96d3-3f52f28c16bc%28Office.15%29.aspx).


**Supported hosts, by platform**


||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|

|||
|:-----|:-----|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history




|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Excel in Office for iPad.|
|1.1|Introduced|
