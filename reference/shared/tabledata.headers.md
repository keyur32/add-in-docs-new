
# TableData.headers property
Gets or sets the headers of the table.

|||
|:-----|:-----|
|**Hosts:**|Excel, Word|
|**Available in [Requirement set](http://msdn.microsoft.com/library/6b6702f2-b0a5-46ab-a356-8dda897ca8ae%28Office.15%29.aspx)**|TableBindings|
|**Last changed in**|1.1|

```
var hasHeaders = tableBindingObj.headers;
```


## Return Value

 **true** if the table has headers; otherwise **false**. 


## Remarks

To specify headers, you must specify an array of arrays that corresponds to the structure of the table. For example, to specify headers for a two-column table you would set the  **header** property to ` [['header1', 'header2']]`.

If you specify  **null** for the **headers** property (or leaving the property empty when you construct a **TableData** object), the following results occur when your code executes:


- If you insert a new table, the default column headers for the table are created.
    
- If you overwrite or update an existing table, the existing headers are not altered.
    

## Example

The following example creates a single-column table with a header and three rows.


```js
function createTableData() {
    var tableData = new Office.TableData();
    tableData.headers = [['header1']];
    tableData.rows = [['row1'], ['row2'], ['row3']];
    return tableData;
}

```



[![Try out this call in the interactive API Tutorial for Excel](../../images/819b84bf-151c-4a12-80c3-d6f8d7c03251.png)](http://officeapitutorial.azurewebsites.net/Redirect.html?scenario=Write+and+Read+a+Table)

## Support details


A capital Y in the following matrix indicates that this property is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this property.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](http://msdn.microsoft.com/library/67340567-bb9a-498c-96d3-3f52f28c16bc%28Office.15%29.aspx).

||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**Available in requirement sets**|TableBindings|
|**Minimum permission level**|[Restricted](http://msdn.microsoft.com/library/da2efadc-4ebf-45fe-be39-397ac1eb1dbd%28Office.15%29.aspx)|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history




|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Word Online.|
|1.1|Added support for Excel and Word in Office for iPad.|
|1.0|Introduced|
