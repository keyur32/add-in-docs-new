
# Document.customXmlParts property
Gets an object that represents the custom XML parts in the document.

|||
|:-----|:-----|
|**Hosts:**|Word|
|**Added in**|1.1|

```js
var xmlParts = Office.context.document.customXmlParts;
```


## Return Value

A [CustomXmlParts](../../reference/shared/customxmlparts.customxmlparts.md) object.


## Example




```js
function getCustomXmlParts(){
    Office.context.document.customXmlParts.getByNamespaceAsync('http://tempuri.org', function (asyncResult) {
        write('Retrieved ' + asyncResult.value.length + ' custom XML parts');
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## Support details


A capital Y in the following matrix indicates that this property is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this property.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](http://msdn.microsoft.com/library/67340567-bb9a-498c-96d3-3f52f28c16bc%28Office.15%29.aspx).


**Supported hosts, by platform**


||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**Minimum permission level**|[Restricted](http://msdn.microsoft.com/library/da2efadc-4ebf-45fe-be39-397ac1eb1dbd%28Office.15%29.aspx)|
|**Add-in types**|Task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history



****


|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Word on Office for iPad.|
|1.0|Introduced|
