
# Context.contentLanguage property
 Gets the locale (language) specified by the user for editing the document or item.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, PowerPoint, Project, Word|
|**Last changed in**|1.1|

```
var myContentLang = Office.context.contentLanguage;
```


## Return Value

A  **string** in the RFC 1766 Language tag format, such as `en-US`.


## Remarks

The  **contentLanguage** value reflects the **Editing Language** setting specified with **File** > **Options** > **Language** in the Office host application.

In content add-ins for Access web apps, the  **contentLanguage** property gets the add-in culture (e.g., "en-GB").


## Example




```js
function sayHelloWithContentLanguage() {
    var myContentLanguage = Office.context.contentLanguage;
    switch (myContentLanguage) {
        case 'en-US':
            write('Hello!');
            break;
        case 'en-NZ':
            write('G\'day mate!');
            break;
    }
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
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
|**Project**|Y|||
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**Minimum permission level**|[Restricted](http://msdn.microsoft.com/library/da2efadc-4ebf-45fe-be39-397ac1eb1dbd%28Office.15%29.aspx)|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history



****


|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for PowerPoint Online.|
|1.1|Added support for Excel, PowerPoint, and Word in Office for iPad.|
|1.1|Added access to this API in content add-ins for Access.|
|1.0|Introduced|
