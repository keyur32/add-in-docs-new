
# What's changed in the JavaScript API for Office
The JavaScript API for Office is periodically updated with new and updated objects, methods, properties, events and enumerations to extend the functionality of your Office Add-ins. Use the links below to see the new and updated API members.

 _**Applies to:** Access apps for SharePoint | apps for Office | Excel | Office Add-ins | Outlook | PowerPoint | Project | Word_

To develop add-ins using new API members, you need to [update the JavaScript API for Office files in your project](../../docs/overview/update-your-javascript-api-for-office-and-manifest-schema-version.md).

To view all API members including those that are unchanged from previous updates, see [JavaScript API for Office](http://msdn.microsoft.com/library/b27e70c3-d87d-4d27-85e0-103996273298%28Office.15%29.aspx).


## New and updated API

 **New and updated objects**



|**Object**|**Description**|**Version added or updated **|
|:-----|:-----|:-----|
|[Item](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html)|Updates and additions to:<br><ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>The <a href="http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html(Office.15).aspx#getSelectedDataAsync" target="_blank">getSelectedDataAsync</a> and <a href="http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html(Office.15).aspx#setSelectedDataAsync" target="_blank">setSelectedDataAsync</a> methods to support getting the user's selection and overwriting it in the subject and body  of a message or appointment.</p></li><li><p>The <a href="http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html(Office.15).aspx#displayReplyAllForm" target="_blank">displayReplyAllForm</a> and <a href="http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html(Office.15).aspx#displayReplyForm" target="_blank">displayReplyForm</a> methods to support adding an attachment to the reply form of an appointment.</p></li></ul>|Mailbox 1.2|
|[Item](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html)|Updated to include methods and fields for creating compose mode Outlook add-ins. |1.1|
|[Binding](http://msdn.microsoft.com/library/42882642-d22b-47d2-a8d3-3aa8c6a4435e%28Office.15%29.aspx)|Updated to support table binding in content add-ins for Access.|1.1|
|[Bindings](http://msdn.microsoft.com/library/09979e31-3bfb-45be-adda-0f7cc2db1fe1%28Office.15%29.aspx)|Updated to support table binding in content add-ins for Access.|1.1|
|[Body](http://dev.outlook.com/reference/add-ins/Body.html)|Added to enable creating and editing the body of a message or appointment in compose mode Outlook add-ins.|1.1|
|[Document](http://msdn.microsoft.com/library/f8859516-cc1f-4b20-a8f3-cee37a983e70%28Office.15%29.aspx)|Updates and additions to: <ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Support <a href="http://msdn.microsoft.com/library/551369c3-315b-428f-8b7e-08987f6b0e00(Office.15).aspx" target="_blank">mode</a>, <a href="http://msdn.microsoft.com/library/77ba7daf-419f-44b6-8747-7fd5618b7053(Office.15).aspx" target="_blank">settings</a>, and <a href="http://msdn.microsoft.com/library/480ac3c6-370e-4505-aba3-1d0dce9fb3dc(Office.15).aspx" target="_blank">url</a> properties in content add-ins for Access.</p></li><li><p>Get the document as PDF with the <a href="http://msdn.microsoft.com/library/35dda81c-235e-4eab-8a77-9acb3b73a380(Office.15).aspx" target="_blank">getFileAsync</a> method in add-ins for PowerPoint and Word.</p></li><li><p>Get file properties with the <a href="http://msdn.microsoft.com/library/2533a563-95ae-4d52-b2d5-a6783e4ef5b4(Office.15).aspx" target="_blank">getFileProperties</a> method in add-ins for Excel, PowerPoint, and Word.</p></li><li><p>Navigate to locations and objects within the document with the <a href="http://msdn.microsoft.com/library/35dda81c-235e-4eab-8a77-9acb3b73a380(Office.15).aspx" target="_blank">goToByIdAsync</a> method in add-ins for Excel and PowerPoint.</p></li><li><p>Get the id, title, and index for selected slides with the <a href="http://msdn.microsoft.com/library/f85ad02c-64f0-4b73-87f6-7f521b3afd69(Office.15).aspx" target="_blank">getSelectedDataAsync</a> method (when you specify the new <span class="keyword">Office.CoercionType.SlideRange</span><a href="http://msdn.microsoft.com/library/735eaab6-5e31-4bc2-add5-9d378900a31b(Office.15).aspx" target="_blank">coercionType</a> enum) in add-ins for PowerPoint.</p></li></ul>|1.1|
|[Location](http://dev.outlook.com/reference/add-ins/Location.html)|Added to enable setting the location of an appointment in compose mode Outlook add-ins.|1.1|
|[Office](http://msdn.microsoft.com/library/c490b13d-ee52-4291-af5d-f4a5a11d3af0%28Office.15%29.aspx)|Updated the select method to support getting bindings in content add-ins for Access.|1.1|
|[Recipients](http://dev.outlook.com/reference/add-ins/Recipients.html)|Added to enable getting and setting the recipients of a message or appointment in compose mode.|1.1|
|[Settings](http://msdn.microsoft.com/library/ad733387-a58c-4514-8fc2-53e64fad468d%28Office.15%29.aspx)|Updated to support creating custom settings in content add-ins for Access.|1.1|
|[Subject](http://dev.outlook.com/reference/add-ins/Subject.html)|Added to enable getting and setting the subject of a message or appointment in compose mode Outlook add-ins.|1.1|
|[Time](http://dev.outlook.com/reference/add-ins/Time.html)|Added to enable getting and setting the start and end time of an appointment in compose mode Outlook add-ins.|1.1|

**New and updated enumerations**


|**Object**|**Description**|**Version**|
|:-----|:-----|:-----|
|[ActiveView](http://msdn.microsoft.com/library/1f1d963e-04e1-4cf2-b161-5329d7ad0a3e%28Office.15%29.aspx)|Specifies the state of the active view of the document, for example, whether the user can edit the document.Added so that add-ins for PowerPoint can determine if the users is viewing the presentation ( **Slide Show**) or editing slides. |1.1|
|[CoercionType](http://msdn.microsoft.com/library/735eaab6-5e31-4bc2-add5-9d378900a31b%28Office.15%29.aspx)|Updated with  **Office.CoercionType.SlideRange** to support getting the selected slide range with the **getSelectedDataAsync** method in add-ins for PowerPoint.|1.1|
|[EventType](http://msdn.microsoft.com/library/82c79659-52da-48b0-92a9-831226eb9a7f%28Office.15%29.aspx)|Updated to include the new ActiveViewChanged event.|1.1|
|[FileType](http://msdn.microsoft.com/library/fadbb4cf-a0e4-47b2-93dd-123f0b06d4ae%28Office.15%29.aspx)|Updated to specify output in PDF format.|1.1|
|[GoToType](http://msdn.microsoft.com/library/8de45be3-de35-4765-a67a-e128a46786bd%28Office.15%29.aspx)|Added to specify the place or object in the document to go to.|1.1|

## Additional resources


- [Office Add-ins API and schema references](../reference/reference.md)
    
- [Office Add-ins](../../docs/overview/office-add-ins.md)
    
