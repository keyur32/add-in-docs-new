
# Get and set Outlook item data in read or compose forms
Learn about the ways to get or set data on an appointment or message from an Outlook add-in, taking into account whether the user is viewing or composing that item. You can use item-level properties in the JavaScript API for Office, Exchange Server callback tokens, or Exchange Web Services.



## Properties available in read and compose forms


Starting in version 1.1 of the Office Add-ins manifests schema, Outlook can activate add-ins when the user is viewing or composing an item. Depending on whether an add-in is activated in a read or compose form, the properties that are available to the add-in on the item differ as well. For example, the [dateTimeCreated](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html) and [dateTimeModified](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html) properties are defined only for an item that has already been sent (item is subsequently viewed in a read form) but not when the item is being created (in a compose form). Another example is the [bcc](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html) property which is only meaningful when a message is being authored (in a compose form), and is not accessible to the user in a read form.

Table 1 shows the item-level properties in the JavaScript API for Office that are available in each of read and compose modes of mail add-ins. Typically, those properties available in read forms are read-only, and those available in compose forms are read/write, with the exception of the [itemId](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html) and [conversationId](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html) properties, which are always read-only regardless. For the remaining item-level properties available in compose forms, because the add-in and user can possibly be reading or writing the same property at the same time, the methods to get or set them in compose mode are asynchronous, and hence the type of the objects returned by these properties are also different in compose forms than in read forms. For more information about using asynchronous methods to get or set item-level properties in compose mode, see [Get and set item data in a compose form in Outlook](../outlook/get-and-set-item-data-in-a-compose-form.md).


**Table 1. Item properties available in compose and read forms**


|**Item type**|**Property**|**Property type in read forms**|**Property type in compose forms**|
|:-----|:-----|:-----|:-----|
|Appointments and messages|[dateTimeCreated](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html)|JavaScript  **Date** object|Property not available|
|Appointments and messages|[dateTimeModified](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html)|JavaScript  **Date** object|Property not available|
|Appointments and messages|[itemClass](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html)|String|Property not available|
|Appointments and messages|[itemId](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html)|String|Property not available|
|Appointments and messages|[itemType](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html)|String in [ItemType](http://dev.outlook.com/reference/add-ins/Office.MailboxEnums.html) enumeration|Property not available|
|Appointments and messages|[attachments](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html)|[AttachmentDetails](http://dev.outlook.com/reference/add-ins/simple-types.html)|Property not available|
|Appointments and messages|[body](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html)|[Body](http://dev.outlook.com/reference/add-ins/Body.html)|[Body](http://dev.outlook.com/reference/add-ins/Body.html)|
|Appointments|[end](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html)|JavaScript  **Date** object|[Time](http://dev.outlook.com/reference/add-ins/Time.html)|
|Appointments|[location](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html)|String|[Location](http://dev.outlook.com/reference/add-ins/Location.html)|
|Appointments and messages|[normalizedSubject](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html)|String|Property not available|
|Appointments|[optionalAttendees](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html)|[EmailAddressDetails](http://dev.outlook.com/reference/add-ins/simple-types.html)|[Recipients](http://dev.outlook.com/reference/add-ins/Recipients.html)|
|Appointments|[organizer](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html)|**EmailAddressDetails**|Property not available|
|Appointments|[requiredAttendees](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html)|**EmailAddressDetails**|**Recipients**|
|Appointments|[resources](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html)|String|Property not available|
|Appointments|[start](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html)|JavaScript  **Date** object|**Time**|
|Appointments and messages|[subject](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html)|String|[Subject](http://dev.outlook.com/reference/add-ins/Subject.html)|
|Messages|[bcc](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html)|Property not available|**Recipients**|
|Messages|[cc](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html)|**EmailAddressDetails**|**Recipients**|
|Messages|[conversationId](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html)|String|String (read only)|
|Messages|[from](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html)|**EmailAddressDetails**|Property not available|
|Messages|[internetMessageId](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html)|Integer|Property not available|
|Messages|[sender](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html)|**EmailAddressDetails**|Property not available|
|Messages|[to](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html)|**EmailAddressDetails**|**Recipients**|

## Using Exchange Server callback tokens from a read add-in


If your Outlook add-in is activated in read forms, you can get an Exchange callback token. This token can be usedin server-side code to access the full item via Exchange Web Services (EWS). By specifying the  **ReadItem** permission in the add-in manifest, you can use the [mailbox.getCallbackTokenAsync](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.html) method to get an Exchange callback token, the [mailbox.ewsUrl](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.html) property to get the URL of the EWS endpoint for the user's mailbox, and [item.itemId](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html) to get the EWS ID for the selected item. You can then pass the callback token, EWS endpoint URL, and the EWS item ID to server-side code to access the [GetItem](http://msdn.microsoft.com/en-us/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx) operation, to get more properties of the item.


## Accessing Exchange Web Services from a read or compose add-in


You can also use the [mailbox.makeEwsRequestAsync](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.html) method to access the Exchange Web Services (EWS) operations [GetItem](http://msdn.microsoft.com/en-us/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx) and [UpdateItem](http://msdn.microsoft.com/en-us/library/5d027523-e0bc-4da2-b60b-0cb9fc1fdfe4%28Office.15%29.aspx) directly from the add-in. You can use these operations to get and set many properties of a specified item. This method is available to Outlook add-ins regardless of whether the add-in has been activated in a read or compose form, as long as you specify the **ReadWriteMailbox** permission in the add-in manifest. For more information in using **makeEwsRequestAsync** to access EWS operations, see [Call web services from an Outlook add-in](../outlook/web-services.md).


## Additional resources



- [Outlook add-ins](../outlook/outlook-add-ins.md)
    
- [Get and set item data in a compose form in Outlook](../outlook/get-and-set-item-data-in-a-compose-form.md)
    
- [Call web services from an Outlook add-in](../outlook/web-services.md)
    


