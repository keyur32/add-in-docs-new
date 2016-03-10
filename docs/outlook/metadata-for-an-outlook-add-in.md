
# Get and set add-in metadata for an Outlook add-in
Learn to load and store custom data that only your Outlook add-in can access - through roaming settings and custom properties.


There are two ways to manage custom data in your Outlook add-in:

- Roaming settings, which manages custom data for a user's mailbox.
    
- Custom properties, which manages custom data for an item in a user's mailbox.
    
Both of these give access to custom data that is only accessible by your Outlook add-in, but each method stores the data separately from the other. That is, the data stored through roaming settings is not accessible by custom properties, and vice versa. The data is stored on the server for that mailbox, and is accessible in subsequent Outlook sessions on all the form factors that the add-in supports. 

## Custom data per mailbox: roaming settings


You can specify data specific to a user's Exchange mailbox using the [RoamingSettings](https://dev.outlook.com/reference/add-ins/RoamingSettings.html) object. Examples of such data include the user's personal data and preferences. Your mail add-in can access roaming settings when it roams on any device it's designed to run on (desktop, tablet, or smartphone).

 Changes to this data are stored on an in-memory copy of those settings for the current Outlook session. You should explicitly save all the roaming settings after updating them so that they will be available the next time the user opens your add-in, on the same or any other supported device.


### Roaming settings format


The data in a  **RoamingSettings** object is stored as a serialized JavaScript Object Notation (JSON) string. The following is an example of the structure, assuming there are three defined roaming settings named `add-in_setting_name_0`,  `add-in_setting_name_1`, and  `add-in_setting_name_2`.


```js
{
  "add-in_setting_name_0":"add-in_setting_value_0",
  "add-in_setting_name_1":"add-in_setting_value_1",
  "add-in_setting_name_2":"add-in_setting_value_2"
}
```


### Loading roaming settings


A mail add-in typically loads roaming settings in the [Office.initialize](http://msdn.microsoft.com/en-us/library/727adf79-a0b5-48d2-99c7-6642c2c334fc%28Office.15%29.aspx) event handler. The following JavaScript code example shows how to load existing roaming settings and get the values of 2 settings, "customerName" and "customerBalance":


```js
var _mailbox;
var _settings;
var _customerName;
var _customerBalance;

// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Initialize instance variables to access API objects.
  _mailbox = Office.context.mailbox;
  _settings = Office.context.roamingSettings;
  _customerName = _settings.get("customerName");
  _customerBalance = _settings.get("customerBalance");
}

```


### Creating or assigning a roaming setting


Continuing with the preceding example, the following JavaScript function,  `setAddInSetting`, shows how to use the [RoamingSettings.set](https://dev.outlook.com/reference/add-ins/RoamingSettings.html) method to set a setting named `cookie` with today's date, and persist the data by using the [RoamingSettings.saveAsync](https://dev.outlook.com/reference/add-ins/RoamingSettings.html) method to save all the roaming settings back to the server. The **set** method creates the setting if the setting does not already exist, and assigns the setting to the specified value. The **saveAsync** method saves roaming settings asynchronously. This code sample passes a callback method, `saveMyAddInSettingsCallback`, to  **saveAsync**. When the asynchronous call finishes,  `saveMyAddInSettingsCallback` is called by using one parameter, _asyncResult_. This parameter is an [AsyncResult](https://dev.outlook.com/reference/add-ins/simple-types.html) object that contains the result of and any details about the asynchronous call. You can use the optional _userContext_ parameter to pass any state information from the asynchronous call to the callback function.


```js
// Set a roaming setting.
function setAddInSetting() {
  _settings.set("cookie", Date());
  // Save roaming settings for the mailbox
  // to the server so that they will be available
  // in the next session.
  _settings.saveAsync(saveMyAddInSettingsCallback);
}

// Callback method after saving custom roaming settings.
function saveMyAddInSettingsCallback(asyncResult) {
  if (asyncResult.status == Office.AsyncResultStatus.Failed) {
    // Handle the failure.
  }
}
```


### Removing a roaming setting


Also extending the preceding examples, the following JavaScript function,  `removeAddInSetting`, shows how to use the [RoamingSettings.remove](https://dev.outlook.com/reference/add-ins/RoamingSettings.html) method to remove the `cookie` setting and save all the roaming settings back to the Exchange Server.


```js
// Remove an add-in setting.
function removeAddInSetting()
{
  _settings.remove("cookie");
  // Save changes to the roaming settings for the mailbox
  // to the server so that they will be available
  // in the next session.
  _settings.saveAsync(saveMyAddInSettingsCallback);
}
```


## Custom data per item in a mailbox: custom properties


You can specify data specific to an item in the user's mailbox using the [CustomProperties](https://dev.outlook.com/reference/add-ins/CustomProperties.html) object. For example, your mail add-in could categorize certain messages and note the category using a custom property `messageCategory`. Or, if your mail add-in creates appointments from meeting suggestions in a message, you can use a custom property to track each of these appointments. This ensures that if the user opens the message again, your mail add-in does not offer to create the appointment a second time.

Similar to roaming settings, changes to custom properties are stored on in-memory copies of the properties for the current Outlook session. To make sure these custom properties will be available in the next session, save all the custom properties to the server.

These add-in-specific, item-specific custom properties can only be accessed by using the  **CustomProperties** object. These properties are different from the custom, MAPI-based, [UserProperties](http://msdn.microsoft.com/library/20b49c86-d74f-9bda-382c-559af278c148%28Office.15%29.aspx) in the Outlook object model, and extended properties in Exchange Web Services (EWS). You cannot access **CustomProperties** by using the Outlook object model or EWS.

However, a mail add-in can get MAPI-based extended properties by using the EWS [GetItem](http://msdn.microsoft.com/en-us/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx) operation. Access **GetItem** on the server side by using a callback token, or on the client side by using the [mailbox.makeEwsRequestAsync](https://dev.outlook.com/reference/add-ins/Office.context.mailbox.html) method. In the **GetItem** request, specify the custom extended properties you need in a property set. A mail add-in can also use **makeEwsRequestAsync** and EWS [CreateItem](http://msdn.microsoft.com/library/78a52120-f1d0-4ed7-8748-436e554f75b6%28Office.15%29.aspx) and [UpdateItem](http://msdn.microsoft.com/library/5d027523-e0bc-4da2-b60b-0cb9fc1fdfe4%28Office.15%29.aspx) operations to create and modify extended properties.




### Using custom properties


Before you can use custom properties, you must load them by calling the [loadCustomPropertiesAsync](https://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html) method. If any custom properties are already set for the current item, they are loaded from the Exchanger server at this point. After you have created the property bag, you can use the [set](https://dev.outlook.com/reference/add-ins/CustomProperties.html) and [get](https://dev.outlook.com/reference/add-ins/CustomProperties.html) methods to add and retrieve custom properties. To save any changes that you make to the property bag, you must use the [saveAsync](https://dev.outlook.com/reference/add-ins/CustomProperties.html) method to persist the changes on the Exchange server.


 >**Note**  Because Outlook for Mac doesn't cache custom properties, if the user's network goes down, mail add-ins in Outlook for Mac would not be able to access their custom properties.


### Custom properties example


The following example shows a simplified set of methods for an Outlook add-in that uses custom properties. You can use this example as a starting point for your add-in that uses custom properties. 

This example includes the following methods:


- [Office.initialize](http://msdn.microsoft.com/en-us/library/727adf79-a0b5-48d2-99c7-6642c2c334fc%28Office.15%29.aspx) -- Initializes the add-in and loads the custom property bag from the Exchange server.
    
-  **customPropsCallback** -- Gets the custom property bag that is returned from the server and saves it for later use.
    
-  **updateProperty** -- Sets or updates a specific property, and then saves the change to the server.
    
-  **removeProperty** -- Removes a specific property from the property bag, and then saves the removal to the server.
    



```js
var _mailbox;
var _customProps;

// The initialize function is required for all add-ins.
Office.initialize = function () {
  _mailbox = Office.context.mailbox;
  _mailbox.item.loadCustomPropertiesAsync(customPropsCallback);
}

// Callback function from loading custom properties.
function customPropsCallback(asyncResult) {
  if (asyncResult.status == Office.AsyncResultStatus.Failed) {
    // Handle the failure.
  }
  else {
    // Successfully loaded custom properties,
    // can get them from the asyncResult argument.
    _customProps = asyncResult.value;
  }
}

// Get individual custom property.
function getProperty() {
  var myProp = customProps.get("myProp");
}

// Set individual custom property.
function updateProperty(name, value) {
  _customProps.set(name, value);
  // Save all custom properties to server.
  _customProps.saveAsync(saveCallback);
}

// Remove a custom property.
function removeProperty(name) {
  _customProps.remove(name);
  // Save all custom properties to server.
  _customProps.saveAsync(saveCallback);
}

// Callback function from saving custom properties.
function saveCallback() {
  if (asyncResult.status == Office.AsyncResultStatus.Failed) {
    // Handle the failure.
  }
}
```


## Additional resources



- [Outlook add-ins](../outlook/outlook-add-ins.md)
    
- [Overview of Outlook add-ins architecture and features](../outlook/overview.md)
    
- [MAPI Property Overview](http://msdn.microsoft.com/library/02e5b23f-1bdb-4fbf-a27d-e3301a359573%28Office.15%29.aspx)
    
- [Outlook Properties Overview](http://msdn.microsoft.com/library/242c9e89-a0c5-ff89-0d2a-410bd42a3461%28Office.15%29.aspx)
    
- [Call web services from an Outlook add-in](../outlook/web-services.md)
    
- [Properties and extended properties in EWS in Exchange](http://msdn.microsoft.com/library/68623048-060e-4602-b3fa-62617a94cf72%28Office.15%29.aspx)
    
- [Property sets and response shapes in EWS in Exchange](http://msdn.microsoft.com/library/04a29804-6067-48e7-9f5c-534e253a230e%28Office.15%29.aspx)
    


