
# Design Office Add-ins for the iPad
Make your Office Add-ins available in Office for iPad.


Office Add-ins provide users with additional functionality within the context of an Office host. To make your Office Add-ins available in Office for iPad, apply the following guidelines:

- Make sure that your add-in meets the design guidelines below and complies with the requirements for add-ins that are available in the Office Store. For more information, see:

  - [Validation policies for apps and add-ins submitted to the Office Store (version 1.9)](http://msdn.microsoft.com/library/cd90836a-523e-42f5-ab02-5123cdf9fefe%28Office.15%29.aspx) (policies 3.4, 4.12, 10.8)

  - [Design guidelines for Office Add-ins](../../docs/design/add-in-design.md)

  - [Best practices for developing Office Add-ins](http://msdn.microsoft.com/library/d455b76b-4d76-493d-a681-6b02ba1f38a8%28Office.15%29.aspx)

  - [Designing for iOS](https://developer.apple.com/library/ios/documentation/UserExperience/Conceptual/MobileHIG/)

- Use version 1.1 of the JavaScript API for Office. For more information, see [Update the version of your JavaScript API for Office and manifest schema files](../../docs/overview/update-your-javascript-api-for-office-and-manifest-schema-version.md). To learn about changes in the API, see [What's changed in the JavaScript API for Office](../../reference/what's-changed-in-the-javascript-api-for-office.md).

- Make your Office Add-in free on the iPad. Only free Office Add-ins are allowed on the iPad. You can sell the same add-in on other platforms, such as Windows or Office 365. Why make a free Office Add-in? A free add-in allows you to reach more users.

- Do not include commerce in your add-in. This policy means that you cannot:

  - Offer trials.

  - Include UI elements or links to paid versions of the add-in.

  - Include UI elements or links to online stores that sell any additional content, apps, or add-ins.

  - Include UI elements or store links in your Privacy Policy or Terms of Use pages.


 >**Note**  Any violations to this policy will result in Microsoft immediately disabling your add-in on the iPad.

 You can share more information about the add-in or services within the add-in. You can also charge for and include commerce in your add-ins that run on other platforms. To do this, display different UIs -- depending on which browser or device runs your add-in -- by using the following properties:

  - [Context.touchEnabled property (JavaScript API for Office)](http://msdn.microsoft.com/library/fd73f94b-7e4a-422c-afdb-fef6fba43766%28Office.15%29.aspx) - Detects whether the host application your add-in runs on is touch enabled.

  - [Context.commerceAllowed property (JavaScript API for Office)](http://msdn.microsoft.com/library/fd3812ac-14c3-485f-8991-d12fcc99c450%28Office.15%29.aspx) - Determines whether your add-in runs on a platform that restricts commerce transactions.

When you've verified that your add-in meets these requirements, republish it to the Office Store. For more information, see [Submit Office and SharePoint Add-ins and Office 365 web apps to the Office Store](http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx) and [Update your app or add-in](http://msdn.microsoft.com/library/7313d32b-5345-4039-ac5d-a1ba0aef890b%28Office.15%29.aspx). In the Seller Dashboard, select the option to make your add-in available on the iPad, agree to the updated store policy, and provide your Apple Developer ID. Also, read and understand the [Office Store Application Provider Agreement](https://sellerdashboard.microsoft.com/Assets/Content/Agreements/en-US/Office_Store_Seller_Agreement_20120927.md).

## Additional resources



- [Debug Office Add-ins on iPad and Mac](../testing/debug-office-add-ins-on-ipad-and-mac.md)

