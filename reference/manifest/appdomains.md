
# AppDomains element
Specifies any additional domains that your Office Add-in will use to load pages.

 **Add-in type:** Content, Task pane, Mail


## Syntax:


```XML
<AppDomains>
   ...
</AppDomains>
```


## Contained in:

 _ [OfficeApp](../../reference/manifest/officeapp.md)_


## Can contain:

[AppDomain](../../reference/manifest/appdomain.md)


## Remarks

The  **AppDomains** and **AppDomain** elements are used to specify any additional domains other than the one specified in the [SourceLocation](../../reference/manifest/sourcelocation.md) element. For more information, see [Office Add-ins XML manifest](http://msdn.microsoft.com/library/4139ff24-afac-472a-af7d-9d069587ac9b%28Office.15%29.aspx#bk_Preventing_Navigation).

