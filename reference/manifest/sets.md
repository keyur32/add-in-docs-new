
# Sets element
Specifies the minimum subset of the JavaScript API for Office that your Office Add-in requires in order to activate.

 **Add-in type:** Content, Task pane, Mail


## Syntax:


```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```


## Contained in:

[Requirements](../../reference/manifest/requirements.md)


## Can contain:

[Set](../../reference/manifest/set.md)


## Attributes



|**Attribute**|**Type**|**Required**|**Description**|
|:-----|:-----|:-----|:-----|
|DefaultMinVersion|string|optional|Specifies the default  **MinVersion** attribute value for all child [Set](../../reference/manifest/set.md) elements. The default value is "1.1".|

## Remarks

For more information about requirement sets, see [Specify Office hosts and API requirements](http://msdn.microsoft.com/library/6b6702f2-b0a5-46ab-a356-8dda897ca8ae%28Office.15%29.aspx).

For more information about the  **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Specify Office hosts and API requirements](http://msdn.microsoft.com/library/6b6702f2-b0a5-46ab-a356-8dda897ca8ae%28Office.15%29.aspx#SpecifyRequirementSets_minversion).

