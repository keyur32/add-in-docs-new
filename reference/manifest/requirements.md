
# Requirements element
Specifies the minimum set of JavaScript API for Office requirements ([requirement sets](http://msdn.microsoft.com/library/6b6702f2-b0a5-46ab-a356-8dda897ca8ae%28Office.15%29.aspx#SpecifyRequirementSets_sets) and/or methods) that your Office Add-in needs to activate.

 **Add-in type:** Content, Task pane, Mail


## Syntax:


```XML
<Requirements>
   ...
</Requirements>
```


## Contained in:

[OfficeApp](../../reference/manifest/officeapp.md)


## Can contain:



|**Element**|**Content**|**Mail**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[Sets](../../reference/manifest/sets.md)|x|x|x|
|[Methods](../../reference/manifest/methods.md)|x||x|

## Remarks

For more information about requirement sets, see [Specify Office hosts and API requirements](http://msdn.microsoft.com/library/6b6702f2-b0a5-46ab-a356-8dda897ca8ae%28Office.15%29.aspx).

