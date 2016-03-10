
# Permissions element
Specifies the level of API access for your Office Add-in; you should request permissions based on the principle of least privilege.

 **Add-in type:** Content, Task pane, Mail


## Syntax:

For content and task pane add-ins:


```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

For mail add-ins:




```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```


## Contained in:

 _[OfficeApp](../../reference/manifest/officeapp.md)_


## Remarks

For more detail, see [Requesting permissions for API use in content and task pane add-ins](http://msdn.microsoft.com/library/da2efadc-4ebf-45fe-be39-397ac1eb1dbd%28Office.15%29.aspx) and [Understanding Outlook add-in permissions](http://msdn.microsoft.com/library/5bca69f2-b287-4e19-8f0f-78d896b2a3d3%28Office.15%29.aspx).

