# SharePoint List Item model

This is an abstract wrapper class that makes it easy to create models based on SharePoint list items in TypeScript. It is made for use with the [SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/sharepoint-framework-overview) and uses the [SharePoint pnp-js](https://github.com/SharePoint/PnP-JS-Core).

## How to use

Import the dependencies:

```
import {SPListItemModel, SPList, SPField} from "sp-list-item-model"
```
Create a class that extends `SPListItemModel` and decorate it with `@SPList(ListName: string, SiteURL: string)`
`SiteURL` is optional and defaults to the site that is the current execution context.

The fields with the `@SPField(InternalName: string)` decorator represent fields in the list item. `InternalName` is optional and defaults to the class property name. If the [internal name](https://social.msdn.microsoft.com/Forums/office/en-US/75ca6fab-56f3-4bf4-aae0-2d29821778a2/how-to-get-internal-names-of-columns-in-sharepoint-lists?forum=sharepointdevelopmentlegacy) of the field is different from your property name, then specify it here.

```typescript
@SPList("employees", "https://mysite.sharepoint.com")
class Employee extends SPListItemModel {
    
    @SPField()
    Name: string
    
    @SPField()
    Address: string
    
    @SPField()
    Salary: number
    
    @SPField("aabbcc")
    HiredOn: Date
        
}
```

##### Creating an item

```
let e = new Employee();
e.Name = "John Doe";
e.Address = "123 Some Street";
e.Salary = 50000;
e.HiredOn = new Date();
e.submit()
    .then(()=>console.log("Item created"))
```

##### Retrieving information from the list
```
Employee.getItemById(123)
    .then(e=>console.log(e.Name))
```

##### Retrieving an item and updating its information
Note: the `submit` method checks if the fields you are updating have been changed since the item was loaded, and if so, will reject the promise. If you do not want this, then you can call `e.submit(false)` .
```
Employee.getItemById(123)
    .then(e=>{
        e.Name = "Joe Bloggs";
        return e.submit();
    })
    .then(()=>console.log("Updated"))
```

##### Updating an item without loading it fist
```
let e = new Employee();
e.ID=123;
e.Name = "Joe Bloggs";
e.submit()
    .then(()=>console.log("Updated"));
```
