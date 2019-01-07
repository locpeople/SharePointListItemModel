# SharePoint List Item model

This is an abstract class that makes it easy to create models based on SharePoint list items in TypeScript. It is made for use with the [SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/sharepoint-framework-overview) and Node.js and uses the [SharePoint pnp-js](https://github.com/SharePoint/PnP-JS-Core).

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
    
    //Add your own methods or properties
    NotASharePointField: string
    
    SomeMethod(): void {
        //do something
    } 
        
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

##### Retrieving a list item by its ID
```
Employee.getItemById(123)
    .then(e=>console.log(e.Name))
```

##### Retrieving all items from the list
```
Employee.getAllItems()
    .then(employees=>{
        employees.map(employee=>console.log(employee.Name))
    })
```

##### Retrieving items from the list using an OData filter
```
Employee.getItemsByFilter(`${Employee.getInternalName("Salary")} Ge 5000`)
    .then(employees=>{
        employees.map(employee=>console.log(employee.Name))
    })
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

##### Updating an item without loading it first
```
let e = new Employee();
e.ID=123;
e.Name = "Joe Bloggs";
e.submit()
    .then(()=>console.log("Updated"));
```

##### Deleting an item
```
Employee.deleteItemById(123)
    .then(()=>{
        //do something
    })
```


##### Authentication
If you are accessing SharePoint data from the Sharepoint Framework, then this will be done on behalf of the currently authenticated user, so authentication is not needed. If you want to access data from Node.JS you will need to authenticate. You can use [pnp-auth](https://www.npmjs.com/package/pnp-auth) for this (see the documentation for pnp-auth for instructions).
