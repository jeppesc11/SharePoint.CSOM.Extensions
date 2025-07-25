# SharePoint CSOM Extensions

A simple .NET 8 library that provides strongly-typed, async extensions for SharePoint CSOM operations with automatic batching and retry support.

## Quick Start

### 1. Create a Model
```csharp
public class Employee : ListItemModel<Employee>
{
    public string? Title { get; set; }
    public FieldUserValue? Manager { get; set; }
    public string? Department { get; set; }

    public override Employee FromListItem(ListItem item)
    {
        if (item == null) return this;
        Id = item.Id;
        Title = item.FieldExistsAndNotNull("Title") ? (string)item["Title"] : null;
        Manager = item.FieldExistsAndNotNull("Manager") ? (FieldUserValue)item["Manager"] : null;
        Department = item.FieldExistsAndNotNull("Department") ? (string)item["Department"] : null;
        return this;
    }

    public override Dictionary<string, object?> ToDictionary()
    {
        return new Dictionary<string, object?> {
            { "Title", Title },
            { "Manager", Manager },
            { "Department", Department }
        };
    }

    public override List<string> GetViewFields()
    {
        return new List<string> { "Title", "Manager", "Department" };
    }
}
```

### 2. Use the Extensions
```csharp
using var context = new ClientContext(siteUrl);
```

#### Read items from a list
```csharp
var items = await context.Web
                         .Lists
                         .GetById(new Guid(listId))
                         .GetItems<TestModel>(
                            CamlQuery.CreateAllItemsQuery()
                         );
```

#### Add a new item
```csharp
var newEmployee = new Employee { Title = "John Doe", Department = "HR" };
await context.Web
             .Lists
             .GetById(new Guid(listId))
             .AddItem(newEmployee);
```

#### Add multiple items
```csharp
var employees = new List<Employee>
{
    new Employee { Title = "Alice Smith", Department = "IT" },
    new Employee { Title = "Bob Brown", Department = "Marketing" }
};

await context.Web
             .Lists
             .GetById(new Guid(listId))
             .AddItems(employees);
```

#### Update an existing item
```csharp
var existingEmployee = await context.Web
                                     .Lists
                                     .GetById(new Guid(listId))
                                     .GetItemById<Employee>(itemId);

existingEmployee.Department = "Finance";

await context.Web
             .Lists
             .GetById(new Guid(listId))
             .UpdateItem(existingEmployee);
```

#### Update multiple items
```csharp
var employeesToUpdate = new List<Employee>
{
    new Employee { Id = 1, Title = "Jane Smith", Department = "IT" },
    new Employee { Id = 2, Title = "Bob Johnson", Department = "Marketing" }
};

await context.Web
             .Lists
             .GetById(new Guid(listId))
             .UpdateItems(employeesToUpdate);
```

## Features

- **Strongly-typed operations** - Type-safe CRUD with automatic field mapping
- **Automatic batching** - Bulk operations process in optimized chunks (100 items)
- **Smart change detection** - Only updates when actual changes exist
- **Retry support** - Optional global retry configuration with correlation tracking
- **Async/await** - Full async support throughout

## Batch Operations
All bulk operations automatically batch for performance

```csharp
var employees = new List<Employee> { /* hundreds of items */ };
await list.AddItems(employees);    // Processes in chunks of 100
await list.UpdateItems(employees); // Only updates changed items
```

## Optional: Global Retry Configuration
Configure once at startup for retry logic and correlation tracking
```csharp
CSOMConfiguration.Configure(async (context, action) =>
{
    // Add your retry logic, logging, correlation IDs, etc.
    await YourRetryHelper.ExecuteAsync(action);
});
```