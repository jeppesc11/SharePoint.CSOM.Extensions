using Microsoft.SharePoint.Client;
using SharePoint.CSOM.Extensions.Extensions;
using SharePoint.CSOM.Extensions.Models;

namespace TestApp.Temp
{

    public static class Position
    {
        public const string Ceo = "CEO";
        public const string Developer = "Developer";
        public const string Nobody = "Nobody";
    }

    public static class InternalName
    {
        public const string Id = "ID";
        public const string Title = "Title";
        public const string People = "People";
        public const string Position = "Position";
    }



    public class TestModel : ListItemModel<TestModel>
    {
        public string? Title { get; set; }
        public FieldUserValue? People { get; set; }
        public string? Position { get; set; }

        public override List<string> GetViewFields()
        {
            return new List<string> {
            InternalName.Title,
            InternalName.People,
            InternalName.Position
        };
        }

        public override TestModel FromListItem(ListItem item)
        {
            if (item == null) return this;
            this.Id = item.Id;
            Title = item.FieldExistsAndNotNull(InternalName.Title) ? (string)item[InternalName.Title] : null;
            People = item.FieldExistsAndNotNull(InternalName.People) ? (FieldUserValue)item[InternalName.People] : null;
            Position = item.FieldExistsAndNotNull(InternalName.Position) ? (string)item[InternalName.Position] : null;
            return this;
        }

        public override Dictionary<string, object?> ToDictionary()
        {
            return new Dictionary<string, object?> {
            { InternalName.Title, Title },
            { InternalName.People, People },
            { InternalName.Position, Position }
        };
        }
    }
}
