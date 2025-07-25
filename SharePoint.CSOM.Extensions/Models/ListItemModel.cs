using Microsoft.SharePoint.Client;
using System.Text.Json.Serialization;

namespace SharePoint.CSOM.Extensions.Models
{
    public abstract class ListItemModel<T>
    {
        [JsonIgnore]
        public ListItem? SourceItem { get; private set; }

        public int? Id { get; set; }
        public string? Title { get; set; }
        public string? ContentType { get; set; }

        public abstract T FromListItem(ListItem item);
        public abstract Dictionary<string, object?> ToDictionary();
        public abstract List<string> GetViewFields();

        public T Load(ListItem item)
        {
            SourceItem = item;
            return FromListItem(item);
        }
    }
}
