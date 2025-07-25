using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;

namespace SharePoint.CSOM.Extensions.Extensions
{
    public static class ListItemExtensions
    {
        internal static ListItem PopulateFromDictionary(this ListItem item, Dictionary<string, object?> dic)
        {
            var caseInsensitiveDictionary = new Dictionary<string, object?>(dic, StringComparer.OrdinalIgnoreCase);

            if (caseInsensitiveDictionary.ContainsKey("id"))
            {
                caseInsensitiveDictionary.Remove("id");
            }

            item.FieldValues.Clear();

            foreach (var keyValuePair in caseInsensitiveDictionary)
            {
                item[keyValuePair.Key] = keyValuePair.Value;
            }

            return item;
        }

        public static bool FieldExistsAndNotNull(this ListItem item, string fieldName)
        {
            if (item == null || string.IsNullOrWhiteSpace(fieldName)) { return false; }
            if (!item.FieldExists(fieldName)) { return false; }

            var value = item[fieldName];
            if (value == null) { return false; }

            return value switch
            {
                FieldLookupValue lookup => lookup.LookupId > 0,
                TaxonomyFieldValue taxonomy => !string.IsNullOrEmpty(taxonomy.TermGuid),
                _ => true
            };
        }

        internal static bool FieldExists(this ListItem item, string fieldName)
        {
            if (item == null || string.IsNullOrWhiteSpace(fieldName)) { return false; }
            return item.FieldValues.ContainsKey(fieldName);
        }
    }
}
