using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using SharePoint.CSOM.Extensions.Models;

namespace SharePoint.CSOM.Extensions.Helpers
{
    internal static partial class CSOMHelpers
    {

        internal static async Task<List<TOut>> ProcessInChunks<TIn, TOut>(List<TIn> source, int chunkSize, Func<List<TIn>, Task<List<TOut>>> action)
        {
            var result = new List<TOut>();
            foreach (var chunk in source.Chunk(chunkSize))
            {
                result.AddRange(await action(chunk.ToList()));
            }
            return result;
        }

        internal static void ValidateItemId<T>(T item) where T : ListItemModel<T>
        {
            if (!item.Id.HasValue)
            {
                throw new ArgumentException("Item ID must be set before operation.");
            }
        }

        internal static bool HasChanges(ListItem listItem, Dictionary<string, object?> updatedValues)
        {
            foreach (var kvp in updatedValues)
            {
                if (!listItem.FieldValues.ContainsKey(kvp.Key))
                {
                    continue;
                }

                var existingValue = listItem[kvp.Key];
                var updatedValue = kvp.Value;

                if (!AreValuesEqual(existingValue, updatedValue))
                {
                    return true;
                }
            }
            return false;
        }

        private static bool AreValuesEqual(object? oldValue, object? newValue)
        {
            if (oldValue == null && newValue == null)
            {
                return true;
            }
            if (oldValue == null || newValue == null)
            {
                return false;
            }

            if (oldValue is FieldUserValue oldUser && newValue is FieldUserValue newUser)
            {
                return oldUser.LookupId == newUser.LookupId;
            }

            if (oldValue is FieldUserValue[] oldUsers && newValue is FieldUserValue[] newUsers)
            {
                return oldUsers.Length == newUsers.Length &&
                       oldUsers.Zip(newUsers, (o, n) => o.LookupId == n.LookupId).All(equal => equal);
            }

            if (oldValue is FieldLookupValue oldLookup && newValue is FieldLookupValue newLookup)
            {
                return oldLookup.LookupId == newLookup.LookupId;
            }

            if (oldValue is FieldLookupValue[] oldLookups && newValue is FieldLookupValue[] newLookups)
            {
                return oldLookups.Length == newLookups.Length &&
                       oldLookups.Zip(newLookups, (o, n) => o.LookupId == n.LookupId).All(equal => equal);
            }

            if (oldValue is FieldUrlValue oldUrl && newValue is FieldUrlValue newUrl)
            {
                return oldUrl.Url == newUrl.Url && oldUrl.Description == newUrl.Description;
            }

            if (oldValue is DateTime oldDate && newValue is DateTime newDate)
            {
                return oldDate.ToUniversalTime().Equals(newDate.ToUniversalTime());
            }

            if (oldValue is string[] oldArray && newValue is string[] newArray)
            {
                return oldArray.SequenceEqual(newArray);
            }

            if (oldValue is TaxonomyFieldValue oldTax && newValue is TaxonomyFieldValue newTax)
            {
                return oldTax.TermGuid == newTax.TermGuid;
            }

            if (oldValue is TaxonomyFieldValueCollection oldTaxSet && newValue is TaxonomyFieldValueCollection newTaxSet)
            {
                var oldGuids = oldTaxSet.Select(t => t.TermGuid).OrderBy(g => g);
                var newGuids = newTaxSet.Select(t => t.TermGuid).OrderBy(g => g);
                return oldGuids.SequenceEqual(newGuids);
            }

            return Equals(oldValue, newValue);
        }
    }
}
