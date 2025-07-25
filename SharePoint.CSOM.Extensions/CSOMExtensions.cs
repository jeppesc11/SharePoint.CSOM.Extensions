using Microsoft.SharePoint.Client;
using SharePoint.CSOM.Extensions.Extensions;
using SharePoint.CSOM.Extensions.Helpers;
using SharePoint.CSOM.Extensions.Models;
using SharePoint.CSOM.Extensions.Configuration;

namespace SharePoint.CSOM.Extensions
{
    public static class CSOMExtensions
    {
        #region CREATE

        public static async Task<T> AddItem<T>(this List list, T item) where T : ListItemModel<T>
        {
            var listItem = CreateListItem(list, item);
            listItem.Update();
            
            await ExecuteQuery(list.Context, () => { /* Setup already done above */ });
            
            return item.Load(listItem);
        }

        public static async Task<List<T>> AddItems<T>(this List list, List<T> items) where T : ListItemModel<T>
        {
            var listItems = await CSOMHelpers.ProcessInChunks(items, 100, async chunk => {
                var created = chunk.Select(item => {
                    var listItem = CreateListItem(list, item);
                    listItem.Update();
                    return listItem;
                }).ToList();

                await ExecuteQuery(list.Context, () => { /* Setup already done above */ });
                
                return created;
            });

            return items.Select((item, index) => item.Load(listItems[index])).ToList();
        }

        #region Private Methods

        private static ListItem CreateListItem<T>(List list, T item) where T : ListItemModel<T>
        {
            return list.AddItem(new ListItemCreationInformation()).PopulateFromDictionary(item.ToDictionary());
        }

        private static async Task ExecuteQuery(ClientRuntimeContext context, Action setupAction)
        {
            if (CSOMConfiguration.HasGlobalConfiguration)
            {
                await CSOMConfiguration.ExecuteWithGlobalConfiguration(context, setupAction);
            }
            else
            {
                setupAction();
                await context.ExecuteQueryAsync();
            }
        }

        #endregion Private Methods

        #endregion CREATE

        #region READ

        public static async Task<T?> GetItemById<T>(this List list, int id) where T : ListItemModel<T>, new()
        {
            var listItem = list.GetItemById(id);
            
            await ExecuteQuery(list.Context, () => {
                list.Context.Load(listItem);
            });
            
            return listItem != null ? new T().Load(listItem) : null;
        }

        public static async Task<List<T>> GetItems<T>(this List list, CamlQuery query) where T : ListItemModel<T>, new()
        {
            var result = new List<ListItem>();
            do
            {
                var items = list.GetItems(query);
                
                await ExecuteQuery(list.Context, () => {
                    list.Context.Load(items);
                });
                
                result.AddRange(items);
                query.ListItemCollectionPosition = items.ListItemCollectionPosition;
            } while (query.ListItemCollectionPosition != null);

            return result.Select(x => new T().Load(x)).ToList();
        }

        public static async Task<List<T>> GetItemsByIds<T>(this List list, IEnumerable<int> ids) where T : ListItemModel<T>, new()
        {
            if (ids == null || !ids.Any())
            {
                return new List<T>();
            }

            var listItems = await CSOMHelpers.ProcessInChunks(ids.ToList(), 100, async chunk => {
                var items = chunk.Select(id => {
                    var item = list.GetItemById(id);
                    list.Context.Load(item);
                    return item;
                }).ToList();

                await ExecuteQuery(list.Context, () => { /* Load calls already done above */ });
                
                return items;
            });

            return listItems.Select(x => new T()?.Load(x)).Where(x => x != null).Cast<T>().ToList();
        }

        #endregion READ

        #region UPDATE

        public static async Task<T> UpdateItem<T>(this List list, T item) where T : ListItemModel<T>
        {
            CSOMHelpers.ValidateItemId(item);
            var listItem = list.GetItemById(item.Id!.Value);
            
            await ExecuteQuery(list.Context, () => {
                list.Context.Load(listItem);
            });

            var updatedValues = item.ToDictionary();

            if (CSOMHelpers.HasChanges(listItem, updatedValues))
            {
                listItem.PopulateFromDictionary(updatedValues);
                listItem.Update();
                
                await ExecuteQuery(list.Context, () => { /* Update calls already done above */ });
            }

            return item.Load(listItem);
        }

        public static async Task<List<T>> UpdateItems<T>(this List list, List<T> items) where T : ListItemModel<T>
        {
            var listItems = await CSOMHelpers.ProcessInChunks(items, 100, async chunk => {
                var toUpdate = new List<ListItem>();

                foreach (var item in chunk)
                {
                    CSOMHelpers.ValidateItemId(item);

                    var listItem = item.SourceItem;
                    if (listItem == null)
                    {
                        throw new InvalidOperationException("Model must be initialized with a ListItem before updating.");
                    }

                    var updatedValues = item.ToDictionary();
                    if (CSOMHelpers.HasChanges(listItem, updatedValues))
                    {
                        listItem.PopulateFromDictionary(updatedValues);
                        listItem.Update();
                        toUpdate.Add(listItem);
                    }
                }

                if (toUpdate.Count > 0)
                {
                    await ExecuteQuery(list.Context, () => { /* Update calls already done above */ });
                }

                return toUpdate;
            });

            return items;
        }

        #endregion UPDATE

        #region DELETE

        public static async Task DeleteItem<T>(this List list, T item, bool permanentDelete = false) where T : ListItemModel<T>
        {
            CSOMHelpers.ValidateItemId(item);
            var listItem = list.GetItemById(item.Id!.Value)
                ?? throw new ArgumentException($"Item with ID {item.Id.Value} does not exist in the list.");

            await DeleteItemInternal(list.Context, listItem, permanentDelete);
        }

        public static async Task DeleteItemById(this List list, int itemId, bool permanentDelete = false)
        {
            var listItem = list.GetItemById(itemId)
                ?? throw new ArgumentException($"Item with ID {itemId} does not exist in the list.");

            await DeleteItemInternal(list.Context, listItem, permanentDelete);
        }

        public static async Task DeleteItemsByIds(this List list, IEnumerable<int> itemIds, bool permanentDelete = false)
        {
            if (itemIds == null || !itemIds.Any())
            {
                return;
            }

            await CSOMHelpers.ProcessInChunks(itemIds.ToList(), 100, async chunk => {
                var listItems = chunk.Select(id => {
                    var listItem = list.GetItemById(id);
                    if (listItem == null)
                    {
                        throw new ArgumentException($"Item with ID {id} does not exist in the list.");
                    }
                    if (permanentDelete)
                    {
                        listItem.DeleteObject();
                    }
                    else
                    {
                        listItem.Recycle();
                    }
                    return listItem;
                }).ToList();

                await ExecuteQuery(list.Context, () => { /* Delete calls already done above */ });
                
                return listItems;
            });
        }

        public static async Task DeleteItems<T>(this List list, List<T> items, bool permanentDelete = false) where T : ListItemModel<T>
        {
            await CSOMHelpers.ProcessInChunks(items, 100, async chunk => {
                var listItems = chunk.Select(item => {
                    CSOMHelpers.ValidateItemId(item);
                    var listItem = list.GetItemById(item.Id!.Value);
                    if (permanentDelete)
                    {
                        listItem.DeleteObject();
                    }
                    else
                    {
                        listItem.Recycle();
                    }
                    return listItem;
                }).ToList();

                await ExecuteQuery(list.Context, () => { /* Delete calls already done above */ });
                
                return listItems;
            });
        }

        #region Private Methods

        private static async Task DeleteItemInternal(ClientRuntimeContext context, ListItem listItem, bool permanentDelete)
        {
            if (listItem == null)
            {
                throw new ArgumentNullException(nameof(listItem), "ListItem cannot be null.");
            }

            if (permanentDelete)
            {
                listItem.DeleteObject();
            }
            else
            {
                listItem.Recycle();
            }

            await ExecuteQuery(context, () => { /* Delete calls already done above */ });
        }

        #endregion Private Methods

        #endregion DELETE
    }
}