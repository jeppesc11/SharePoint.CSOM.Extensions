using Microsoft.SharePoint.Client;
using System.Diagnostics;

namespace SharePoint.CSOM.Extensions.Configuration
{
    /// <summary>
    /// Global configuration for CSOM operations retry logic
    /// </summary>
    public static class CSOMConfiguration
    {
        private static Func<ClientRuntimeContext, Action, Task>? _globalExecutor;

        /// <summary>
        /// Configure global CSOM execution logic
        /// </summary>
        /// <param name="executor">Function that takes a context and action, and executes it with retry/logging logic</param>
        public static void Configure(Func<ClientRuntimeContext, Action, Task> executor)
        {
            _globalExecutor = executor;
        }

        /// <summary>
        /// Clear global configuration - returns to default ExecuteQueryAsync behavior
        /// </summary>
        public static void Clear()
        {
            _globalExecutor = null;
        }

        /// <summary>
        /// Check if global configuration is set
        /// </summary>
        internal static bool HasGlobalConfiguration => _globalExecutor != null;

        /// <summary>
        /// Execute using the global configuration
        /// </summary>
        internal static async Task ExecuteWithGlobalConfiguration(ClientRuntimeContext context, Action setupAction)
        {
            if (_globalExecutor != null)
            {
                await _globalExecutor(context, setupAction);
            }
            else
            {
                setupAction(); // Execute the setup (Load, Update, etc.)
                await context.ExecuteQueryAsync(); // Then execute the query
            }
        }
    }
}