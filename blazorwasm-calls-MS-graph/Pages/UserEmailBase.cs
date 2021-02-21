using Microsoft.AspNetCore.Components;
using Microsoft.Graph;
using System;
using System.Threading.Tasks;

namespace blazorwasm_calls_MS_graph.Pages
{
    /// <summary>
    /// Base class for UserProfile component.
    /// Injects GraphServiceClient and calls Microsoft Graph /me endpoint.
    /// </summary>
    public class UserEmailBase : ComponentBase
    {
        [Inject]
        GraphServiceClient GraphClient { get; set; }
        protected IUserMessagesCollectionPage _messages;

        protected override async Task OnInitializedAsync()
        {
            await GetMessages();
        }

        /// <summary>
        /// Retrieves uemail messages from Microsoft Graph /me endpoint.
        /// https://docs.microsoft.com/en-us/graph/api/user-list-messages?view=graph-rest-1.0&tabs=csharp
        /// </summary>
        /// <returns></returns>
        private async Task GetMessages()
        {
            try
            {
                _messages = await GraphClient.Me.Messages
                                .Request()
                                .GetAsync();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}

