using System.Collections.Generic;
using System.Threading.Tasks;

namespace MicrosoftGraphAspNetCoreConnectSample.Helpers
{
    public interface IIManageService
    {
        Task<List<ItemDetails>> GetRecentDocumentsAsync(string email);

        Task<string> GetRecentDocumentsJsonAsync(string email);
    }
}
