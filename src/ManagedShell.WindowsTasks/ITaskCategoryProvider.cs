using System;
using System.Threading.Tasks;

namespace ManagedShell.WindowsTasks
{
    public interface ITaskCategoryProvider : IDisposable
    {
        Task<string> GetCategoryAsync(ApplicationWindow window);

        void SetCategoryChangeDelegate(TaskCategoryChangeAsyncDelegate changeAsyncDelegate);
    }
}
