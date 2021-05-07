using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace ManagedShell.WindowsTasks
{
    public interface ITaskCategoryProvider : IDisposable
    {
        string GetCategory(ApplicationWindow window);

        void SetCategoryChangeDelegate(TaskCategoryChangeDelegate changeDelegate);
    }
}
