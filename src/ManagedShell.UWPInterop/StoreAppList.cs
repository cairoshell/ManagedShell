using System.Collections.Generic;

namespace ManagedShell.UWPInterop
{
    public class StoreAppList : List<StoreApp>
    {
        public void FetchApps()
        {
            Clear();
            AddRange(StoreAppHelper.GetStoreApps());
        }

        public StoreApp GetAppByAumid(string appUserModelId)
        {
            // first attempt to get an app in our list already
            foreach (var storeApp in this)
            {
                if (storeApp.AppUserModelId == appUserModelId)
                {
                    return storeApp;
                }
            }

            // not in list, get from StoreAppHelper
            StoreApp app = StoreAppHelper.GetStoreApp(appUserModelId);

            if (app != null)
            {
                Add(app);
                return app;
            }

            // no app found for given AUMID
            return null;
        }
    }
}
