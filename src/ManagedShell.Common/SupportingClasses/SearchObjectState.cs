using System.Threading;

namespace ManagedShell.Common.SupportingClasses
{
    class SearchObjectState
    {
        public string SearchString;
        public ManualResetEvent Reset;

        public SearchObjectState(string searchStr)
        {
            SearchString = searchStr;
            Reset = new ManualResetEvent(false);
        }

        public SearchObjectState()
        {
            SearchString = string.Empty;
            Reset = new ManualResetEvent(false);
        }
    }
}
