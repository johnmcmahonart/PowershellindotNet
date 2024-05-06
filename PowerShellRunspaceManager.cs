using System.Management.Automation.Runspaces;

namespace PowershellFromdotNet
{
    internal class PowerShellRunspaceManager
    {
        private Lazy<RunspacePool> _runspacePool;
        private InitialSessionState _initialSessionState;

        public RunspacePool PowerShellRunspacePool
        {
            get
            {
                // Initialize the Lazy instance if it hasn't been already
                if (_runspacePool == null)
                {
                    _runspacePool = new Lazy<RunspacePool>(InitializeRunspacePool, true);
                }
                return _runspacePool.Value;
            }
        }

        private RunspacePool InitializeRunspacePool()
        {
            //create empty pool that we can load modules into with PowerShellBootStrapper
            //_initialSessionState = InitialSessionState.Create();
            RunspacePool runspacePool = RunspaceFactory.CreateRunspacePool(1, 5);
            //RunspacePool runspacePool = RunspaceFactory.CreateRunspacePool();
            runspacePool.Open();
            return runspacePool;
        }

        public void Dispose()
        {
            PowerShellRunspacePool.Dispose();
        }
    }
}