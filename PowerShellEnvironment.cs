using System.Management.Automation;

namespace PowershellFromdotNet
/// <summary>
/// Manages a PowerShell environment by creating and maintaining a runspace pool and facilitating module management and cmdlet execution.
/// </summary>
{
    public class PowershellEnvironment
    {
        private PowerShellRunspaceManager _runspaceManager;
        private PowerShellModuleManager _moduleManager;
        private PowerShellCmdlet? _cmdlet;
        private PowerShell _powerShellSession;

        /// <summary>
        /// Initializes a new instance of the PowershellEnvironment class.
        /// </summary>
        /// <remarks>
        /// This constructor sets up a new PowerShell environment with a runspace pool and a session manager,
        /// ready to load modules and execute cmdlets.
        /// </remarks>
        public PowershellEnvironment()
        {
            // Wire up environment for use by
            // 1. Create Powershell RunspacePool
            // 2. Create Powershell instance and associate it to the Runspace pool

            _runspaceManager = new PowerShellRunspaceManager();
            _powerShellSession = PowerShell.Create();
            _powerShellSession.RunspacePool = _runspaceManager.PowerShellRunspacePool;
            _moduleManager = new PowerShellModuleManager(_powerShellSession);
        }

        /// <summary>
        /// Installs the specified modules into the PowerShell session.
        /// </summary>
        /// <param name="modulesToInstall">A list of module names to install.</param>
        public void InstallModule(List<string> modulesToInstall)
        {
            _moduleManager.InstallModule(modulesToInstall);
        }

        /// <summary>
        /// Imports the specified modules into the PowerShell session.
        /// </summary>
        /// <param name="modulesToImport">A list of module names to import.</param>
        public void ImportModule(List<string> modulesToImport)
        {
            _moduleManager.ImportModule(modulesToImport);
        }

        /// <summary>
        /// Creates a new <see cref="PowerShellCmdlet"/> instance that can be used to execute cmdlets dynamically.
        /// </summary>
        /// <returns>A new instance of <see cref="PowerShellCmdlet"/>.</returns>
        public PowerShellCmdlet Cmdlet()
        {
            return new PowerShellCmdlet(_powerShellSession);
        }

        /// <summary>
        /// Closes the PowerShell environment and releases all associated resources.
        /// </summary>
        public void Close()
        {
            _powerShellSession.Dispose();
            _runspaceManager.Dispose();
        }
    }
}