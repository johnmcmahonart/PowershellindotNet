using System.Management.Automation;

namespace PowershellFromdotNet
{
    internal class PowerShellModuleManager
    {
        private PowerShell _powerShellInstance;

        public PowerShellModuleManager(PowerShell PowerShellInstance)
        {
            _powerShellInstance = PowerShellInstance;
        }

        /// <summary>
        /// Loads a list of PowerShell modules into the shared PowerShell session.
        /// </summary>
        /// <param name="modulesToLoad">A list of module names to load. Each module name must correspond to a module
        /// that is available in the session's module path.</param>
        /// <exception cref="System.Management.Automation.CommandNotFoundException">
        /// Thrown if any specified module does not exist or is not discoverable in the session's module path.
        /// </exception>
        /// <exception cref="System.Management.Automation.RuntimeException">
        /// Thrown if an error occurs during the loading of a module, such as conflicts or errors within the module scripts.
        /// </exception>
        /// <remarks>
        /// Ensure that all modules listed in <paramref name="modulesToLoad"/> are installed and accessible before calling this method.
        /// This method may affect application performance if large modules are loaded, due to initialization overhead.
        /// </remarks>
        /// <example>
        /// Here is how you can use the LoadModule method:
        /// <code>
        /// try
        /// {
        ///     List<string> modules = new List<string> { "ActiveDirectory", "SqlServer" };
        ///     PowerShellBootstrapper.LoadModule(modules);
        /// }
        /// catch (RuntimeException ex)
        /// {
        ///     Console.WriteLine($"Error loading PowerShell modules: {ex.Message}");
        /// }
        /// catch (CommandNotFoundException ex)
        /// {
        ///     Console.WriteLine($"Module not found: {ex.Message}");
        /// }
        /// </code>
        /// </example>

        public void ImportModule(List<string> modulesToLoad)
        {
            foreach (var module in modulesToLoad)
            {
                _powerShellInstance.AddCommand("Import-Module").AddArgument(module);
                _powerShellInstance.Invoke();
            }
        }

        // Check if the module is installed
        private bool IsModuleInstalled(string moduleName)
        {
            _powerShellInstance.AddCommand("Get-Module");
            _powerShellInstance.AddParameter("Name", moduleName);
            _powerShellInstance.AddParameter("ListAvailable", true);

            var result = _powerShellInstance.Invoke();
            return result.Count > 0;
        }

        /// <summary>
        /// Installs a list of PowerShell modules using a shared PowerShell session.
        /// </summary>
        /// <param name="modulesToInstall">A list of module names to be installed. Each module name must correspond to a module
        /// that is available in the repository accessible by the PowerShell environment.</param>
        /// <remarks>
        /// This method uses the 'Install-Module' cmdlet to install each specified module into the current user's environment.
        /// 'AllowClobber' is set to true to allow overwriting existing commands, and 'Scope' is set to 'CurrentUser'.
        /// Errors during installation of any module will terminate the process and throw an exception.
        /// </remarks>
        /// <exception cref="System.Exception">
        /// Thrown if any module cannot be installed due to errors during the execution of the command, such as conflicts, permissions issues, or network problems.
        /// </exception>
        /// <example>
        /// How to use the InstallModule method:
        /// <code>
        /// try
        /// {
        ///     List<string> modules = new List<string> { "Pester", "AzureRM" };
        ///     PowerShellBootstrapper.InstallModule(modules);
        /// }
        /// catch (Exception ex)
        /// {
        ///     Console.WriteLine("Failed to install modules: " + ex.Message);
        /// }
        /// </code>
        /// </example>
        public void InstallModule(List<string> modulesToInstall)
        {
            Console.WriteLine("Starting installation of PowerShell modules.");

            foreach (var moduleName in modulesToInstall)
            {
                _powerShellInstance.Commands.Clear();
                _powerShellInstance.AddCommand("Install-Module")
                    .AddParameter("Name", moduleName)
                    .AddParameter("AllowClobber", true)
                    .AddParameter("Scope", "CurrentUser");

                _powerShellInstance.Invoke();

                if (_powerShellInstance.HadErrors)
                {
                    throw new Exception($"Failed to install PowerShell module: {moduleName}");
                }

                Console.WriteLine($"Successfully installed module: {moduleName}");
            }
        }
    }
}