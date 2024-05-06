using System.Dynamic;
using System.Management.Automation;

/// <summary>
/// Represents a dynamic object to execute PowerShell cmdlets in a fluent manner.
/// This class allows for dynamic invocation of cmdlets and parameters handling.
/// </summary>
public class PowerShellCmdlet : DynamicObject
{
    private PowerShell _powerShellSession;
    private string _currentCmdlet;
    private Dictionary<string, object> _parameters = new Dictionary<string, object>();

    /// <summary>
    /// Initializes a new instance of the PowerShellCmdlet class using an existing PowerShell session.
    /// </summary>
    /// <param name="powerShellSession">The PowerShell session this object will operate with.</param>

    public PowerShellCmdlet(PowerShell powerShellSession)
    {
        _powerShellSession = powerShellSession;
    }

    /// <summary>
    /// Provides a dynamic way to invoke a PowerShell cmdlet. The method name called on this instance
    /// is interpreted as a cmdlet name, and arguments are treated as parameters for the cmdlet.
    /// </summary>
    /// <param name="binder">Provides information about the dynamic operation.</param>
    /// <param name="args">The arguments passed to the dynamic method call.</param>
    /// <param name="result">The result of the dynamic call, which is this instance for fluent chaining.</param>
    /// <returns>True if the operation is successful, otherwise throws an exception.</returns>
    /// <exception cref="InvalidOperationException">Thrown when the cmdlet does not exist in the PowerShell session.</exception>
    public override bool TryInvokeMember(InvokeMemberBinder binder, object[] args, out object result)
    {
        string cmdletName = binder.Name.Replace('_', '-');  // Convert method name to potential cmdlet name
        _parameters.Clear(); // Reset parameters for each cmdlet invocation

        // Check if the cmdlet is available in the current session
        if (!CmdletExists(cmdletName))
        {
            throw new InvalidOperationException($"The cmdlet '{cmdletName}' is not available in the current PowerShell session.");
        }

        _currentCmdlet = cmdletName;  // Set the current cmdlet

        if (args.Length > 0 && binder.CallInfo.ArgumentNames.Count > 0)
        {
            // Assume that each method call with parameters includes the names
            for (int i = 0; i < args.Length; i++)
            {
                _parameters[binder.CallInfo.ArgumentNames[i]] = args[i];
            }
        }

        result = this;  // Return this to continue the fluent interface
        return true;
    }

    /// <summary>
    /// Checks if a cmdlet exists in the current PowerShell session.
    /// </summary>
    /// <param name="cmdletName">The name of the cmdlet to check.</param>
    /// <returns>True if the cmdlet exists, false otherwise.</returns>
    private bool CmdletExists(string cmdletName)
    {
        _powerShellSession.Commands.Clear();
        _powerShellSession.AddCommand("Get-Command").AddParameter("Name", cmdletName);

        var results = _powerShellSession.Invoke();
        return results.Count > 0;
    }

    /// <summary>
    /// Executes the configured cmdlet with any specified parameters.
    /// </summary>
    /// <returns>A collection of PSObject that represents the output of the cmdlet.</returns>
    public ICollection<PSObject> Run()
    {
        _powerShellSession.Commands.Clear();
        _powerShellSession.AddCommand(_currentCmdlet);
        foreach (var param in _parameters)
        {
            _powerShellSession.AddParameter(param.Key, param.Value);
        }

        return _powerShellSession.Invoke();
    }

    /// <summary>
    /// Disposes the underlying PowerShell session.
    /// </summary>
    public void Dispose()
    {
        _powerShellSession.Dispose();
    }
}