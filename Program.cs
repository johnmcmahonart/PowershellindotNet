using System.Management.Automation;

namespace PowershellFromdotNet
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            var testEnvironment = new PowershellEnvironment();
            dynamic cmdlet = testEnvironment.Cmdlet();
            var results = cmdlet.Get_Process().Run();
            foreach (PSObject psObject in results)
            {
                Console.WriteLine(psObject.Properties["ProcessName"].Value);
            }
        }
    }
}