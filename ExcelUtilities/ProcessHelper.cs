using System.Diagnostics;

namespace ExcelUtilities
{
  public static class ProcessHelper
  {
    /// <summary>
    /// Executes the process specified.
    /// </summary>
    /// <param name="filePath">The executable file path.</param>
    /// <param name="arguments">The arguments passed to the process.</param>
    /// <param name="runAsAdmin">Run as admin, if set to <c>true</c>.</param>
    /// <returns>Success of the execution.</returns>
    public static bool ExecuteProcess(string filePath, string arguments = null, bool runAsAdmin = false)
    {
      using (var process = new Process())
      {
        var processInfo = new ProcessStartInfo(filePath, arguments)
        {
          Verb = runAsAdmin ? "runas" : null
        };

        process.StartInfo = processInfo;
        return process.Start();
      }
    }
  }
}
