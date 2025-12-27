using VbaMcpServer.GUI.Forms;

namespace VbaMcpServer.GUI;

static class Program
{
    /// <summary>
    ///  The main entry point for the application.
    /// </summary>
    [STAThread]
    static void Main()
    {
        try
        {
            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            ApplicationConfiguration.Initialize();
            Application.Run(new MainForm());
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"Failed to start VBA MCP Server Manager:\n\n{ex.GetType().Name}: {ex.Message}\n\nStack Trace:\n{ex.StackTrace}",
                "Startup Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }
}