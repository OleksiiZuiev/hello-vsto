using System;
using System.IO;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace HelloVsto
{
    public partial class ThisAddIn
    {
        // P/Invoke declarations for console management
        [DllImport("kernel32.dll")]
        private static extern bool AllocConsole();

        [DllImport("kernel32.dll")]
        private static extern bool FreeConsole();

        private HelloRibbon ribbon;

        /// <summary>
        /// This is the FIRST method called by VSTO runtime - even before Startup event.
        /// Use this to create and return the ribbon extensibility object.
        /// </summary>
        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            LogLifecycleEvent("CreateRibbonExtensibilityObject() - FIRST EVENT CALLED");

            // Create our custom ribbon
            ribbon = new HelloRibbon(this);

            return ribbon;
        }

        /// <summary>
        /// Called when the add-in is loaded. Excel Application object is available here.
        /// </summary>
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            LogLifecycleEvent("ThisAddIn_Startup() - Add-in is loading");
            LogLifecycleEvent($"Excel Version: {Application.Version}");
            LogLifecycleEvent($"Excel Build: {Application.Build}");
            LogLifecycleEvent($"Add-in is ready!");
        }

        /// <summary>
        /// Called when the add-in is unloading. Excel is shutting down.
        /// </summary>
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            LogLifecycleEvent("ThisAddIn_Shutdown() - Add-in is unloading");
            LogLifecycleEvent("Cleaning up resources...");

            // Give user a moment to see shutdown message
            System.Threading.Thread.Sleep(500);

            // Free the console before Excel closes
            FreeConsole();
        }

        /// <summary>
        /// Helper method to log lifecycle events to console.
        /// Creates console on first call.
        /// </summary>
        public void LogLifecycleEvent(string message)
        {
            try
            {
                // Create console on first log call
                if (!IsConsoleAllocated)
                {
                    AllocConsole();

                    // Reinitialize console streams for proper output
                    Console.SetOut(new StreamWriter(Console.OpenStandardOutput()) { AutoFlush = true });
                    Console.SetError(new StreamWriter(Console.OpenStandardError()) { AutoFlush = true });

                    IsConsoleAllocated = true;

                    // Write header
                    Console.WriteLine("=====================================");
                    Console.WriteLine("   HELLO VSTO - LIFECYCLE LOGGER");
                    Console.WriteLine("=====================================");
                    Console.WriteLine();
                }

                // Write timestamped message
                string timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
                Console.WriteLine($"[{timestamp}] {message}");
            }
            catch (Exception ex)
            {
                // If console logging fails, fail silently (don't crash the add-in)
                System.Diagnostics.Debug.WriteLine($"Console logging failed: {ex.Message}");
            }
        }

        private bool IsConsoleAllocated { get; set; } = false;

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
