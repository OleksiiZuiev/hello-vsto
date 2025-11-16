using System;
using System.Diagnostics;
using System.IO;
using Office = Microsoft.Office.Core;

namespace HelloVsto
{
    public partial class ThisAddIn
    {
        private HelloRibbon ribbon;
        private readonly string sessionId = Guid.NewGuid().ToString("N").Substring(0, 8);
        private bool isFirstLog = true;

        /// <summary>
        /// Gets the log file path in %TEMP%\HelloVsto\HelloVsto.log
        /// </summary>
        private string LogFilePath
        {
            get
            {
                var logDirectory = Path.Combine(Path.GetTempPath(), "HelloVsto");
                return Path.Combine(logDirectory, "HelloVsto.log");
            }
        }

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
            LogSessionEnd();
        }

        /// <summary>
        /// Logs the end of the current session.
        /// </summary>
        private void LogSessionEnd()
        {
            try
            {
                var sessionFooter = new System.Text.StringBuilder();
                sessionFooter.AppendLine($"[{sessionId}] ========================================");
                sessionFooter.AppendLine($"[{sessionId}] SESSION END: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                sessionFooter.AppendLine($"[{sessionId}] ========================================");

                File.AppendAllText(LogFilePath, sessionFooter.ToString());
                Debug.Write(sessionFooter.ToString());
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to log session end: {ex.Message}");
            }
        }

        /// <summary>
        /// Helper method to log lifecycle events to file.
        /// Appends timestamped entries to log file.
        /// </summary>
        public void LogLifecycleEvent(string message)
        {
            try
            {
                // Ensure log directory exists
                var logDirectory = Path.GetDirectoryName(LogFilePath);
                if (!Directory.Exists(logDirectory))
                {
                    Directory.CreateDirectory(logDirectory);
                }

                // Write session header on first log entry
                if (isFirstLog)
                {
                    var sessionHeader = new System.Text.StringBuilder();
                    sessionHeader.AppendLine();
                    sessionHeader.AppendLine("========================================");
                    sessionHeader.AppendLine($"SESSION START: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                    sessionHeader.AppendLine($"Session ID: {sessionId}");
                    sessionHeader.AppendLine("========================================");

                    File.AppendAllText(LogFilePath, sessionHeader.ToString());
                    Debug.Write(sessionHeader.ToString());

                    isFirstLog = false;
                }

                // Write timestamped message to file
                string timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
                string logEntry = $"[{sessionId}] [{timestamp}] {message}";

                File.AppendAllText(LogFilePath, logEntry + Environment.NewLine);

                // Also write to Debug output for Visual Studio debugging
                Debug.WriteLine(logEntry);
            }
            catch (Exception ex)
            {
                // If file logging fails, only write to Debug output (don't crash the add-in)
                Debug.WriteLine($"File logging failed: {ex.Message}");
            }
        }

        /// <summary>
        /// Gets the full path to the log file.
        /// </summary>
        public string GetLogFilePath()
        {
            return LogFilePath;
        }

        /// <summary>
        /// Opens the log file in the default text editor.
        /// </summary>
        public void OpenLogFile()
        {
            try
            {
                if (File.Exists(LogFilePath))
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = LogFilePath,
                        UseShellExecute = true
                    });
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show(
                        $"Log file does not exist yet.\n\nLog file will be created at:\n{LogFilePath}",
                        "Hello VSTO - Log File",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    $"Failed to open log file:\n{ex.Message}",
                    "Hello VSTO - Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

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
