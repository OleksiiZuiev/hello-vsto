using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace HelloVsto
{
    /// <summary>
    /// Custom ribbon implementation using IRibbonExtensibility interface.
    /// This approach generates ribbon XML dynamically in code (not using Ribbon Designer).
    /// </summary>
    [ComVisible(true)]
    public class HelloRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        private readonly ThisAddIn addIn;

        public HelloRibbon(ThisAddIn addIn)
        {
            this.addIn = addIn;
        }

        #region IRibbonExtensibility Members

        /// <summary>
        /// Called by VSTO to get the custom ribbon XML.
        /// This defines the structure of our custom ribbon tab and controls.
        /// </summary>
        public string GetCustomUI(string ribbonID)
        {
            addIn.LogLifecycleEvent($"GetCustomUI() called with RibbonID: {ribbonID}");

            return GetRibbonXml();
        }

        #endregion

        #region Ribbon Callbacks

        /// <summary>
        /// Called when the ribbon is loaded. Provides access to the ribbon UI object.
        /// </summary>
        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
            addIn.LogLifecycleEvent("Ribbon_Load() - Ribbon UI initialized");
        }

        /// <summary>
        /// Called when the "Hello" button is clicked.
        /// Writes "Hello VSTO" to cell A1 of the active worksheet.
        /// </summary>
        public void OnHelloButtonClick(Office.IRibbonControl control)
        {
            try
            {
                addIn.LogLifecycleEvent("OnHelloButtonClick() - Button clicked!");

                // Get the active worksheet using the addIn reference (not Globals.ThisAddIn which may be null)
                var worksheet = (Microsoft.Office.Interop.Excel.Worksheet)addIn.Application.ActiveSheet;

                // Write to cell A1
                var range = worksheet.Range["A1"];
                range.Value2 = "Hello VSTO";

                addIn.LogLifecycleEvent("Successfully wrote 'Hello VSTO' to cell A1");
            }
            catch (Exception ex)
            {
                addIn.LogLifecycleEvent($"ERROR in OnHelloButtonClick: {ex.Message}");
                System.Windows.Forms.MessageBox.Show(
                    $"Error: {ex.Message}",
                    "Hello VSTO Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Called when the "View Logs" button is clicked.
        /// Opens the log file in the default text editor.
        /// </summary>
        public void OnViewLogsButtonClick(Office.IRibbonControl control)
        {
            try
            {
                addIn.LogLifecycleEvent("OnViewLogsButtonClick() - Opening log file...");
                addIn.OpenLogFile();
            }
            catch (Exception ex)
            {
                addIn.LogLifecycleEvent($"ERROR in OnViewLogsButtonClick: {ex.Message}");
                System.Windows.Forms.MessageBox.Show(
                    $"Error opening log file: {ex.Message}",
                    "Hello VSTO Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        #endregion

        #region Ribbon XML Generation

        /// <summary>
        /// Generates the ribbon XML that defines our custom tab and buttons.
        /// This is the code-based approach (alternative to using Ribbon Designer).
        /// </summary>
        private string GetRibbonXml()
        {
            return @"
                <customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='Ribbon_Load'>
                  <ribbon>
                    <tabs>
                      <tab id='HelloVstoTab' label='Hello VSTO'>
                        <group id='HelloGroup' label='Greetings'>
                          <button
                            id='HelloButton'
                            label='Hello'
                            size='large'
                            onAction='OnHelloButtonClick'
                            imageMso='HappyFace' />
                          <button
                            id='ViewLogsButton'
                            label='View Logs'
                            size='large'
                            onAction='OnViewLogsButtonClick'
                            imageMso='FileOpen' />
                        </group>
                      </tab>
                    </tabs>
                  </ribbon>
                </customUI>";
        }

        #endregion
    }
}
