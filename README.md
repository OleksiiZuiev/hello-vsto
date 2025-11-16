# Hello VSTO - Excel Add-in Learning Project

A simple Excel VSTO (Visual Studio Tools for Office) add-in created for learning purposes. This project demonstrates the fundamental concepts of VSTO development including ribbon customization, lifecycle event handling, and real-time console logging.

## What This Add-in Does

- ✅ Creates a custom "Hello VSTO" tab in Excel ribbon
- ✅ Adds a "Hello" button with a happy face icon
- ✅ Writes "Hello VSTO" to cell A1 when button is clicked
- ✅ Logs all VSTO lifecycle events to a console window
- ✅ Includes automated build and launch script

## Quick Start

### Prerequisites

- Visual Studio 2019 or 2022 with:
  - .NET desktop development workload
  - Office/SharePoint development workload
- Microsoft Excel (Office 2016 or later recommended)
- PowerShell 5.1 or later

### Build and Run

1. Clone this repository
2. Open PowerShell in the repository root
3. Run the build script:

```powershell
.\Build.ps1
```

The script will:
- Restore NuGet packages
- Build the project
- Register the add-in
- Launch Excel with the add-in loaded

### What to Expect

When Excel launches, you'll see:
1. **Excel window** with your workbook
2. **Console window** (separate) showing lifecycle events in real-time
3. **"Hello VSTO" tab** in the Excel ribbon

Click the "Hello" button in the ribbon to write "Hello VSTO" to cell A1.

## Project Structure

```
hello-vsto/
├── HelloVsto/
│   ├── HelloVsto.csproj       # VSTO project file with Office references
│   ├── App.config              # Runtime configuration
│   ├── ThisAddIn.cs            # Main add-in entry point
│   ├── HelloRibbon.cs          # Ribbon customization (code-based)
│   └── Properties/
│       └── AssemblyInfo.cs     # Assembly metadata
├── Build.ps1                   # Automated build and launch script
└── README.md                   # This file
```

## VSTO Architecture Overview

### Project Configuration

The `HelloVsto.csproj` file contains critical VSTO-specific configurations:

**Project Type GUID:**
```xml
<ProjectTypeGuids>{BAA0C2D2-18E2-41B9-852F-F413020CAA33};{FAE04EC0-301F-11D3-BF4B-00C04F79EFBC}</ProjectTypeGuids>
```
- First GUID identifies this as a VSTO project
- Required for Visual Studio to recognize it as an Office add-in

**Key Properties:**
```xml
<OutputType>Library</OutputType>          <!-- VSTO add-ins are DLLs -->
<TargetFrameworkVersion>v4.8</TargetFrameworkVersion>
<DefineConstants>VSTO40</DefineConstants>
<OfficeApplication>Excel</OfficeApplication>
<LoadBehavior>3</LoadBehavior>           <!-- Load on Excel startup -->
```

**Essential References:**
- `Microsoft.Office.Interop.Excel` - Excel object model
- `Microsoft.Office.Tools.*` - VSTO runtime and tools
- `Microsoft.VisualStudio.Tools.Applications.Runtime` - VSTO runtime
- `Office` (Microsoft.Office.Core) - Common Office functionality

### VSTO Lifecycle Events

VSTO add-ins have a specific lifecycle with multiple events. Understanding the order is crucial:

#### 1. CreateRibbonExtensibilityObject() - FIRST EVENT

```csharp
protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
{
    // Called BEFORE any other events
    // Must return the ribbon object
    // Excel Application is NOT available yet
}
```

**Important:**
- This is called **before** `Startup` event
- Use for early initialization
- Must return an `IRibbonExtensibility` implementation

#### 2. GetCustomUI() - Ribbon XML Generation

```csharp
public string GetCustomUI(string ribbonID)
{
    // Return XML that defines ribbon structure
    return GetRibbonXml();
}
```

**Purpose:**
- VSTO calls this to get your custom ribbon definition
- You return XML as a string (programmatic approach)
- Alternative to using Ribbon Designer

#### 3. Ribbon_Load() - Ribbon Initialization

```csharp
public void Ribbon_Load(Office.IRibbonUI ribbonUI)
{
    // Ribbon UI object is now available
    // Can invalidate/update ribbon controls later
}
```

#### 4. ThisAddIn_Startup() - Main Initialization

```csharp
private void ThisAddIn_Startup(object sender, EventArgs e)
{
    // Excel Application object is now available
    // Perform main add-in initialization here
}
```

**Available:**
- `this.Application` - Excel application object
- Full Excel object model access
- Time to set up event handlers, load data, etc.

#### 5. ThisAddIn_Shutdown() - Cleanup

```csharp
private void ThisAddIn_Shutdown(object sender, EventArgs e)
{
    // Excel is closing
    // Clean up resources
    // Cannot be async - use dispatcher if needed
}
```

**Limitations:**
- Cannot use `async/await` (event is synchronous)
- Excel is shutting down - limited time
- Free resources quickly

### Lifecycle Event Order Summary

```
1. CreateRibbonExtensibilityObject()  ← Create and return ribbon
   ↓
2. GetCustomUI(ribbonID)              ← Return ribbon XML
   ↓
3. Ribbon_Load(ribbonUI)              ← Ribbon initialized
   ↓
4. ThisAddIn_Startup()                ← Main initialization
   ↓
   [Add-in is running - user interacts with Excel]
   ↓
5. ThisAddIn_Shutdown()               ← Cleanup on exit
```

## Ribbon Implementation: Code-Based Approach

This project uses **code-based ribbon customization** via `IRibbonExtensibility` interface, not the visual Ribbon Designer.

### Why Code-Based?

**Advantages:**
- More control and flexibility
- Easier to generate dynamic ribbons
- Better for understanding how VSTO works
- Matches DataSnipper and other professional add-ins

**Alternative:** Visual Ribbon Designer (easier for beginners, less flexible)

### Ribbon XML Structure

```xml
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
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
```

**Key Elements:**
- `tab` - Creates a new ribbon tab
- `group` - Groups related controls
- `button` - Clickable button control
- `onAction` - Callback method name
- `imageMso` - Built-in Office icon (HappyFace, FileSave, etc.)

### Button Click Handler

```csharp
public void OnHelloButtonClick(Office.IRibbonControl control)
{
    var worksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
    worksheet.Range["A1"].Value2 = "Hello VSTO";
}
```

**Important:**
- Method signature must match: `void Method(Office.IRibbonControl control)`
- Method name must match `onAction` in XML
- Use `Globals.ThisAddIn.Application` to access Excel

## Console Logging Implementation

This project uses **AllocConsole()** Win32 API to create a console window for real-time logging.

### How It Works

```csharp
[DllImport("kernel32.dll")]
private static extern bool AllocConsole();

// In startup:
AllocConsole();
Console.SetOut(new StreamWriter(Console.OpenStandardOutput()) { AutoFlush = true });
Console.WriteLine("Hello from VSTO!");
```

**Behavior:**
- Creates a **new console window** owned by Excel process
- Console appears alongside Excel window
- Output is immediate and visible
- Console closes when Excel closes

### Why Not PowerShell Console?

Excel is a GUI application (`/SUBSYSTEM:WINDOWS`), so it doesn't inherit the PowerShell console that launched it. `AllocConsole()` creates a separate console window instead.

**Alternatives for production:**
- **File logging** (log4net, NLog, Serilog)
- **Debug.WriteLine()** (Visual Studio Output window)
- **DebugView** (Sysinternals tool)

### What Gets Logged

The add-in logs all major lifecycle events:
- `CreateRibbonExtensibilityObject()` - First event
- `GetCustomUI()` - Ribbon XML generation
- `Ribbon_Load()` - Ribbon initialization
- `ThisAddIn_Startup()` - Main startup
- Excel version and build info
- Button clicks
- `ThisAddIn_Shutdown()` - Cleanup

## Build and Deployment

### MSBuild Process

The `Build.ps1` script uses MSBuild to compile the VSTO project:

```powershell
# 1. Restore NuGet packages
msbuild /t:Restore /p:Configuration=Debug HelloVsto.csproj

# 2. Build project
msbuild /p:Configuration=Debug HelloVsto.csproj
```

**Build Output:**
```
HelloVsto/bin/Debug/
├── HelloVsto.dll           # Your add-in code
├── HelloVsto.dll.manifest  # Assembly manifest
├── HelloVsto.vsto          # ClickOnce deployment manifest
└── [dependencies]          # VSTO runtime, Office PIAs
```

### Registry-Based Installation

For development, VSTO add-ins are registered in the Windows Registry:

**Registry Path:**
```
HKEY_CURRENT_USER\Software\Microsoft\Office\Excel\Addins\HelloVsto
```

**Required Values:**
- `Description` (String) - Human-readable description
- `FriendlyName` (String) - Display name in Excel
- `LoadBehavior` (DWORD) - `3` means "load on startup"
- `Manifest` (String) - Path to .vsto file in format:
  ```
  file:///C:/path/to/HelloVsto.vsto|vstolocal
  ```

**LoadBehavior Values:**
- `0` - Unload (disabled)
- `1` - Load on demand
- `2` - Loaded (currently running)
- `3` - Load on startup (typical for development)
- `8` - Load on first use
- `9` - Load on demand (load immediately)
- `16` - Connect first time only

### ClickOnce Deployment

VSTO uses ClickOnce for deployment. The `.vsto` file is an XML manifest that points to:
- Assembly location
- Dependencies (VSTO runtime, .NET Framework)
- Trust/security settings
- Version information

**Production Deployment Options:**
1. **Network share** - Deploy to file share, users install from UNC path
2. **Web server** - Host on web server, users install from URL
3. **ClickOnce installer** - Self-updating deployment
4. **Group Policy** - Enterprise deployment via AD

## Building Blocks for Further Learning

Now that you understand the basics, here are next steps:

### 1. Task Panes
Add custom UI panels alongside Excel:
```csharp
var taskPane = CustomTaskPanes.Add(new MyUserControl(), "My Panel");
taskPane.Visible = true;
```

### 2. Excel Events
Respond to workbook/worksheet events:
```csharp
this.Application.WorkbookOpen += Application_WorkbookOpen;
worksheet.Change += Worksheet_Change;
```

### 3. Custom Functions (UDFs)
Create Excel formulas in .NET (requires separate project type)

### 4. Ribbon Updates
Dynamically update ribbon controls:
```csharp
ribbon.Invalidate(); // Refresh entire ribbon
ribbon.InvalidateControl("ButtonId"); // Refresh specific control
```

### 5. COM Add-in Registry
Explore `HKEY_CURRENT_USER\Software\Microsoft\Office\Excel\Addins`

## Troubleshooting

### Add-in Doesn't Load

1. **Check registry** - Verify manifest path is correct
2. **Check LoadBehavior** - Should be `3` for auto-load
3. **Check Excel Trust Center**:
   - File → Options → Trust Center → Trust Center Settings
   - Add-ins → Check settings
4. **Look for VSTO errors** in Event Viewer:
   - Windows Logs → Application
   - Filter by source: "VSTO 4.0"

### Console Doesn't Appear

- Verify `AllocConsole()` is called in `CreateRibbonExtensibilityObject()`
- Check if console is hidden behind other windows
- Ensure logging code doesn't throw exceptions

### Button Doesn't Work

- Check method signature matches: `void OnHelloButtonClick(Office.IRibbonControl control)`
- Verify `onAction='OnHelloButtonClick'` in ribbon XML
- Look for errors in console window

### Build Fails

- Ensure Visual Studio with Office development tools installed
- Check .NET Framework 4.8 SDK is installed
- Verify Office PIAs (Primary Interop Assemblies) are installed

## Learning Resources

### Official Documentation
- [VSTO Overview (Microsoft)](https://learn.microsoft.com/en-us/visualstudio/vsto/office-solutions-development-overview-vsto)
- [Excel VSTO Add-ins](https://learn.microsoft.com/en-us/visualstudio/vsto/excel-solutions)
- [Ribbon Overview](https://learn.microsoft.com/en-us/visualstudio/vsto/ribbon-overview)

### Ribbon XML Reference
- [Office Fluent UI XML Schema](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-customui/31f152cf-c8b6-4c1e-82f5-6e7c32e0e8c7)
- [Built-in imageMso Icons](https://www.microsoft.com/en-us/download/details.aspx?id=21103)

### Advanced Topics
- [Custom Task Panes](https://learn.microsoft.com/en-us/visualstudio/vsto/custom-task-panes)
- [Actions Panes](https://learn.microsoft.com/en-us/visualstudio/vsto/actions-pane-overview)
- [Document-Level vs Application-Level](https://learn.microsoft.com/en-us/visualstudio/vsto/application-level-and-document-level-customizations)

## License

This is a learning project - feel free to use, modify, and learn from it!

## Acknowledgments

Inspired by the DataSnipper Excel Add-in architecture and real-world VSTO development practices.