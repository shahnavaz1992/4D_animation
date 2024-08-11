# 4D_animation
This is a Navisworks plugin using C\# to automate the creation of 4D model animations based on information about module and crane locations and schedules.

How to Run a Navisworks Plugin Developed in Visual Studio
Set Up the Development Environment

Ensure you have Navisworks installed on your machine.
Ensure you have Visual Studio installed with the .NET desktop development workload.
Open the Solution in Visual Studio

Launch Visual Studio.
Go to File > Open > Project/Solution.
Navigate to the location of your Autoanimation.sln file and open it.
Build the Solution

Set the build configuration to Debug or Release from the dropdown menu in the toolbar.
Right-click on the solution in Solution Explorer and select Build Solution (or press Ctrl + Shift + B).
Ensure the build succeeds without errors. This will generate the DLL file for your plugin.
Deploy the Plugin

Locate the output DLL file, typically found in the bin/Debug or bin/Release directory of your project.
Use the following xcopy command to copy the DLL file to the Navisworks plugin directory:
bash
Copy code
xcopy /Y /I "path\to\your\plugin.dll" "C:\ProgramData\Autodesk\Navisworks\Plugins"
Replace "path\to\your\plugin.dll" with the actual path to your DLL file.
Run Navisworks

Launch Navisworks.
Go to the Add-Ins tab or section in the Navisworks interface.
You should see your plugin listed there if it was successfully deployed.
Test the Plugin

Run the plugin by selecting it from the Add-Ins tab.
Verify that the plugin functions as expected within Navisworks.
Troubleshooting
Plugin Not Visible: Ensure that the DLL was copied to the correct folder and Navisworks was restarted.
Errors: Review the output window in Visual Studio for any errors and consult the plugin documentation for additional configuration requirements.
