using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;

[assembly:AssemblyCopyright("Smart Ace Designs")]
[assembly:AssemblyVersion("1.1.0.0")]
[assembly:AssemblyFileVersion("1.1.0.0")]
[assembly:AssemblyInformationalVersion("1.1.0.0")]
[assembly:AssemblyProduct("PowerShell SmartLaunch Utility")]
[assembly:AssemblyTitle("PowerShell SmartLaunch Utility")]

namespace SmartLaunch
{
    static class Globals
    {
        // Default script name
        public const string ScriptName = "Script.ps1";

        // PowerShell versions
        public const string PWSH7 = @"C:\Program Files\PowerShell\7\pwsh.exe";
        public const string PSWIN = @"C:\Windows\System32\WindowsPowerShell\v1.0\PowerShell.exe";

        // Command-line arguments
        public const string Args_ExcludePWSH7 = "-EXCLUDE:PS7";
        public const string Args_ExcludeAll = "-EXCLUDE:ALL";
        public const string Args_Debug = "-DEBUG";
    }

    class Program
    {
        // Function: Resolves either the MSIX alias path or traditional MSI path
        private static string GetPWSH7Path()
        {
            // Test for the MSIX execution alias
            string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string msixAliasPath = Path.Combine(localAppData, @"Microsoft\WindowsApps\pwsh.exe");

            if (File.Exists(msixAliasPath))
            {
                return msixAliasPath;
            }
                        
            // Fallback test for the MSI installation path
            if (File.Exists(Globals.PWSH7))
            {
                return Globals.PWSH7;
            }
            
            // Return empty string if neither is found
            return string.Empty;
        }

        // Function: Generate unique extraction directory
        private static string GenerateExtractionDirectory()
        {
            string extractionDirectory = String.Empty;
            int loopCounter = 0;

            do
            {
                // Break out of loop if number of attempts exceeds arbitrary limit
                if (loopCounter > 100)
                {
                    break;
                }
                else
                {
                    extractionDirectory = Path.GetTempPath() + "\\" + Assembly.GetEntryAssembly().GetName().Name + "-" + Guid.NewGuid().ToString().ToUpper();
                    if (!(Directory.Exists(extractionDirectory)))
                    {
                        Directory.CreateDirectory(extractionDirectory);
                    }
                    else
                    {
                        loopCounter++;
                    }
                }
            }
            while (!(Directory.Exists(extractionDirectory)));
            return extractionDirectory;
        }

        // Function: Extract embedded resource files
        private static void ExtractEmbeddedResource(string outputDirectory, List<string> files)
        {
            if (Directory.Exists(outputDirectory))
            {
                foreach (string file in files)
                {
                    using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(file))
                    {
                        using (FileStream fileStream = new FileStream(Path.Combine(outputDirectory, file), FileMode.Create))
                        {
                            for (int i = 0; i < stream.Length; i++)
                            {
                                fileStream.WriteByte((byte)stream.ReadByte());
                            }
                            fileStream.Close();
                        }
                    }
                }
            }
            else
            {
                throw new DirectoryNotFoundException();
            }
        }

        static void Main(string[] args)
        {
            // Extraction location for the script and support files
            string extractionPath = GenerateExtractionDirectory();

            // Version of PowerShell to be excluded from use
            string excludedPowershellVersion = String.Empty;

            // Initial PowerShell console window behavior
            bool showPowerShellConsole = false;

            // Parse command-line arguments
            if (args.Length == 1)
            {
                if (args[0].ToUpper() == Globals.Args_Debug)
                {
                    showPowerShellConsole = true;
                }
                else
                {
                    excludedPowershellVersion = args[0].ToUpper();
                }
            }
            else if (args.Length > 1)
            {
                if (args[0].ToUpper() == Globals.Args_Debug)
                {
                    showPowerShellConsole = true;
                    excludedPowershellVersion = args[1].ToUpper();
                }
                else if (args[1].ToUpper() == Globals.Args_Debug)
                {
                    showPowerShellConsole = true;
                    excludedPowershellVersion = args[0].ToUpper();
                }
            }

            try
            {
                // Extract the embedded resources
                List<string> resourceList = new List<string>();
                foreach (string resource in Assembly.GetExecutingAssembly().GetManifestResourceNames())
                {
                    resourceList.Add(resource);
                }
                ExtractEmbeddedResource(extractionPath, resourceList);

                // Specify the execution parameters for the process
                Process powershellProcess = new Process();
                if (showPowerShellConsole)
                {
                    // Show PowerShell console window for debugging purposes
                    powershellProcess.StartInfo.Arguments = "-NoLogo -NoProfile -NoExit -File " + extractionPath + "\\" + Globals.ScriptName;
                }
                else
                {
                    // Hide PowerShell console window
                    powershellProcess.StartInfo.Arguments = "-NoLogo -NoProfile -Window Hidden -File " + extractionPath + "\\" + Globals.ScriptName;
                    powershellProcess.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                }

                // Allow Windows to unwrap an MSIX app execution alias smoothly
                powershellProcess.StartInfo.UseShellExecute = true;

                // Select the version of PowerShell to use
                string pwsh7Path = GetPWSH7Path();
                if (!string.IsNullOrEmpty(pwsh7Path) && excludedPowershellVersion != Globals.Args_ExcludePWSH7 && excludedPowershellVersion != Globals.Args_ExcludeAll)
                {
                    powershellProcess.StartInfo.FileName = pwsh7Path;
                }
                else
                {
                    powershellProcess.StartInfo.FileName = Globals.PSWIN;
                }
                
                // Execute the script using the selected version of PowerShell
                powershellProcess.Start();
                powershellProcess.WaitForExit();

                // Remove the extracted files and directory after the PowerShell process exits
                if (Directory.Exists(extractionPath))
                {
                    Directory.Delete(extractionPath, true);
                }
            }

            catch
            {
                // Remove the extracted files and directory when an exception is caught
                if (Directory.Exists(extractionPath))
                {
                    Directory.Delete(extractionPath, true);
                }
            }
        }
    }
}
