# New-ServerRestart
**Warning:** This ScriptoForm is for demonstration purposes only and not intended for use in a production environment.

This repository contains all files required to build the **New-ServerRestart** ScriptoForm project.  A *ScriptoForm* is a PowerShell script that generates and displays a [Microsoft Windows Forms](https://learn.microsoft.com/en-us/dotnet/desktop/winforms/overview/?view=netdesktop-10.0#introduction) application that can be used for a specific management or system administration task in a network environment.  A *ScriptoForm project* is the set of files and folders, including the PowerShell script, that can be compiled into an executable file using the Microsoft .NET CLI utility (dotnet.exe) which is available with any [Microsoft .NET SDK](https://dotnet.microsoft.com/en-us/download/dotnet).  Included in the repository is the Build.cs C# file which the compiler will use as the source for the executable, and the Build.csproj C# project file which provides the set of instructions used to compile the executable.

## Purpose
The **New-ServerRestart** script launches a Windows form that provides a method to schedule a one time restart on a remote server.

![Screenshot 2023-05-17 211734](https://github.com/Smart-Ace-Designs/New-ServerRestart/assets/132539186/5c606b54-caf6-4198-b525-60e85f241daa)

## Requirements
- This project supports the following command shells:
    - [Windows PowerShell 5.1](https://docs.microsoft.com/en-us/powershell/scripting/overview?view=powershell-5.1)
    - [PowerShell 7.4.x](https://docs.microsoft.com/en-us/powershell/scripting/overview?view=powershell-7.4)
    - [PowerShell 7.5.x](https://docs.microsoft.com/en-us/powershell/scripting/overview?view=powershell-7.5)
- This project supports the following Microsoft .NET frameworks:
    - [Microsoft .NET 4.x](https://dotnet.microsoft.com/en-us/download/dotnet-framework)
    - [Microsoft .NET 8.x](https://dotnet.microsoft.com/en-us/download/dotnet/8.0)
    - [Microsoft .NET 9.x](https://dotnet.microsoft.com/en-us/download/dotnet/9.0)
    - [Microsoft .NET 10.x](https://dotnet.microsoft.com/en-us/download/dotnet/10.0)

## Compile Instructions
Perform the following prerequisite steps:
- Install the [Microsoft .NET 10.x SDK](https://dotnet.microsoft.com/en-us/download/dotnet/10.0) on your development machine.
- Clone the repository to your development machine.

Use any of the below workflows to create an executable file of the PowerShell script that is compatible with the framework version specified:

Microsoft .NET 4.x Framework
- Open a supported command shell and navigate to the *Build* subdirectory in your local repository directory.
- Run the following command from within your *Build* subdirectory:<br>
``dotnet publish -f net48 -v q -o ..\Release\Legacy; dotnet clean -f net48 -v q``
- The compiled executable will be created in the *Release\Legacy* subdirectory of your local repository directory.  This location can be changed by modifying the ``-o`` argument in the above command.
- The latest [Microsoft .NET 4.x Framework Runtime](https://dotnet.microsoft.com/en-us/download/dotnet-framework/net48) will be required on any computer used to run the executable.

Microsoft .NET 8.x Framework
- Open a supported command shell and navigate to the *Build* subdirectory in your local repository directory.
- Run the following command from within your *Build* subdirectory:<br>
``dotnet publish -f net8.0-windows -v q -o ..\Release\LTS-Legacy; dotnet clean -f net8.0-windows -v q``
- The compiled executable will be created in the *Release\LTS-Legacy* subdirectory of your local repository directory.  This location can be changed by modifying the ``-o`` argument in the above command.
- The latest [Microsoft .NET 8.x Runtime](https://dotnet.microsoft.com/en-us/download/dotnet/8.0) will be required on any computer used to run the executable.

Microsoft .NET 9.x Framework
- Open a supported command shell and navigate to the *Build* subdirectory in your local repository directory.
- Run the following command from within your *Build* subdirectory:<br>
``dotnet publish -f net9.0-windows -v q -o ..\Release\STS; dotnet clean -f net9.0-windows -v q``
- The compiled executable will be created in the *Release\STS* subdirectory of your local repository directory.  This location can be changed by modifying the ``-o`` argument in the above command.
- The latest [Microsoft .NET 9.x Runtime](https://dotnet.microsoft.com/en-us/download/dotnet/9.0) will be required on any computer used to run the executable.

Microsoft .NET 10.x Framework
- Open a supported command shell and navigate to the *Build* subdirectory in your local repository directory.
- Run the following command from within your *Build* subdirectory:<br>
``dotnet publish -f net10.0-windows -v q -o ..\Release\LTS; dotnet clean -f net10.0-windows -v q``
- The compiled executable will be created in the *Release\LTS* subdirectory of your local repository directory.  This location can be changed by modifying the ``-o`` argument in the above command.
- The latest [Microsoft .NET 10.x Runtime](https://dotnet.microsoft.com/en-us/download/dotnet/10.0) will be required on any computer used to run the executable.

## Executable Notes
- When the executable file is run it will extract all resource files that were included in the compilation process to a unique temporary extraction directory in the user's profile directory.
- The executable will attempt to use the latest version of PowerShell discovered on the local machine to execute the extracted script file unless excluded by one of the command-line arguments noted below.  If a version of PowerShell is excluded by a command-line argument then the executable will attempt to use the next latest version of PowerShell discovered on the system.  The executable will always default to using Windows PowerShell if no other versions are available for use.
- After the PowerShell script execution has completed, all extracted files and the extraction directory are deleted and the executable terminates.
- The following optional command-line arguments can be used to control operation of the executable:
  | Argument     | Purpose                           | Notes                                          |
  | ------------ | --------------------------------- | ---------------------------------------------- |
  | -exclude:all | Exclude PowerShell x.x.x versions | Do not use with other *exclude* arguments      |
  | -exclude:ps7 | Exclude PowerShell 7.x.x versions | Do not use with other *exclude* arguments      |
  | -debug       | Show console window               | Use individually or with an *exclude* argument |
