; ExcelAddinInstaller
; InnoSetup script to install and activate native Excel addins.
; Originally developed for Daniel's XL Toolbox (xltoolbox.sf.net).
; Requires the InnoSetup Preprocessor (ISPP).
; Copyright (C) 2013  Daniel Kraus <http://github.com/bovender>
; 
; This program is free software: you can redistribute it and/or modify
; it under the terms of the GNU General Public License as published by
; the Free Software Foundation, either version 3 of the License, or
; (at your option) any later version.
; 
; This program is distributed in the hope that it will be useful,
; but WITHOUT ANY WARRANTY; without even the implied warranty of
; MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.	See the
; GNU General Public License for more details.
; 
; You should have received a copy of the GNU General Public License
; along with this program.	If not, see <http://www.gnu.org/licenses/>.

[Setup]
#include "config.iss"

AppName={#product}
VersionInfoProductName={#product}
AppVerName={#product} {#version}
AppPublisher={#company}
VersionInfoCompany={#company}
AppCopyright={#yearspan} {#company}
VersionInfoCopyright={#yearspan} {#company}
VersionInfoDescription=custom wrappers and Windows libraries.
VersionInfoVersion={#longversion}
VersionInfoProductVersion={#longversion}
VersionInfoTextVersion={#version}

; Make this setup program work with 32-bit and 64-bit Windows
ArchitecturesAllowed=x86 x64
ArchitecturesInstallIn64BitMode=x64

; Always write a log file
SetupLogging=true

; Addins do not need a program group and no user-configurable
; installation folder.
DisableProgramGroupPage=true
DisableDirPage=true
CreateAppDir=true
AppendDefaultDirName=false
DisableReadyPage=true

; Allow normal users to install the addin into their profile.
; This directive also ensures that the uninstall information is
; stored in the user profile rather than a system folder (which
; would require administrative rights).
PrivilegesRequired=lowest

; The destination folder is also determined by code since
; different language versions of Excel expect addins in
; localized folders.
DefaultDirName={#DEFINSDIR}

; The uninstall information is to be put into a subfolder of
; the installation in the user's profile. If this were not set,
; the uninstall information would be put into the Windows system
; folder by InnoSetup.
UninstallFilesDir={#DEFINSDIR}\uninstall

InternalCompressLevel=max
SolidCompression=true

ShowLanguageDialog=no
ChangesEnvironment=True

[Files]
Source: "images\logo.ico"; DestDir: "{#DEFINSDIR}\";

Source: "{#sourcedir}\CoolPropLib.h"; DestDir: "{#DLLINSDIR}\"; Tasks: SharedLibs
Source: "{#sourcedir}\CoolProp_stdcall.dll"; DestDir: "{#DLLINSDIR}\"; Tasks: SharedLibs ExcelAddin
Source: "{#sourcedir}\CoolProp_cdecl.dll"; DestDir: "{#DLLINSDIR}\"; Tasks: SharedLibs
Source: "{#sourcedir}\CoolProp_x64.dll"; DestDir: "{#DLLINSDIR}\"; Tasks: SharedLibs ExcelAddin

Source: "{#sourcedir}\CoolProp_stdcall.dll"; DestDir: "{#DLLINSDIR}\"; DestName: "CoolProp.dll"; Tasks: SharedLibs\32BitStdcall
Source: "{#sourcedir}\CoolProp_cdecl.dll"; DestDir: "{#DLLINSDIR}\"; DestName: "CoolProp.dll"; Tasks: SharedLibs\32BitCdecl
Source: "{#sourcedir}\CoolProp_x64.dll"; DestDir: "{#DLLINSDIR}\"; DestName: "CoolProp.dll"; Tasks: SharedLibs\64Bit

Source: "{#sourcedir}\TestExcel.xlsx"; DestDir: "{#EXAMPLDIR}\"; Flags: uninsneveruninstall; Tasks: ExcelAddin
Source: "{#sourcedir}\*.xlam"; DestDir: "{code:GetDestDir}\"; Check: ShouldInstallFile(12,16); AfterInstall: ActivateAddin(12,16)
Source: "{#sourcedir}\*.xla"; DestDir: "{code:GetDestDir}\"; Excludes: "*.xlam"; Check: ShouldInstallFile(9,11); AfterInstall: ActivateAddin(9,11)

Source: "{#sourcedir}\EES\CoolProp.htm"; DestDir: "{#EESINSDIR}\"; Tasks: EesUserLib
Source: "{#sourcedir}\EES\CoolProp.LIB"; DestDir: "{#EESINSDIR}\"; Tasks: EesUserLib
Source: "{#sourcedir}\EES\COOLPROP_EES.dlf"; DestDir: "{#EESINSDIR}\"; Tasks: EesUserLib
Source: "{#sourcedir}\EES\CoolProp_EES_Sample.EES"; DestDir: "{#EXAMPLDIR}\"; Flags: uninsneveruninstall; Tasks: EesUserLib

[Tasks]
; We make it optional for users to have the addin activated for use in
; Excel. In most cases, this will be left enabled by users (everything
; else does not make sense).

; Name: ActivateAddin; Description: {cm:taskActivate}; 
; Name: AddDirToPath; Description: {cm:taskAddDirToPath}; 
; Name: InstallEES; Description: {cm:taskInstallEES}; 

Name: SharedLibs;              Description: {cm:taskSharedLibs};             GroupDescription: "CoolProp Library";
;Name: AddToPath;  Description: {cm:taskAddToPath};  GroupDescription: "CoolProp Library"; 
Name: SharedLibs\32BitCdecl;   Description: {cm:taskSharedLibs32BitCdecl};   GroupDescription: "CoolProp Library"; Flags: exclusive unchecked
Name: SharedLibs\32BitStdcall; Description: {cm:taskSharedLibs32BitStdcall}; GroupDescription: "CoolProp Library"; Flags: exclusive unchecked
Name: SharedLibs\64Bit;        Description: {cm:taskSharedLibs64Bit};        GroupDescription: "CoolProp Library"; Flags: exclusive

Name: ExcelAddin;          Description: {cm:taskExcelAddin};         GroupDescription: "Custom wrappers"; Flags: checkablealone
; Name: ExcelAddin\Example;  Description: {cm:taskExcelAddinExample};  GroupDescription: "Custom wrappers";
; Name: ExcelAddin\Activate; Description: {cm:taskExcelAddinActivate}; GroupDescription: "Custom wrappers";

Name: EesUserLib;         Description: {cm:taskEesUserLib};        GroupDescription: "Custom wrappers"; Flags: checkablealone
; Name: EesUserLib\Example; Description: {cm:taskEesUserLibExample}; GroupDescription: "Custom wrappers";


; Name: ExcelAddin;               Description: {cm:taskExcelAddin};   GroupDescription: "Microsoft Excel";
; Name: ExcelAddin\ActivateAddin; Description: {cm:taskActivate};     GroupDescription: "Microsoft Excel";
; Name: ExcelAddin\ActivateAddin; Description: {cm:taskActivate};     GroupDescription: "Microsoft Excel";

; Name: EESwrapper;               Description: {cm:taskEESwrapper}; GroupDescription: "Engineering Equation Solver (EES)"; Components: main
; Name: MATLABfiles;              Description: {cm:taskMATLABfiles}; GroupDescription: "MATLAB interface"; Components: main


; Name: desktopicon; Description: "Create a &desktop icon"; GroupDescription: "Additional icons:"; Components: main
; Name: desktopicon\common; Description: "For all users"; GroupDescription: "Additional icons:"; Components: main; Flags: exclusive
; Name: desktopicon\user; Description: "For the current user only"; GroupDescription: "Additional icons:"; Components: main; Flags: exclusive unchecked
; Name: quicklaunchicon; Description: "Create a &Quick Launch icon"; GroupDescription: "Additional icons:"; Components: main; Flags: unchecked
; Name: associate; Description: "&Associate files"; GroupDescription: "Other tasks:"; Flags: unchecked

; Define any additional tasks in the custom tasks.iss file.
#ifexist "tasks.iss"
  #include "tasks.iss"
#endif

[Registry]
Root: "HKCU"; Subkey: "Environment"; ValueType: expandsz; ValueName: "Path"; ValueData: "{olddata};{#DLLINSDIR}"; Check: NeedsAddPath('{#DLLINSDIR}')
;Root: "HKCU"; Subkey: "Environment"; ValueType: expandsz; ValueName: "Path"; ValueData: "{olddata};{code:GetDestDir}\CoolProp"; Check: NeedsAddPath('{code:GetDestDir}\CoolProp')
;http://www.jrsoftware.org/isfaq.php#env
;http://stackoverflow.com/questions/3304463/how-do-i-modify-the-path-environment-variable-when-running-an-inno-setup-install


; Use the original file and an additional file for custom functions.
[Code]
#include "inc/code.iss"
#include "code.iss"


[Languages]
Name: English; MessagesFile: compiler:Default.isl; 
Name: Deutsch; MessagesFile: compiler:Languages\German.isl; 
Name: Dansk; MessagesFile: compiler:Languages\Danish.isl; 


[CustomMessages]
; Use a different file to simplify the handling of unicode messages
#include "messages.iss"

; vim: set ts=2 sts=2 sw=2 noet tw=60 fo+=lj cms=;%s 
