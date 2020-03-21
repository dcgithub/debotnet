<p align="center"> <sup> Currently are over 250 privacy scripts available (03-21-2020) </sup> <p align="center">
	
<p align="center">
	 
<a href="https://github.com/mirinsoft/debotnet/stargazers"><img src="https://img.shields.io/github/stars/mirinsoft/debotnet.svg?style=flat-square"/></a>
<a href="https://github.com/mirinsoft/debotnet/issues">
<img src="https://img.shields.io/github/issues-raw/mirinsoft/debotnet.svg?style=flat-square&logo=github&logoColor=white"
alt="GitHub issues">   
 
<a href="https://github.com/Mirinsoft/Debotnet/releases/latest" target="_blank">
 <img alt="Latest GitHub release" src="https://img.shields.io/github/release/mirinsoft/debotnet.svg?style=flat-square" />
</a>
<a href="https://github.com/Mirinsoft/Debotnet/releases" target="_blank">
 <img alt="Downloads on GitHub" src="https://img.shields.io/github/downloads/Mirinsoft/debotnet/total.svg?style=flat-square" />
</a>
  
<a href="https://twitter.com/intent/tweet?text=Debotnet is a tiny portable tool for controlling Windows 10's many privacy-related settings and keep your personal data private.&url=https%3A%2F%2Fgithub.com%2Fmirinsoft%2Fdebotnet">
<img src="https://img.shields.io/twitter/url/https/github.com/mirinsoft/debotnet.svg?style=flat-square&logo=twitter"
alt="GitHub tweet">
		 
<a href="https://twitter.com/intent/follow?screen_name=Mirinsoft">
 <img src="https://img.shields.io/twitter/follow/Mirinsoft?style=social&logo=twitter"alt="follow on Twitter"></a>

</p>

<p align="center">Be sure to ⭐️ Debotnet's repo so you can keep up to date on releases and progress!</p>

*** 
<h1 align="center">
<sub>
<img  src="https://github.com/Mirinsoft/Debotnet/raw/master/debotnet.png"
      height="38"
      width="38">
</sub>
	Debotnet - Your copy of Windows 10 is now Debotnetted!
	
</h1>

<p align="center">
A free and portable tool for controlling Windows 10's many privacy-related settings and keep your personal data private.
	
![alt text](https://github.com/mirinsoft/debotnet/blob/master/debotnet-intro.gif)

</p>

*** 

# Stop this!
![alt text](https://github.com/mirinsoft/debotnet/blob/master/debotnet-this.png)

# with Debotnet!

The Windows 10 default privacy settings leave a lot to be desired when it comes to protecting you and your private information. Whenever I set up a new computer or update a current setup for my family or job, I always carefully go through the privacy settings for each install, making sure to lock it down to make it as private as possible.

Windows 10 has raised several concerns about privacy due to the fact that it has a lot of telemetry and online features, which send your data (sensitive and not) to Microsoft and can't be disabled, which means Microsoft can:

* Run software on your computer without your consent
* Get data from your computer without your consent
* Remove software and files from your computer without your consent

This qualifies Windows 10 as malware, and more specifically, a botnet.

#### Debotnet requires Windows 10 including both 32-bit and 64-bit versions.

#### You can read a more complete introduction about the latest changes in the blog post [here](https://www.mirinsoft.com/blog/18-debotnet/16-improvements-in-debotnet-release-0-6).

## 🎨 Features

* Disable telemetry and online features, which send your data (sensitive and not) to Microsoft
* Choose which unwanted functions you wish to disable
* Debotnet will show you what it's doing. You have full control, because the executing code can be viewed in a simple text editor
* Simple scripting engine for adding custom privacy rules
* Debug mode. E.g. the Test mode lets you see which values are twisted in registry or commands executed
* Scripts updated on GitHub
* Support for [Ninite](https://ninite.com) service, which allows you to download and install more than 70 popular apps for Windows.
* Download Windows 10 ISO files using Microsoft Media Creation Tool with command line switches, also through [MCT Wrapper](https://gist.github.com/AveYo/c74dc774a8fb81a332b5d65613187b15) or just a [PowerShell script](https://github.com/pbatard/Fido)
* Download deprecated Windows apps, e.g. Classic Calculator
* Modern and familiar UI, with [theme support](https://github.com/mirinsoft/debotnet/blob/master/themes/theme.md)
* Small footprint. No installation required (Portable)

## 🔨 Usage
As above mentioned Debotnet's main tools (in this case the scripts for debotnetting Windows) are not hard coded. Debotnet is based upon simple .DS1 files which define exactly which registry keys, files and or/services should be disabled, blocked, deleted etc. and preserved by the program. These script files allows you to execute command-line tools and parameter and also simple PowerShell code.

## 🐾 Example
```
[Info]
ID=No more forced updates
Ver=1.0
Desc=This will notify when updates are available, and you decide when to install them.\n\nThe values added to registry with this script will prevent forced updates.
Dev=Federico Dossena
DevURL=https://github.com/adolfintel/Windows10-Privacy
WinVer=Compatible with Windows 10
Evaluation=Recommended
EvaluationColor=009e5e

[Code]
Task1=Try,query "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" /v NoAutoUpdate,STDOUT
Task2=Try,query "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" /v AUOptions,STDOUT
Task3=Try,query "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" /v ScheduledInstallDay,STDOUT
Task4=Try,query "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" /v ScheduledInstallTime,STDOUT
File5=Reg,add "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" /v NoAutoUpdate /t REG_DWORD /d 0 /f,STDOUT
File6=Reg,add "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" /v AUOptions /t REG_DWORD /d 2 /f,STDOUT
File7=Reg,add "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" /v ScheduledInstallDay /t REG_DWORD /d 0 /f,STDOUT
File8=Reg,add "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" /v ScheduledInstallTime /t REG_DWORD /d 3 /f,STDOUT
```

## 💾 Download 
* (Latest stable on Dev site) https://www.mirinsoft.com/debotnet
* (Latest stable on GitHub) https://github.com/mirinsoft/debotnet/releases
* (Chocolatey release) https://chocolatey.org/packages/debotnet
* (Nightly builds) https://www.mirinsoft.com/debotnet-nightly
* (Latest script packages) https://www.mirinsoft.com/debotnet-packs

## 💬 Suggesting features / providing feedback
Please send me [feedback](https://www.mirinsoft.com/feedback) or report an [Issue on GitHub](https://github.com/mirinsoft/debotnet/issues/new)

## 🚀 Acknowledgments
* More Information about Debotnet can be found on the [official website](https://www.mirinsoft.com/ms-apps/debotnet)
* Several scripts are based upon the Windows 10 privacy guide from Federico Dossena, https://github.com/adolfintel/Windows10-Privacy
* Icon by EatosDesign, https://www.deviantart.com/eatosdesign/art/Windows-9-Icons-359505915

## 📜 License
100% Free, also for commercial purposes

<br>

<p align="center">
	<strong>
		<a href="https://www.mirinsoft.com">Website</a>
		•
		<a href="https://github.com/mirinsoft/debotnet/tree/master/scripts">Scripts</a>
		•
		<a href="https://github.com/mirinsoft/debotnet/wiki/How-to's">How-to's</a>
		•
		<a href="https://github.com/mirinsoft/debotnet/wiki/FAQ">FAQ</a>
		•
		<a href="https://github.com/mirinsoft/debotnet/wiki">Documentation</a>
		•
		<a href="https://github.com/mirinsoft/debotnet/releases">Download</a>
	</strong>
</p>


Built with ❤︎ by © [Mirinsoft](https://www.mirinsoft.com "Good apps are as little apps as possible").
