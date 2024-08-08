# Steamydock-Enhanced-Settings-VB6

ICON SETTINGS for Steamydock, written in VB6. A WoW64 dock settings
utility for Reactos, XP, Win7, 8 and 10+.

![rocketdock-embeddedIcons](https://github.com/yereverluvinunclebert/rocketdock/assets/2788342/a525e0e1-50fc-42c9-8cb5-d578e3a9efaf)

This utility is for use with SteamyDock, is Beta-grade software, under
development, not yet ready to use on a production system - use at your
own risk.

This utility is a functional reproduction of the original settings screen that 
came from Rocketdock. Please note that the design is limited to enhancing what 
Rocketdock already provides in order to make the utility familiar to Rocketdock 
users. If you hover your mouse cursor on the various components that comprise 
the utility a tooltip will appear that will give more information on each item. 
There is a help button on the bottom right that will provide further detail at 
any time. Presing CTRL+H will give you an instant HELP pop up.

![lowContrasr](https://github.com/yereverluvinunclebert/rocketdock/assets/2788342/8fee79a9-bb0a-4338-bc83-e251ba6de562)

The tool is designed to operate with SteamyDock, the open source replacement for 
Rocketdock. SteamyDock is a work in progress so please bear that in mind when 
any reference to SteamyDock is made in this documentation.

This tool was developed on Windows 7 using 32 bit VisualBasic 6 as a FOSS 
project creating a WoW64 program for the desktop. 

It is open source to allow easy configuration, bug-fixing, enhancement and 
community contribution towards free-and-useful VB6 utilities that can be created
by anyone. The first step was the creation of this template program to form the 
basis for the conversion of other desktop utilities or widgets. A future step 
is conversion to RADBasic/TwinBasic for future-proofing and 64bit-ness. 

This utility is one of a set of steampunk and dieselpunk desktop widgets. That 
you can find here on Deviantart: https://www.deviantart.com/yereverluvinuncleber/gallery

I do hope you enjoy using this utility and others. Your own software 
enhancements and contributions will be gratefully received if you choose to 
contribute.

Credits : Standing on the shoulders of the following giants:

	LA Volpe (VB Forums) for his transparent picture handling.  
	Shuja Ali (codeguru.com) for his settings.ini code.  
	KillApp code from an unknown, untraceable source, possibly on MSN.  
	Registry reading code from ALLAPI.COM.

	Punklabs for the original inspiration and for Rocketdock, Skunkie in particular.  
	Active VB Germany for information on the undocumented PrivateExtractIcons API.  
	Elroy on VB forums for his Persistent debug window  
	Rxbagain on codeguru for his Open File common dialog code without dependent OCX  
	Krool on the VBForums for his impressive common control replacements  
	si_the_geek for his special folder code  
	KPD-Team for the code to trawl a folder recursively KPDTeam@Allapi.net http://www.allapi.net  
	Elroy on the VBForums for the balloon tooltips  
	Rod Stephens vb-helper.com Resize controls to fit when a form resizes  
	KPD-Team 1999 http://www.allapi.net/ Recursive search  
	IT researcher https://www.vbforums.com/showthread.php?784053-Get-installed-programs-list-both-32-and-64-bit-programs  
	For the idea of extracting the ununinstall keys from the registry  
	CREDIT Jacques Lebrun http://www.vb-helper.com/howto_get_shortcut_info.html

Built using: VB6, MZ-TOOLS 3.0, CodeHelp Core IDE Extender Framework 2.2 & Rubberduck 2.4.1

	MZ-TOOLS https://www.mztools.com/  
	CodeHelp http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=62468&lngWId=1  
	Rubberduck http://rubberduckvba.com/  
	Rocketdock https://punklabs.com/  
	VBAdvance https://classicvb.net/tools/vbAdvance/  
	Registry code ALLAPI.COM  
	La Volpe http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=67466&lngWId=1  
	PrivateExtractIcons code http://www.activevb.de/rubriken/  
	Persistent debug code http://www.vbforums.com/member.php?234143-Elroy  
	Open File common dialog code without dependent OCX - http://forums.codeguru.com/member.php?92278-rxbagain  
	Open font dialog code without dependent OCX - unknown URL  
	Krools replacement Controls http://www.vbforums.com/showthread.php?698563-CommonControls-%28Replacement-of-the-MS-common-controls%29

Tested on :  

	ReactOS 0.4.14 32bit on virtualBox  
	Windows 7 Professional 32bit on Intel  
	Windows 7 Ultimate 64bit on Intel  
	Windows 7 Professional 64bit on Intel  
	Windows XP SP3 32bit on Intel  
	Windows 10 Home 64bit on Intel  
	Windows 10 Home 64bit on AMD  
	Windows 11 64bit on Intel

Dependencies:

o A windows-alike o/s such as Windows XP, 7-11 or ReactOS.

o Microsoft VB6 IDE installed with its runtime components. The program runs 
without any additional Microsoft OCX components, just the basic controls that 
ship with VB6.

o Requires the SteamyDock program source code to be downloaded and available in 
an adjacent folder as some of the BAS modules are common and shared.

Example folder structure:
	
	E:\VB6\steamydock    ! from https://github.com/yereverluvinunclebert/SteamyDock
	E:\VB6\docksettings	! from https://github.com/yereverluvinunclebert/dockSettings
	E:\VB6\rocketdock		! this repo.

o Krool's replacement for the Microsoft Windows Common Controls found in
mscomctl.ocx (treeview, slider) are replicated by the addition of three
dedicated OCX files that are shipped with this package.

During development these should be copied to C:\windows\syswow64 and should be registered.

- CCRImageList.ocx
- CCRSlider.ocx
- CCRTreeView.ocx

Register these using regsvr32, ie. in a CMD window with administrator privileges.
- c:
- cd \windows\syswow64
- regsvr32 CCRImageList.ocx

Do the same for all three OCX. This will allow the custom controls to be accessible to the VB6 IDE
at design time and the sliders, treeview and imagelist will function as intended (if these ocx are
not registered correctly then the relevant controls will be replaced by picture boxes).

No need to do the above at runtime. At runtime these OCX will reside in the program folder. The program reference to these
OCX is contained within the supplied resource file, IconSettings.RES. The reference to these 
files is compiled into the binary. As long as the three OCX are in the same folder as the binary
the program will run without need to register the three OCX.

o OLEGuids.tlb

This is a type library that defines types, object interfaces, and more specific
API definitions needed for COM interop / marshalling. It is only used at design
time (IDE). This is a Krool-modified version of the original .tlb from the
vbaccelerator website. The .tlb is compiled into the executable.
For the compiled .exe this is NOT a dependency, only during design time.

From the command line, copy the tlb to a central location (system32 or wow64
folder) and register it.

COPY OLEGUIDS.TLB %SystemRoot%\System32\
REGTLIB %SystemRoot%\System32\OLEGUIDS.TLB

In the VB6 IDE - project - references - browse - select the OLEGuids.tlb

Project References:  

	VisualBasic for Applications  
	VisualBasic Runtime Objects and Procedures  
	VisualBasic Objects and Procedures  
	OLE Automation - drag and drop  
	Microsoft Shell Controls and Automation

LICENCE AGREEMENTS:

Copyright 2023 Dean Beedell

In addition to the GNU General Public Licence please be aware that you may use
any of my own imagery in your own creations but commercially only with my
permission. In all other non-commercial cases I require a credit to the
original artist using my name or one of my pseudonyms and a link to my site.
With regard to the commercial use of incorporated images, permission and a
licence would need to be obtained from the original owner and creator, ie. me.

![desktop1](https://github.com/yereverluvinunclebert/rocketdock/assets/2788342/f2d3be1e-c98f-4597-9c8d-503486cf5afb)
