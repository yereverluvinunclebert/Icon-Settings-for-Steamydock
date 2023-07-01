# Rocketdock-Enhanced-Settings-VB6

ICON SETTINGS for Steamydock, written in VB6. A WoW64 dock settings
utility for Reactos, XP, Win7, 8 and 10+.

This utility is for use with SteamyDock, is Beta-grade software, under
development, not yet ready to use on a production system - use at your
own risk.

This utility is a functional reproduction of the original settings screen that came from Rocketdock. Please note that the design is limited to enhancing what Rocketdock already provides in order to make the utility familiar to Rocketdock users. If you hover your mouse cursor on the various components that comprise the utility a tooltip will appear that will give more information on each item. There is a help button on the bottom right that will provide further detail at any time. Presing CTRL+H will give you an instant HELP pop up.

The tool is designed to operate with SteamyDock, the open source replacement for Rocketdock. SteamyDock is a work in progress so please bear that in mind when any reference to SteamyDock is made in this documentation.

![rocketdock-embeddedIcons](https://github.com/yereverluvinunclebert/rocketdock/assets/2788342/513587e5-4a15-4c24-8e2c-a2c6fe34b1d5)

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
           Elroys for the balloon tooltips

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
           Registry code ALLAPI.COM
           La Volpe  http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=67466&lngWId=1
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

The program runs without any additional Microsoft plugins.

Krools replacement for the Microsoft Windows Common Controls found in
mscomctl.ocx (treeview, slider) are replicated by the addition of two
dedicated OCX files that are shipped with this package.

           CCRImageList.ocx
           CCRSlider.ocx
           CCRTreeView.ocx

           These OCX will reside in the same folder as the utility that uses it.

           OLEGuids.tlb

           This is a type library that defines types, object interfaces, and more specific API definitions
           needed for COM interop / marshalling. It is only used at design time (IDE). This is a Krool-modified
           version of the original .tlb from the vbaccelerator website. The .tlb is compiled into the executable.
           For the compiled .exe this is not a dependency.

           From the command line, copy the tlb to a central location and register it.

           COPY OLEGUIDS.TLB %SystemRoot%\System32\
           REGTLIB %SystemRoot%\System32\OLEGUIDS.TLB

Building a Manifest:
Using La Volpes program

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

![windowsGeneralIcons](https://github.com/yereverluvinunclebert/rocketdock/assets/2788342/b5730f8c-f8d9-4007-930b-f398a41450d9)

