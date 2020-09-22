    Make DLL's in Visual Basic!
   (c) 2004 DanSoft Australia.

Here are the instructions to install the addin:

0.5) Unzip the zip file -- I assume this has already been done, so this is step 0.5 instead of 1 ;)
1) Open 'Linker.vbp' (in the 'linker' folder) and compile it.
2) Go into your Visual Basic folder (usually C:\Program Files\Microsoft Visual Studio\VB98) and rename 'LINK.EXE' to 'LINK1.EXE'
3) Copy the 'Link.exe' file (in the 'compiled' folder) to your Visual Basic Folder
4) Open 'MakeDLLAddin.vbp' (in the 'addin' folder) and compile it
5) Go into Visual Basic, and click Add-Ins -> Add-In Manager. There should be an addin listed called 'Make DLL's In Visual Basic'. Make sure both 'Loaded' and 'Load On Startup' are ticked.
5a) If the addin wasen't listed, copy 'MakeDLL.DLL' (in the 'compiled' folder) into your Visual Basic directory and restart Visual Basic.
6) Copy all the files in the 'dll project' folder to your Visual Basic Project Templates folder (usually C:\Program Files\Microsoft Visual Studio\VB98\template\projects)

Yay! It is now installed!
Sample DLL: 'TestDLL.vbp' (in the 'test dll' folder)
Sample prog that uses that DLL: 'TestProg.vbp' (in the 'test program' folder)
If you want to create a DLL yourself, go into Visual Basic and choose to create a 'Standard DLL' project.

If you liked this, please vote! (Even if you didn't like it, post comments saying why).