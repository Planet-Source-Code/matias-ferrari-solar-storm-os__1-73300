-----------------------------------------------------------------------
     Copyright © 1997 Microsoft Corporation. All rights reserved.
You have a royalty-free right to use, modify, reproduce and distribute
the Sample Application Files (and/or any modified version) in any way
    you find useful, if you agree that Microsoft has no warranty,
     obligations or liability for any Sample Application Files.
-----------------------------------------------------------------------

                         Visual Basic Dbgwproc.dll
		(Debug Object for AddressOf Subclassing)
                              June 13, 1997
1. Introduction:

Subclassing is a technique that enables you to intercept Windows 
messages being sent to a form or control. By intercepting these 
messages, you can then write your own code to change or extend the 
behavior of the object. VB5 provides the AddressOf keyword, which can 
be used to reroute Windows messages to your own message processing 
procedure.  Subclassing using the AddressOf is very efficient, but 
makes debugging a project difficult.  If the window you are 
subclassing receives a message when you are in break mode, VB will 
crash. The DbgWProc.Dll (Debug Object for AddressOf Subclassing) 
enables you to debug normally while a subclass is active without 
adding any unnecessary overhead to your finished product or 
distributing an extra component.  

For a more detailed discussion of the risks and rewards of subclassing,
and for additional sample code, please refer to the Books Online topic 
"Passing Function Pointers to DLL Procedures and Type Libraries"


NOTE - Before jumping into subclassing with AddressOf
=====================================================
This is an advanced programming technique.  It is very easy to crash VB 
using these methods.  In particular, the Stop button and End statement 
are completely off limits for projects with an active subclass.  It is 
your responsibility as a programmer to correctly turn off the subclass 
before VB shuts down.  

Advantages
==========
Rolling your own subclass has the performance and distribution 
advantages mentioned earlier, as well as the ability to process 
messages at any time.  Traditional, event based subclassing techniques 
required with earlier versions of VB have the disadvantage of being 
unable to process messages while a modal form is showing over the 
subclassed window.  For example, if you're relying on receiving 
WM_PAINT or WM_DRAWITEM messages to properly display a custom window, 
then you'll have to use AddressOf directly, or a subclassing component 
which relies on an Implements based direct callback model instead of 
the more common Event model.

Using DbgWProc.dll:
===================
To use the DbgWProc.Dll to enable debugging:
1) Copy DbgWProc.Dll to your Windows\System[32] or VB5 directory.
2) regsvr32 DbgWProc.Dll
3) Add a reference to 'Debug Object for AddressOf Subclassing'
4) In the Make tab of the Project Properties dialog, add the 
  DEBUGWINDOWPROC=-1 : USEGETPROP=-1 conditional compilation arguments
5) Now include the following code in your class, control, or form 
   module (this is only the parts of the subclassing code which 
   actually change):
'At the top of the file
Private m_wndprcNext As Long
#If DEBUGWINDOWPROC Then
Private m_SCHook As WindowProcHook
#End If

'In the SubClass procedure.  For the finished version, this code is a 
'single call to SetWindowLong
#If DEBUGWINDOWPROC Then
    On Error Resume Next
    Set m_SCHook = CreateWindowProcHook
    If Err Then
        MsgBox Err.Description
        Err.Clear
        UnSubClass
        Exit Sub
    End If
    On Error GoTo 0
    With m_SCHook
        .SetMainProc AddressOf Form1Proc
        m_wndprcNext = SetWindowLong(hWnd, GWL_WNDPROC, .ProcAddress)
        .SetDebugProc m_wndprcNext
    End With
#Else
    m_wndprcNext = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf_
			Form1Proc)
#End If

Each WindowProcHook object provides a ProcAddress value to 
use in SetWindowLong.  You call SetMainProc with AddressOf -
your subclassing procedure - and SetDebugProc with the original 
window procedure.  The WindowProcHook will call the debug 
procedure directly if you're in break mode, skipping the 
subclassing code and the invalid zero return produced by the 
IDE.  You can create up to 100 simulateous WindowProcHook 
objects (hence the error trapping code), so be sure to release 
them when they're not in use.

Samples
=======
SubClass.vbp contains two samples: The AboutBox (simple 
sample which puts an AboutBox item on the system menu) and 
DrawItem (advanced ownerdraw listbox OCX sample).  Both 
give the option of associating a subclassed hWnd with a class 
instance using either SetWindowLong(GWL_USERDATA) or 
SetProp/GetProp.  USEGETPROP=-1 uses the more robust 
SetProp/GetProp method, which also requires the Atomizer.cls 
file.

To view these samples, load them (AboutBox\SubClass.vbp or 
DrawItem\Group1.vbg) and press F5 to run them. To exit, close the 
forms - do not press the Stop button.

To view DbgWProc in action, load the samples and begin single-stepping 
through the code with the F8 key. As an alternative, put a breakpoint 
in Friend Function WindowProc (in Form1 of the AboutBox sample and the 
OwnerDrawListBox usercontrol of the DrawItem sample). Both of these 
samples initiate subclassing with a routine called SubClass and end it 
with a routine called UnSubClass. If you press the Stop button while 
subclassing is active (between SubClass and UnSubClass), VB will GPF.

Note on Compiling after using DbgWproc.dll:
===========================================
Before you compile your final OCX or EXE file, change the 
conditional compilation value for DEBUGWINDOWPROC to 0.  If 
you run a project which uses DbgWProc when the IDE isn't 
active, it will give you a warning error message when the Dll 
loads.  This is meant to strongly discourage you from actually 
shipping the Dll with your project (your users will see the 
same message).  Also be sure to remove the Dll from the list 
distributed files in SetupWizard.

For more information and samples with AddressOf subclass, see 
the Black-Belt column in the February and June '97 issues of the 
Visual Basic Programmer's Journal magazine (www.windx.com).
