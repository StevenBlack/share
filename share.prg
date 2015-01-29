*  Program...........: share.prg
*  Author............: Steven M Black
*  Project...........: Tools
*  Version...........: Build 7 (May 9 1998)
*  Created...........: 02/19/98  19:34:35
*) Description.......: VFP Class Browser Addin
*)                   : for extracting components into
*)                   : a new VCX.
*  Calling Samples...:
*  Parameter List....: oBrowser: Reference to the calling
*                    : instance of the class browser.
* -------------------: ----------------------------------------
*  ToDo:             : Reset the original class browser status
*  ToDo:             : Hide interim Browsers
*  ToDo:             : Copy and repointer icons, .H files
* -------------------: ----------------------------------------

LPARAMETERS oBrowser

*********************************************
*-- Registration services start
IF TYPE("oBrowser") = "L"
  LOCAL lcName
  lcName = "Steven Black's Component Extractor"
  IF TYPE( "_oBrowser" ) = "O"
    _oBrowser.Addin( lcName, STRTRAN( SYS(16), ".FXP", ".PRG" ))
  ELSE
    IF FILE( HOME() + "BROWSER.DBF" )
      USE ( HOME() + "BROWSER" ) AGAIN SHARED ALIAS StevenBlacksComponentExtractor
      LOCATE FOR Type = "ADDIN" AND NAME = lcName
      IF EOF()
        APPEND BLANK
      ENDIF
      REPLACE Type WITH "ADDIN", ;
        Id WITH "MENU", ;
        Name WITH lcName, ;
        Program WITH LOWER( STRTRAN( SYS(16), ".FXP", ".PRG"))
      USE
    ENDIF
  ENDIF
  RETURN
ENDIF
*-- Registration services end
*********************************************

*-- Trick-or-treat test to ensure we've really been
*-- passed a class browser
IF !PEMSTATUS( oBrowser, "lFileMode", 5)
  RETURN .F.
ENDIF

*-- No support for copying whole files
IF oBrowser.lFileMode
  RETURN .F.
ENDIF

*-- Save environmental things
LOCAL oProcess, oEnvironment
oEnvironment = CREATE( "SaveEnvironment" )
SET EXCLUSIVE OFF

*-- Here's the meat of the matter
oProcess= CREATE( "ProcessServices", oBrowser )
oProcess.ShareClass()


*==================================================
DEFINE CLASS ProcessServices AS CUSTOM
*==================================================
oBrowser           = .NULL.
oClassSvc          = .NULL.
cClassSvcClass     = "ClassServices"
oHierarchySvc      = .NULL.
cHierarchySvcClass = "HierarchyServices"

*--------------------------------------------
FUNCTION Init( toBrowser)
*--------------------------------------------
this.oBrowser      = toBrowser
this.oClassSvc     = CREATE( this.cClassSvcClass,     toBrowser)
this.oHierarchySvc = CREATE( this.cHierarchySvcClass, toBrowser)

*--------------------------------------------
FUNCTION ShareClass()
*--------------------------------------------
*-- Export the selected class
*-- Where?
LOCAL lcVcx, loTarget
lcVcx = this.oClassSvc.GetTarget()
IF !".vcx" $ LOWER( lcVcx)
  RETURN .F.
ENDIF

*-- Create the VCX, bail if we can't
IF ! this.oClassSvc.CreateCLassLibrary( lcVcx)
  WAIT WINDOW "Canot create class library " + lcVcx
  RETURN .F.
ENDIF

this.oHierarchySvc.TraversePedigree( this.oBrowser, lcVcx)

*-- We now have the basic class hierarchy loaded
*-- in the browser, but not including the hierarchy
*-- of its members.
LOCAL llFoundNew
DO WHILE .T.
  SELECT 0
  USE ( lcVcx ) AGAIN
  COPY TO Junk
  USE Junk
  llFoundNew = .F.
  SCAN FOR ! EMPTY ( ClassLoc ) AND ;
      LOWER( ALLTRIM( this.oBrowser.TrimPath( ClassLoc )))<> LOWER( ALLTRIM( this.oBrowser.TrimPath( lcVcx )))

    llFoundNew = .T.
    *-- Open a new browser
    DO ( _BROWSER) WITH ClassLoc, Class
    _SCREEN.ActiveForm.Caption = "Interim Browser -- Please wait"
    *-- Add Class if it isn't already therein
    this.oHierarchySvc.TraversePedigree( _SCREEN.ActiveForm, lcVcx )
    _SCREEN.ActiveForm.Release
  ENDSCAN
  USE IN Junk
  ERASE Junk.DBF
  IF !llFoundNew
    EXIT
  ENDIF
ENDDO

DO (_BROWSER ) WITH lcVCX
_SCREEN.ActiveForm.Backcolor  = rgb( 255, 100, 100 )
_SCREEN.ActiveForm.Caption    = "Component Export Results"
_SCREEN.ActiveForm.AutoCenter = .T.

ENDDEFINE

*==================================================
DEFINE CLASS ClassServices AS Custom
*==================================================
oBrowser = .NULL.

*--------------------------------------------
FUNCTION Init( toBrowser )
*--------------------------------------------
this.oBrowser = toBrowser

*--------------------------------------------
FUNCTION CreateCLassLibrary( tcLibName )
*--------------------------------------------
* Create a class library.
IF !FILE( tcLibName )
  CREATE CLASSLIB ( tcLibName )
ENDIF
RETURN FILE( tcLibName )

*--------------------------------------------
FUNCTION GetTarget()
*--------------------------------------------
* Get the target class from the user.
PRIVATE oReturn
oReturn = CREATE( "CharacterParameter" )
LOCAL oGetFileForm
oGetFileForm = CREATE( "TargetVCXform", oReturn )
oGetFileForm.SHOW()
RETURN oReturn.Caption

ENDDEFINE

*==================================================
DEFINE CLASS HierarchyServices AS Custom
*==================================================
oBrowser = .NULL.

*--------------------------------------------
FUNCTION Init( toBrowser )
*--------------------------------------------
this.oBrowser = toBrowser

*--------------------------------------------
FUNCTION TraversePedigree( toBrowser, tcVCX )
*--------------------------------------------
SET ESCAPE ON
*-- Remember the Browser pointer
LOCAL lcClass, oB
lcClass = toBrowser.cClass
DO WHILE .T.

  IF USED( tcVCX )
    USE IN ( tcVCX )
  ENDIF
  ADD CLASS ( toBrowser.cClass ) OF ( toBrowser.cFileName ) TO ( tcVCX ) OVERWRITE

  SELECT 0
  USE ( tcVCX ) AGAIN SHARED
  SCAN FOR Class = toBrowser.cClass
    REPLACE ClassLoc WITH LOWER( this.oBrowser.TrimPath( tcVCX ))
  ENDSCAN
  USE
  IF ! toBrowser.SeekParentClass()
    EXIT
  ENDIF
ENDDO

*-- Restore the Browser poiner
toBrowser.FindClass(lcClass )

*--------------------------------------------
FUNCTION ERROR( nIndex, nError, cMethod, nLine )
*--------------------------------------------
* Ignore these errors!
DO CASE
  CASE nIndex = 1948 && Cannot modify a class that is in use
OTHERWISE
  this.oBrowser.ERROR( nIndex, nError, cMethod, nLine )
ENDCASE

ENDDEFINE


**** SUPPORT STUFF BELOW ************************
DEFINE CLASS SaveEnvironment AS CUSTOM
cExclusive = ''

FUNCTION Init()
this.cExclusive = SET("Exclusive")

FUNCTION Destroy()
LOCAL lcExcl
lcExcl = this.cExclusive
SET EXCLUSIVE &lcExcl

ENDDEFINE

DEFINE CLASS TargetVCXform AS form
*-- Form used to specify where to export this component
Height      = 104
Width       = 300
AutoCenter  = .T.
BorderStyle = 2
Caption     = "Where to export this class?"
FontName    = "MS Sans Serif"
FontSize    = 8
WindowType  = 1
Name        = "FrmVCXWhere"
oParameter  = .NULL.

ADD OBJECT lbl1 AS Baselabel WITH ;
Caption = "Specify a VCX wherein to export the class hierarchy", ;
Left    = 10, ;
Top     = 11, ;
Width   = 249

ADD OBJECT txtvcx AS FileTextbox WITH ;
Left  = 10, ;
Top   = 41, ;
Width = 248

ADD OBJECT cmdsearch AS VCXSearchCommandbutton WITH ;
Top     = 41, ;
Left    = 266, ;
Width   = 23, ;
Caption = "..."

ADD OBJECT cmdok AS OKCommandButton WITH ;
Top     = 71, ;
Left    = 115, ;
Width   = 84, ;
Caption = "OK", ;
Enabled = .F.

ADD OBJECT cmdcancel AS CancelCommandButton WITH ;
Top     = 71, ;
Left    = 205, ;
Width   = 84, ;
Caption = "Cancel"

FUNCTION INIT( toPassed )
this.oParameter = toPassed

ENDDEFINE

DEFINE CLASS Baselabel AS Label
FontName = "MS Sans Serif"
FontSize = 8
Height   = 23
Name     = "lbl1"
ENDDEFINE

DEFINE CLASS FileTextbox AS TextBox
FontName = "MS Sans Serif"
FontSize = 8
Height   = 23
Name     = "txtVCX"

FUNCTION Valid()
thisform.cmdok.Enabled= !EMPTY( this.Value )
this.Parent.oParameter.SetReturn( this.Value )

FUNCTION KeyPress()
LPARAMETERS nKeyCode, nShiftAltCtrl
this.Valid()

ENDDEFINE

DEFINE CLASS BaseCommandButton AS CommandButton
Height   = 23
FontName = "MS Sans Serif"
FontSize = 8
ENDDEFINE

DEFINE CLASS VCXSearchCommandbutton AS BaseCommandButton
Width   = 23
Caption = "..."

FUNCTION Click()
CD CURDIR()
thisform.txtvcx.Value  = LOWER( GETFILE("VCX"))
thisform.cmdok.Enabled = !EMPTY( thisform.txtvcx.Value )

ENDDEFINE

DEFINE CLASS OKCommandButton AS BaseCommandButton
Width   = 84
Caption = "OK"
Default = .T.

FUNCTION Click()
LOCAL lcValue
lcValue = ALLT( thisform.txtvcx.Value)
IF ATC( ".", lcValue ) =0
  lcValue = lcValue + ".vcx"
ENDIF
thisform.oParameter.SetReturn( lcValue )
thisform.Release

ENDDEFINE

DEFINE CLASS CancelCommandButton AS BaseCommandButton
Width   = 84
Caption = "Cancel"

FUNCTION Click()
thisform.oParameter.SetReturn( '' )
thisform.Release
ENDDEFINE

DEFINE CLASS CharacterParameter AS Label
FUNCTION SetReturn( tcValue )
IF TYPE( "tcValue" ) = "C"
  this.Caption = tcValue
ENDIF
ENDDEFINE

