#include <IE.au3>
#include <Excel.au3>
#include <Array.au3>
#include <MsgBoxConstants.au3>

;open Excel Workbook
Local $oExcel = _Excel_Open()
Local $sWorkbook = @ScriptDir & "\.com Logins.xlsx"
Local $oWorkbook = _Excel_BookOpen($oExcel, $sWorkbook, True, True)

;declare iterative variable
Local $i = Random(1, 15, 1)
Local $alreadyDone[100]

;wait 5 sec
sleep(5000)

;declare local variable for webpage, create webpage
Global $oIE = _IECreate ("http://.com/Account/Login")

;For loop to login and log out
For $x = 1 To 25 Step +1

;set tenancy, username, password fields to data from Excel
Local $tenancyName = _Excel_RangeRead($oWorkbook, Default, "A"&$i)
Local $username = _Excel_RangeRead($oWorkbook, Default, "B"&$i)
Local $password = _Excel_RangeRead($oWorkbook, Default, "C"&$i)

Sleep(5000)

;Sign in
Call ("signIn")

;wait some time
sleep(Random(5000, 10000, 1))

;sign out
call ("signOut")
sleep (2000)

;add i to alreadyDone array
$alreadyDone[$x] = $i

;randomize iterative variable
Do

$i = Random(1,15,1)

Until _ArraySearch($alreadyDone, $i) = -1 ;until you find an i that is not done yet

Next

Sleep(5000)
$oExcel.quit



;sign in to the webpage
Func signIn ()

Local $WEBtenancyName = _IEGetObjByName ($oIE,"tenancyName")
Local $WEBusername = _IEGetObjByName ($oIE,"usernameOrEmailAddress")
Local $WEBpassword = _IEGetObjByName ($oIE,"password")

_IEFormElementSetValue ($WEBtenancyName, $tenancyName)
_IEFormElementSetValue ($WEBusername, $username)
_IEFormElementSetValue ($WEBpassword, $password)

_IEAction ($WEBpassword, "focus")
Sleep(2000)

$colTags = _IETagNameGetCollection($oIE, "button")
For $oTag In $colTags
    If $oTag.classname = "btn btn-success uppercase" Then
        _IEAction($oTag, "click")
    EndIf
Next

EndFunc

;sign out of the webpage
Func signOut ()

_IENavigate($oIE, "http://cloudbusinessdesk.com/Account/Logout")

EndFunc

;go to Login page
Func goToLogin ()

_IENavigate($oIE, "http://cloudbusinessdesk.com/Account/Login")

EndFunc

;function to randomize iterative variable
Func randomize ($i)

$i = Random(1,15,1)

EndFunc



;TO-DO

; implement random time the user is logged in for
; implement random time the user logs in at
; make sure already logged in before users don't log in again
; use dynamic boundaries for excel sheet size and random function
