#Include G:\My Drive\Historic Images\Chrome.ahk
#Warn  ; Enable warnings to assist with detecting common errors.
#Persistent
#SingleInstance Off
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

global page := Chrome.getPageByTitle("batch_ocr.gs - Project Editor - Apps Script")
toggle = 0

$+^1::
If (toggle := !toggle)
	SetTimer, PressTheKey, -1
Return

PressTheKey:
while toggle
{
	page.evaluate("document.querySelector('#yDmH0d > c-wiz > div > div.Hu42fb > div.OX2gTc.DqVtX.dLSs8b.p7Awzb.fb0g6 > div.Kp2okb.lFgX0b > div > div > div.Kp2okb.SQyOec > div > div > div.voS0mf.EF5LSd > div > div:nth-child(3) > div:nth-child(1) > div > div > button > span').click();")
	Sleep, 280000
}
Return