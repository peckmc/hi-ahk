#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Persistent
#Include G:\My Drive\Historic Images\Chrome.ahk
#Include G:\My Drive\Historic Images\CSV-master\csv.ahk
#SingleInstance,Force
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

CSV_Load("G:\My Drive\Historic Images\aOCR.csv","aOCR")
CSV_Load("G:\My Drive\Historic Images\bOCR.csv","bOCR")

global page := Chrome.getPageByTitle("Historic Images Template")
page.WaitForLoad()

;$^+p::
;  Send ^c
;  page.Evaluate("document.getElementById('title').value += '" clipboard "';")
;  page.Evaluate("document.getElementById('title').focus();")
;  return

$^v::   
        Clipboard := workFormat(Clipboard)
	send ^v
	return

$^+u:: ; set category to "music/singers"
	page.evaluate("document.querySelector('#collapse1 > div > ul > li:nth-child(4) > a').click();")
	return

$^+o:: ; set category to "movie actors/actresses"
	page.evaluate("document.querySelector('#collapse1 > div > ul > li:nth-child(8) > a').click();")
	return

;$^+t::  
;   page.evaluate("document.querySelector('#ui-id-78').click();")
;	return

getImgInfo(partNo) {
	ResultA:=CSV_Search("aOCR", partNo)
	ResultB:=CSV_Search("bOCR", partNo)

	ResultA:=StrSplit(ResultA,",")
	ResultB:=StrSplit(ResultB,",")

	ResultA:=CSV_ReadCell("aOCR",ResultA[1],2)
	ResultB:=CSV_ReadCell("bOCR",ResultB[1],2)

	FoundPos := RegExMatch(ResultA ResultB, "(?i)(Photo credit: |Photo by: )(.*?\s.*?)(\/|\s|\.)", Match)
	if (FoundPos > 0) {
	   Match2 := RegExReplace(Match2, "'", "\'")
	   page.evaluate("document.querySelector('#photographer').value = '" Match2 "';")
	}

	if (InStr(ResultA ResultB,"ABC Television") | InStr(ResultA ResultB,"ABC Photograph.")) {
	   handleCreditSelection("ABC (American Broadcasting Company", ")")
	} else if InStr(ResultA ResultB,"CBS") {
		handleCreditSelection("CBS (Columbia Broadcasting System", ")")
	} else if (InStr(ResultA ResultB,"NBC") | InStr(ResultB,"NBC Photo")) {
		handleCreditSelection("NBC (National Broadcasting Company", ")")
	}
	return
}

handleCreditSelection(credit, endChar) {
	page.evaluate("document.querySelector('#credits').focus();")
	page.evaluate("document.querySelector('input#search_credits_autocomplete.form-control.ui-autocomplete-input').value = '';")
	page.evaluate("document.querySelector('input#search_credits_autocomplete.form-control.ui-autocomplete-input').value = '" credit "';")
	page.evaluate("document.querySelector('input#search_credits_autocomplete.form-control.ui-autocomplete-input').setSelectionRange(1000,1000);")
	Sleep,500
	Send %endChar%
	Sleep,1000
	SendInput {enter}
	page.evaluate("document.getElementById('date_month').focus();")
	return
}

$^+f:: ; save caption to file and click "save image"
	title := page.evaluate("document.getElementById('title').value").value
	desc := page.evaluate("document.getElementById('description').value").value

	category := page.evaluate("document.getElementById('search_credits_autocomplete').value").value
	fileAppend,%desc%`r`n`r`n,G:\My Drive\Historic Images\AllClips.txt
	page.evaluate("document.getElementById('save_next_image_for_review').click();")
	sleep,3000
   partNo := page.evaluate("document.querySelector('#partnumber').innerHTML").value 
   getImgInfo(partNo)
;	page.evaluate("document.querySelector('#collapse1 > div > ul > li:nth-child(4) > a').click();")
	return

$^+s:: ; don't save caption, click "save image"
	page.evaluate("document.getElementById('save_next_image_for_review').click();")
	sleep,1500
   partNo := page.evaluate("document.querySelector('#partnumber').innerHTML").value
   getImgInfo(partNo)

;	page.evaluate("document.querySelector('#select_credits_autocomplete').setAttribute('value','Houston Post');")
;	page.evaluate("document.querySelector('#select_credits_autocomplete > button').click();")
;	page.evaluate("document.getElementById('date_month').focus();")
	return

;$^+p::
;	imgid := page.evaluate("document.querySelector('#partnumber').innerHTML").value
;	bimgid := imgid . "b"
;	pre := SubStr(imgid, 1, 3)
;	FileRead, ocrText, C:\Users\thinkpoop\Documents\test images\%pre%\%imgid%.txt
;	ocrText := escape(ocrText)
;	ocrText := workFormat(ocrText)
;	ocrText := formatForSavedFile(ocrText)
;	page.Evaluate("document.getElementById('description').value += '" ocrText "'")
;	return

;OnClipboardChange:
;	sugArray := []
;	pst := 1
;	clip := clipboard
;	Loop, {
;		found_pos := RegExMatch(clip,"O)((?:[A-Z][a-z]*\s?){2,})",sug,pst)
;		pst:=found_pos+StrLen(sug)+1
;		if not sug
;			break
;		sugArray.push(sug.1)
;	}
;	Gui,destroy
;	for phr,but in sugArray
;		Gui, Add, Button, y+2 vButton1, % but
;	Gui, +AlwaysOnTop
;	Gui, Show, x1415 y870 w500 h150,Suggested Captions
;	pastePhr(but) {
;		page.Evaluate("document.getElementById('title').value += '" but "'")
;	return
;	}

escape(Str) {
   static escapeChars := { 8: ""        ; Backspace is replaced with \b
                        ,  9: ""        ; Horizontal Tab is replaced with \t
                        , 10: ""        ; Newline is replaced with \n
                        , 11: ""        ; Horizontal Tab is replaced with \t
                        , 12: ""        ; Form feed is replaced with \f
                        , 13: ""        ; Carriage return is replaced with \r
                        , 34: "\"""       ; Double quote is replaced with \"
                        , 39: "\'"}       ; Single quote is replaced with \'
   escaped := StrReplace(str, "\", "\\")  ; Backslash is replaced with \\ first
   for i, v in escapeChars
      escaped := StrReplace(escaped, chr(i), v)
   return escaped
   }

formatForSavedFile(originalString) {
	originalString := RegExReplace(originalString, "^.*(?i)(warner bros\.|images).*$","")
	return originalString
}

workFormat(Str) {
	StringReplace,str,str,`n,,A ; remove line breaks
	StringReplace,str,str,`r,%A_Space%,A ; same as above
	StringReplace,str,str,\s,%A_Space%,A ; ""
	StringReplace, str, str,”,",A 
	StringReplace, str, str,”,",A 
	StringReplace, str, str,“,",A
	StringReplace, str, str,’,',A
	StringReplace, str, str,‘,',A
	StringReplace, str, str,*,,A
	StringReplace, str, str,^,,A
	StringReplace, str, str,•,-,A
	StringReplace, str, str,>,i,A
	StringReplace, str, str,<,i,A
	StringReplace, str, str,»,,A
	StringReplace, str, str,|,i,A
	StringReplace, str, str,„,,A
	StringReplace, str, str,¦,,A
	StringReplace, str, str,—,--,A
	StringReplace, str, str,nght,right,A
	StringReplace, str, str,%A_Space%ot%A_Space%,%A_Space%of%A_Space%,A
	StringReplace, str, str,%A_Space%bnng%A_Space%,%A_Space%bring%A_Space%,A
	StringReplace, str, str,%A_Space%trom%A_Space%,%A_Space%from%A_Space%,A
	StringReplace, str, str,%A_Space%m%A_Space%,%A_Space%in%A_Space%,A
	StringReplace, str, str,%A_Space%ol%A_Space%,%A_Space%of%A_Space%,A
	StringReplace, str, str,%A_Space%lather%A_Space%,%A_Space%father%A_Space%,A
	StringReplace, str, str,stamng,starring,A
	StringReplace, str, str,Distnbution,Distribution,A
	StringReplace, str, str,Distnbuted,Distributed,A
	StringReplace, str, str,distnbution,distribution,A
	StringReplace, str, str,distnbuted,distributed,A
	StringReplace, str, str,%A_Space%ln%A_Space%,%A_Space%in%A_Space%,A
	StringReplace, str, str,%A_Space%tor%A_Space%,%A_Space%for%A_Space%,A
	StringReplace, str, str,%A_Space%lhe%A_Space%,%A_Space%the%A_Space%,A
	StringReplace, str, str,{,(,A
	StringReplace, str, str,},),A
	StringReplace, str, str,],),A
	StringReplace, str, str,[,(,A
	StringReplace, str, str,PROGRAM:,in,A
	StringReplace, str, str, thnller,thriller,A
	StringReplace, str, str, scries,series,A
	StringReplace, str, str,Rep.,Representative,A 
	StringReplace, str, str,Gov.,Governor,A 
	StringReplace, str, str,Md.,Maryland,A 
	StringReplace, str, str,Tex.,Texas,A 
	StringReplace, str, str,U.S.,United States,A
	str := RegExReplace(str, "(?i)\bCalif\.", "California")
	str := RegExReplace(str, "(?i)\bCapt\.", "Captain")
	str := RegExReplace(str, "(?i)\bCol\.", "Colonel")
	str := RegExReplace(str, "(?i)\bCpl\.","Corporal")
	str := RegExReplace(str, "(?i)\bMaj\.","Major")
	str := RegExReplace(str, "(?i)\bN\. ?J\.","New Jersey")
	str := RegExReplace(str, "(?i)\bOkla\.","Oklahoma")
	str := RegExReplace(str, "(?i)\bSen\.", "Senator")
	str := RegExReplace(str, "(?i)\bGen\.", "General")
	str := RegExReplace(str, "(?i)\bRev\.", "Reverend")
	str := RegExReplace(str, "(?i)\bLt\.", "Lieutenant")
	str := RegExReplace(str, "(?i)\bCmdr\.", "Commander")
	str := RegExReplace(str, "(?i)\bFla\.", "Florida")
	str := RegExReplace(str, "(?i)\bAdm\.", "Admiral")
;	str := RegExReplace(str, "(,)([^\d\x22])", "$1 $2") ; add spaces after commas as needed
;	str := RegExReplace(str, "(,)( )(\x22)", "$1$3") ; eliminate spaces between periods or commas and quotation marks
	str := RegExReplace(str, "\.{3,}","")
	str := RegExReplace(str, " ,",",")
	str := RegExReplace(str, " \.",".") ; eliminates spaces before periods
	str := RegExReplace(str, "(?i)\bp\.m[^.]","PM ")
	str := RegExReplace(str, "(?i)\bhts\b","his")
  	str := RegExReplace(str, "(?i)\bdistnbutes\b","distributes")
	str := RegExReplace(str, "(?i)\b(s)(enes)\b","$1eries")
	str := RegExReplace(str, "(?i)\b(w)(ile)\b","$1ife")
	str := RegExReplace(str, "(?i)\b(l)(ett)\b","$1eft")
	str := RegExReplace(str, "\b(tn)\b","in")
	;str := RegExReplace(str, "(?i)\b(\(?i\.)","L.")
	str := RegExReplace(str, "(\s\s+)|\t", " ")
	str := RegExReplace(str, "( +)i\t", " ")
	str := RegExReplace(str, "\b([0-9]+)(ST|ND|TH)\b", "$L0")
	str := RegExReplace(str, "(?!ABC|AIDS|AMC|BBC|CBS|CMT|CDT|CIA|CST|EDT|EST|ET|FBI|HBO|KGB|MCA|MGM|MTV|NBA|NBC|NFL|NYT|PBS|PDT|PM|PST|PT|RCA|SFM|TNN|TV|USA|USO|USSR|WGBH|WNET\b)(\b[A-Z]+'?[A-Z]\b)", "$T1")
;	str := RegExReplace(str, "(?<!^|\. |: |['""“])(A|An|And|By|Is|In|Of|On|The|To)[^\.]\b", "$L0")
	str := RegExReplace(str, "(\bMc)([A-Z]{2,}'?[A-Z]\b)", "$1$T2")
	str := RegExReplace(str, "\b((?i)jr|dr|mr|mrs|ms)( )\b", "$T1. ")
	str := RegExReplace(str, "\b((?i)DR\. |JR\. |MR\. |MRS\. |MS\. )\b", "$T0")
	str := RegExReplace(str, "\b(bom)\b", "born")
	str := RegExReplace(str, "\b(Bom)\b", "Born")
	str := RegExReplace(str, "(_+)", " ") ; replace underscores with spaces
	str := RegExReplace(str, " 1n ", " in ")
	str := RegExReplace(str, " 1s ", " is ")
	str := RegExReplace(str, "«|®|■|•", "")
	str := RegExReplace(str, "\b((?i)Warner Bros )\b", "Warner Bros. ")
	str := RegExReplace(str, "'s ", "$L0") ; converts possessive words, i.e. "ROSE'S" to "rose's" instead of "rose'S"
	str := RegExReplace(str, "\b ((?i)i|ii|iii|iv|v)\b", "$U0")
	return str
	}