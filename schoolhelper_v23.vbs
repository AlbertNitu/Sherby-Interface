' DEBATE TOPIC: REPLACE 'CALCULATE' COMMAND WITH WOLFRAM?!?
' THINGS TO MAKE SHERBY LIKE SIRI: SPORTS DATA, MAPS DATA, RESTAURANTS DATA AND RESERVATIONS, MOVIES DATA, WEATHER DATA, STOCK DATA.

' ALWAYS UPDATE COMMAND LIST BOTH FOR THE 'HELP' COMMAND AS WELL AS ON THE WEBSITE!
' CHANGE VERSION NUMBER (IN THE VERSION FUNCTION) FOR EVERY SHERBY UPDATE!
' FOR EACH UPDATE, MAKE SURE TO CHANGE SHORTCUT PATH OF SHERBY.LNK TO THE LATEST VERSION OF THE FILE!


' NEW COMMANDS IN THIS VERSION:
'		VERSION = COMMAND THAT INFORMS THE USER THE CURRENT SHERBY INTERFACE VERSION. USEFUL FOR VERIFYING IF YOU'VE SUCCESSFULLY UPDATED THE SHERBY INTERFACE.
'		WHATEVER/BYE = ** ADDED VOCABULARY FOR THE NEVERMIND/CANCEL FUNCTION. **
'		DISABLE VOICE RECOGNITION = COMMAND THAT DISABLES VOICE RECOGNITION. (YOU CAN ALWAYS TURN VOICE RECOGNITION BACK ON BY TYPING IN THE COMMAND 'VOICE RECOGNITION)
'		SHUTDOWN = COMMAND THAT SHUTS DOWN THE COMPUTER.
'		RESTART/REBOOT = COMMAND THAT RESTARTS/REBOOTS THE COMPUTER
'		NOTEPAD: COMMANDS THAT OPENS A NEW, EMPTY NOTEPAD DOCUMENT
'		FILE EXPLORER: COMMAND THAT OPENS THE FILE EXPLORER TO THE DOCUMENTS FOLDER
'				****************** MAKE SETTING TO CHANGE OPENING PATH! *********************
'		CHROME = COMMAND THAT OPENS UP A NEW CHROME WINDOW
'		CHROME TAB = COMMAND THAT OPENS UP A NEW CHROME TAB IN AN EXISTING CHROME WINDOW. (IF THERE IS NO CHROME WINDOW ALREADY OPEN, IT WILL OPEN THE NEW CHROME TAB IN A NEW WINDOW)
'		FIREFOX = COMMAND THAT OPENS UP A NEW FIREFOX WINDOW
'		FIREFOX TAB = COMMAND THAT OPENS UP A NEW FIREFOX TAB IN AN EXISTING FIREFOX WINDOW (IF THERE IS NO FIREFOX WINDOW ALREADY OPEN, IT WILL OPEN THE NEW FIREFOX TAB IN A NEW WINDOW)
'		CONTROL PANEL = COMMAND THAT OPENS THE CONTROL PANEL FOR YOU
'		WORD = COMMAND THAT TELLS SHERBY TO CREATE A NEW WORD DOCUMENT FOR You
'		EXCEL = COMMAND THAT TELLS SHERBY TO CREATE A NEW EXCEL/SPREADSHEET DOCUMENT
'		POWERPOINT = COMMAND THAT TELLS SHERBY TO CREATE A NEW POWERPOINT/SLIDESHOW DOCUMENT
'		[number of] DEGREES CELSIUS TO FAHRENHEIT = COMMAND THAT ALLOWS YOU TO CONVERT CELSIUS TO FAHRENHEIT
'		[number of] DEGREES FAHRENHEIT TO CELSIUS = COMMAND THAT ALLOWS YOU TO CONVERT FAHRENHEIT TO CELSIUS
'		CALCULATE [insert calculation here] = COMMAND THAT ALLOWS YOU CALCULATE ANYTHING YOU WANT. (NOTE: THE OPERATION FACTORIAL DOESN'T WORK, BUT OTHER COMPLEX OPERATIONS SUCH AS SQUARE ROOT OR TRIGONOMETRIC OPERATIONS DO WORK)
'		FLIP A COIN = COMMAND THAT RETURNS EITHER HEADS OR TAILS
'		ROLL A DIE = SIMULATES ROLLING A NORMAL 6 SIDED DIE
'		ROLL A [number of sides] SIDED DIE = COMMAND THAT RETURNS A RANDOM NUMBER FROM ONE (INCLUSIVE) TO THE NUMBER OF SIDES THE DIE HAS. (ALSO INCLUSIVE) IN OTHER WORDS, IT SIMULATES ROLLING ADIE WITH A SPECIFIED NUMBER OF SIDES.
'		YES OR NO = COMMAND THAT CHOOSES EITHER YES OR NO
'		CONVERT [unit to another unit] = COMMAND THAT (AS LONG AS YOU HAVE THE KEYWORD 'CONVERT' IN YOUR COMMAND) WILL CONVERT PRACTICALLY ANY UNIT TO ANY UNIT (INCLUDING CURRENCY)
'		SEARCH FOR "[INSERT TEXT YOU WANT TO SEARCH HERE]" ON GOOGLE IMAGES = COMMAND THAT GOOGLES THE TEXT FOUND IN BETWEEN THE QUOTATION MARKS IN GOOGLE IMAGES
'		SEARCH FOR "[INSERT TEXT YOU WANT TO SEARCH HERE]" ON YOUTUBE = COMMAND THAT GOOGLES THE TEXT FOUND IN BETWEEN THE QUOTATION MARKS IN GOOGLE IMAGES
'		GOOGLE SEARCH = COMMAND THAT PROMPTS/ASKS YOU FOR WHAT YOU WANT TO SEARCH ON GOOGLE FOR. SHERBY THEN GOOGLES YOUR QUERY ON GOOGLE. (THIS IS SIMILAR TO THE OTHER GOOGLE SEARCH COMMAND, BUT IS USEFUL IF YOU WANT TO SEARCH LONGER TEXTS)
'		GOOGLE IMAGE SEARCH = COMMAND THAT PROMPTS/ASKS YOU FOR WHAT YOU WANT TO SEARCH ON GOOGLE IMAGES FOR. SHERBY THEN GOOGLES YOUR QUERY ON GOOGLE IMAGES. (THIS IS SIMILAR TO THE OTHER GOOGLE IMAGES SEARCH COMMAND, BUT IS USEFUL IF YOU WANT TO SEARCH LONGER TEXTS ON GOOGLE IMAGES)
'		YOUTUBE SEARCH = COMMAND THAT PROMPTS/ASKS YOU FOR WHAT YOU WANT TO SEARCH ON YOUTUBE FOR. SHERBY THEN GOOGLES YOUR QUERY ON YOUTUBE. (THIS IS SIMILAR TO THE OTHER YOUTUBE SEARCH COMMAND, BUT IS USEFUL IF YOU WANT TO SEARCH LONGER TEXTS ON YOUTUBE)
'		DELETE ALL BOOKMARKS = COMMAND THAT DELETES ALL YOUR BOOKMARKS
'		DELETE ALL GOALS = COMMAND THAT DELETES ALL YOUR GOALS
'		CREATE NOTE = COMMAND THAT ALLOWS YOU CREATE A 'NOTE'. A NOTE IS SIMILAR TO A GOAL, IT'S LIKE A MESSAGE, WHERE YOU CAN WRITE ANYTHING YOU WANT. THIS COMMAND IS PARTICULARLY USEFUL IF YOU NEED TO REMEMBER SOMETHING ON THE SPOT. (EX. 'THE CAPITAL OF MONGOLIA IS ULAANBAATAR')
'		LIST NOTES = COMMAND THAT LISTS ALL OF YOUR CURRENT NOTES
'		DELETE NOTE = COMMAND THAT DELETES A SPECIFIED NOTE. TO DELETE A NOTE, YOU NEED TO SPECIFY THE INDEX OF THE NOTE. FOR EXAMPLE, (AFTER VIEWING ALL YOUR CURRENT NOTES) IF YOU WANTED TO DELETE YOUR 3RD NOTE, YOU WOULD SIMPLY HAVE TO ENTER THE NUMBER '3' TO DELETE THE THIRD NOTE.
'		DELETE ALL NOTES = CMMAND THAT DELETES ALL OF YOUR CURRENT NOTES.
'		NEWS = COMMAND THAT DISPLAYS THE CURRENT NEWS (NEWS COMES FROM BBC NEWS)
'		OPEN NEWS = COMMAND THAT OPENS CNN, WHICH DISPLAYS ALL CURRENT WORLD NEWS
'		OPEN BBC NEWS = COMMAND THAT OPENS BBC NEWS IN A NEW TAB
'		OPEN YAHOO NEWS = COMMAND THAT OPENS YAHOO NEWS IN A NEW TAB

Dim masterName
Set Sapi = CreateObject("Sapi.SpVoice")
Set a = createobject("wscript.shell")

Sub Play(SoundFile)
	Set Sound = CreateObject("WMPlayer.OCX")
	Set fsoSound = CreateObject("Scripting.FileSystemObject")
	If (fsoSound.FileExists(SoundFile)) Then
		Sound.URL = SoundFile
		Sound.settings.volume = 100
		Sound.Controls.play
		do while Sound.currentmedia.duration = 0
			wscript.sleep 1
		loop
		wscript.sleep(int(Sound.currentmedia.duration)+1)*1000
	Else
	   wscript.sleep(0)
	End If
End Sub

Function PCase(stringToConvert)
	stringToConvert = LTrim(stringToConvert)
	PCase = UCase(Left(stringToConvert, 1)) & LCase(Mid(stringToConvert, 2))
End Function

Function rand(arrayone)
	maximum = UBound(arrayone)
	minimum = 0
	Randomize
	randomOutcomeRandom=Int((maximum-minimum+1)*Rnd+minimum)
	rand=arrayone(randomOutcomeRandom)
End Function

Function readFromRegistry(strRegistryKey, strDefault)
	On Error Resume Next
	Set WSHShellSetUp = CreateObject("WScript.Shell")
	value = WSHShellSetUp.RegRead(strRegistryKey)
		if err.number <> 0 then
		readFromRegistry= strDefault
	else
		readFromRegistry=value
	end if
		set WSHShell = nothing
End function

Function OpenWithChrome(strURL)
	Dim strChrome
	Dim WShellChrome
	strChrome = readFromRegistry ( "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe\Path", "") 
	if (strChrome = "") then
		strChrome = "chrome.exe"
	else
		strChrome = strChrome & "\chrome.exe"
	end if
	Set WShellChrome = CreateObject("WScript.Shell")
	strChrome = """" & strChrome & """" & " " & strURL
	WShellChrome.Run strChrome, 1, false
End function


Dim suffix
Function convertDate(dateSpec)
	suffix = "th"
	months = Array("HI!", "January", "Febuary", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
	dateSplit = Split(dateSpec, "-")
	month = months(dateSplit(1))
	If dateSplit(0) = 1 OR dateSplit(0) = 21 OR dateSplit(0) = 31 Then
		suffix = "st"
	ElseIf dateSplit(0) = 2 OR dateSplit(0) = 22 Then
		suffix = "nd"
	ElseIf dateSplit(0) = 3 OR dateSplit(0) = 23 Then
		suffix = "nd"
	End If
	Sapi.speak month + " " + CStr(dateSplit(0)) + suffix + ", " + CStr(dateSplit(2))
End Function

'Still working on it!!!!
Function doWolframProfileSearch(commandProfile)
	name = ""
	proffession = ""
	fullName = ""
	birthDate = ""
	placeOfBirth = ""
	deathDate = ""
	age = ""
	profileInfo = ""
	startedSearch = 0
	foundplaneWolfram(commandProfile)
	Dim WinHttpReq
	Set WinHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
	Set objFSO=CreateObject("Scripting.FileSystemObject")
	Dim sTemp
	foundError=0
	foundplaintext=0
	startingURL = "http://api.wolframalpha.com/v2/query?input="
	If isplaneOverhead = 1 Then
		endingURL = "&appid=AAXTLK-5XE3J9E8L3&location=Ottawa,%20Ontario&podindex=2&format=image"
	Else
		endingURL = "&appid=AAXTLK-5XE3J9E8L3&location=Ottawa,%20Ontario&format=plaintext&podindex=1&podindex=2&podtitle=Notable%20facts"
	End If
	commandRefinedXML = commandProfile
	commandRefinedXML = Replace(commandRefinedXML, "+", " plus ")
	commandRefinedXML = Replace(commandRefinedXML, "#", "%23")
	commandRefinedXML = Replace(commandRefinedXML, "%", "%25")
	commandRefinedXML = Replace(commandRefinedXML, "^", "%5E")
	commandRefinedXML = Replace(commandRefinedXML, "&", "%26")
	commandRefinedXML = Replace(commandRefinedXML, "=", "%3D")
	commandRefinedXML = Replace(commandRefinedXML, " ", "+")
	commandRefinedXML = Replace(commandRefinedXML, "\", "%5C")
	commandRefinedXML = Replace(commandRefinedXML, "|", "%7C")
	commandRefinedXML = Replace(commandRefinedXML, "{", "%7B")
	commandRefinedXML = Replace(commandRefinedXML, "{", "%7D")
	commandRefinedXML = Replace(commandRefinedXML, "[", "%5B")
	commandRefinedXML = Replace(commandRefinedXML, "]", "%5D")
	commandRefinedXML = Replace(commandRefinedXML, "'", "%27")
	commandRefinedXML = Replace(commandRefinedXML, "?", "%3F")
	commandRefinedXML = Replace(commandRefinedXML, "/", "%2F")
	fullXMLUrl = startingURL + commandRefinedXML + endingURL
	WinHttpReq.Open "GET", fullXMLURL, False

	WinHttpReq.Send
	 If (WinHttpReq.Status = 200) Then
		  correctString=""
		  sTemp = WinHttpReq.ResponseText

		  sTemp = Replace(sTemp, vbCr, vbcrlf)
		  sTemp = Replace(sTemp, vbLf, vbcrlf)
		 
		   WebText = sTemp

		   Set objFile = objFSO.OpenTextFile("C:\Sherby Interface\xmlWolframFile.txt", 2, True)
		   'on error resume next
		   objFile.WriteLine sTemp
			'If Err Then
			'	WScript.StdErr.WriteLine "error "
			'End If 
		   objFile.Close

			Set objFileXML = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\xmlWolframFile.txt",1)

		If isplaneOverhead = 0 Then
			do while not objFileXML.AtEndOfStream
				plaintextSearch = objFileXML.ReadLine()
				If InStr(1, plaintextSearch, "scanner='Identity'") > 0 Then
					startedSearch = 1
				ElseIf startedSearch = 1 And InStr(1, plaintextSearch, "<plaintext>") > 0 Then
					plaintextSearch = Replace(plaintextSearch, "<plaintext>", "")
					plaintextSearch = Replace(plaintextSearch, "</plaintext>", "")
					If Right(plaintextSearch, 1) = ")" And Left(plaintextSearch, 1) <> "(" Then
						lastOcurrenceOfBracket = InStrRev(plaintextSearch,"(",-1)
						name = Left(plaintextSearch, lastOcurrenceOfBracket - 1)
					Else
						wscript.sleep(0)
					End If
wscript.echo "Name:" + name
				ElseIf startedSearch = 1 And InStr(1, plaintextSearch, "full name | ") > 0 Then
					plaintextSearch = Replace(plaintextSearch, "<plaintext>", "")
					plaintextSearch = Replace(plaintextSearch, "</plaintext>", "")
					plaintextSearch = Replace(plaintextSearch, "full name | ", "")
					fullName = plaintextSearch
wscript.echo "Full name: " + fullName
				ElseIf startedSearch = 1 And InStr(1, plaintextSearch, "date of birth | ") > 0 Then
					plaintextSearch = Replace(plaintextSearch, "<plaintext>", "")
					plaintextSearch = Replace(plaintextSearch, "</plaintext>", "")
					plaintextSearch = Replace(plaintextSearch, "date of birth | ", "")
					If InStr(1, plaintextSearch, "(") > 0 Then
						birthDate = Left(plaintextSearch, InStr(1, plaintextSearch, "(") - 1)
					Else
						birthDate = plaintextSearch						
					End If
wscript.echo "Birthday: " + birthDate
				ElseIf startedSearch = 1 And InStr(1, plaintextSearch, "place of birth | ") > 0 Then
					plaintextSearch = Replace(plaintextSearch, "<plaintext>", "")
					plaintextSearch = Replace(plaintextSearch, "</plaintext>", "")
					plaintextSearch = Replace(plaintextSearch, "place of birth | ", "")
					If InStr(1, plaintextSearch, "(") > 0 Then
						placeOfBirth = Left(plaintextSearch, InStr(1, plaintextSearch, "("))
					Else
						placeOfBirth = plaintextSearch						
					End If
wscript.echo "Place of birth: " + placeOfBirth
				ElseIf startedSearch = 1 And InStr(1, plaintextSearch, "date of death | ") > 0 Then
					plaintextSearch = Replace(plaintextSearch, "<plaintext>", "")
					plaintextSearch = Replace(plaintextSearch, "</plaintext>", "")
					plaintextSearch = Replace(plaintextSearch, "date of death | ", "")
					If InStr(1, plaintextSearch, "(age: ") > 0 Then
						age = Left(Mid(plaintextSearch, InStr(1, plaintextSearch, "(age: ")), InStr(1, plaintextSearch, ")"))
					Else
						age = ""
					End If
					If InStr(1, plaintextSearch, "(") > 0 Then
						deathDate = Left(plaintextSearch, InStr(1, plaintextSearch, "("))
					Else
						deathDate = plaintextSearch						
					End If
wscript.echo "Date of death: " + deathDate + " ..... age: "
				ElseIf InStr(1, plaintextSearch, "<plaintext>") > 0 Then
					plaintextSearch = Replace(plaintextSearch, "<plaintext>", "")
					plaintextSearch = Replace(plaintextSearch, "</plaintext>", "")
					If Trim(plaintextSearch) = "" OR plaintextSearch = " " OR plaintextSearch = "  " OR plaintextSearch = "   " OR plaintextSearch="(data not available)" Then
						Sapi.speak Replace(rand(Array("Okay, here's what I found for, #%command%#, on google", "Okay, here's what I found on the web for, #%command%#.", "Okay, I found this on the web for, #%command%#", "Okay, check it out. Here's what I found for, #%command%#, on google", "Okay, check it out. Here's what I found for, #%command%#, on the web", "I found this on google for, #%command%#", "Here's what I found for, #%command%#, on the web")), "#%command%#", commandProfile)
						googleSearchQueryRefined=Trim(commandProfile)
						googleSearchQueryRefined=Replace(googleSearchQueryRefined, "+", "%2B")
						googleSearchQueryRefined=Replace(googleSearchQueryRefined, "#", "%23")
						googleSearchQueryRefined=Replace(googleSearchQueryRefined, "%", "%25")
						googleSearchQueryRefined=Replace(googleSearchQueryRefined, "&", "%26")
						googleSearchQueryRefined=Replace(googleSearchQueryRefined, " ", "%20")
						startingurl="https://www.google.ca/?gws_rd=cr&ei=acGMV629B8TQ-QH6vZroDQ&safe=active&ssui=on#safe=active&q="
						fullurl=startingurl+googleSearchQueryRefined
						a.run fullurl
						foundError = 1
					Else
						plaintextSplitByLetter = Array()
						x=Len(plaintextSearch)
						for i=0 to x
							letterToAdd = Mid(plaintextSearch,i+1,1)
							ReDim Preserve plaintextSplitByLetter(UBound(plaintextSplitByLetter) + 1)
							plaintextSplitByLetter(UBound(plaintextSplitByLetter)) = letterToAdd
						Next
						plaintextSearch = Replace(plaintextSearch, "au (astronomical units)", "astronomical units")
						plaintextSearch = Replace(plaintextSearch, "^", " to the power of ")
						plaintextSearch = Replace(plaintextSearch, "|", ",")
						plaintextSearch = Replace(plaintextSearch, "(degrees celsius)", "")
						plaintextSearch = Replace(plaintextSearch, "(meters per second)", "")
						plaintextSearch = Replace(plaintextSearch, "m/s", " meters per second")
						plaintextSearch = Replace(plaintextSearch, " N ", " north ")
						plaintextSearch = Replace(plaintextSearch, " S ", " south ")
						plaintextSearch = Replace(plaintextSearch, " E ", " east ")
						plaintextSearch = Replace(plaintextSearch, " W ", " west ")
						plaintextSearch = Replace(plaintextSearch, " NE ", " north east ")
						plaintextSearch = Replace(plaintextSearch, " NW ", " north west ")
						plaintextSearch = Replace(plaintextSearch, " SE ", " south east ")
						plaintextSearch = Replace(plaintextSearch, " SW ", " south west ")
						plaintextSearch = Replace(plaintextSearch, " NNE ", " north north east ")
						plaintextSearch = Replace(plaintextSearch, " NNW ", " north north west ")
						plaintextSearch = Replace(plaintextSearch, " ENE ", " east north east ")
						plaintextSearch = Replace(plaintextSearch, " WNW ", " west north west ")
						plaintextSearch = Replace(plaintextSearch, " SSE ", " south south east ")
						plaintextSearch = Replace(plaintextSearch, " SSW ", " south south west ")
						plaintextSearch = Replace(plaintextSearch, " ESE ", " east south east ")
						plaintextSearch = Replace(plaintextSearch, " WSW ", " west south west ")
						If Right(plaintextSearch, 1) = ")" And Left(plaintextSearch, 1) <> "(" Then
							lastOcurrenceOfBracket = InStrRev(plaintextSearch,"(",-1)
							plaintextSearch = Left(plaintextSearch, lastOcurrenceOfBracket - 1)
						Else
							wscript.sleep(0)
						End If
						Dim day
						Dim daySplit
						day = CStr(plaintextSearch)
						daySplit = Split(day, "-")
						daySplitNum = UBound(daySplit)
						If InStr(1, day, "-") > 0 And daySplitNum = 2 Then
							If isNumeric(daySplit(0)) And isNumeric(daySplit(1)) And isNumeric(daySplit(2)) Then
								convertDate(day)
							Else
								If UBound(Split(plaintextSearch, " ")) < 15 Then
									If UBound(Split(plaintextSearch, " ")) < 10 Then
										Sapi.speak plaintextSearch
									Else
										Set VObj = CreateObject("SAPI.SpVoice")

										With VObj
											.Volume = 100
											.Rate = -3
											.Speak plaintextSearch
										End With
									End If
								End If
								If UBound(Split(plaintextSearch, " ")) > 10 Then
									plaintextSearch = LTrim(plaintextSearch)
									wscriptMessageAnswer = msgbox("Answer: " + PCase(plaintextSearch),,"Answer to your question")
								End If
							End If
						Else
							If UBound(Split(plaintextSearch, " ")) < 15 Then
								If UBound(Split(plaintextSearch, " ")) < 10 Then
									Sapi.speak plaintextSearch
								Else
									Set VObj = CreateObject("SAPI.SpVoice")

									With VObj
										.Volume = 100
										.Rate = -1
										.Speak plaintextSearch
									End With
								End If
							End If
							If UBound(Split(plaintextSearch, " ")) > 10 Then
								plaintextSearch = LTrim(plaintextSearch)
								wscriptMessageAnswer = msgbox("ANSWER: " & vbcrlf & "" & vbcrlf & PCase(plaintextSearch),,"Answer to your question")
							End If
						End If
					End If
					foundplaintext = 1
				ElseIf InStr(1, plaintextSearch, "didyoumean") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<tip ") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<error>") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<plaintext/>") > 0 And foundError=0 Then
					Sapi.speak Replace(rand(Array("Okay, here's what I found for, #%command%#, on google", "Okay, here's what I found on the web for, #%command%#.", "Okay, I found this on the web for, #%command%#", "Okay, check it out. Here's what I found for, #%command%#, on google", "Okay, check it out. Here's what I found for, #%command%#, on the web", "I found this on google for, #%command%#", "Here's what I found for, #%command%#, on the web")), "#%command%#", commandProfile)
					googleSearchQueryRefined=Trim(commandProfile)
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "+", "%2B")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "#", "%23")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "%", "%25")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "&", "%26")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, " ", "%20")
					startingurl="https://www.google.ca/?gws_rd=cr&ei=acGMV629B8TQ-QH6vZroDQ&safe=active&ssui=on#safe=active&q="
					fullurl=startingurl+googleSearchQueryRefined
					a.run fullurl
					foundError = 1
					foundplaintext = 1
				End If
			loop
			objFileXML.Close
			If foundplaintext=0 Then
				Sapi.speak Replace(rand(Array("Okay, here's what I found for, #%command%#, on google", "Okay, here's what I found on the web for, #%command%#.", "Okay, I found this on the web for, #%command%#", "Okay, check it out. Here's what I found for, #%command%#, on google", "Okay, check it out. Here's what I found for, #%command%#, on the web", "I found this on google for, #%command%#", "Here's what I found for, #%command%#, on the web")), "#%command%#", commandProfile)
				googleSearchQueryRefined=Trim(commandProfile)
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "+", "%2B")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "#", "%23")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "%", "%25")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "&", "%26")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, " ", "%20")
				startingurl="https://www.google.ca/?gws_rd=cr&ei=acGMV629B8TQ-QH6vZroDQ&safe=active&ssui=on#safe=active&q="
				fullurl=startingurl+googleSearchQueryRefined
				a.run fullurl
			End If
		Else
			do while not objFileXML.AtEndOfStream
				plaintextSearch = objFileXML.ReadLine()
				If InStr(1, plaintextSearch, "<img src=") > 0 Then
					plaintextSearch = Replace(plaintextSearch, "<img src='", "")
					quotationPosition = InStr(1, plaintextSearch, "'")
					imageURL = CStr(Left(plaintextSearch, quotationPosition - 1))
					imageURL = Replace(imageURL, "amp;", "")
					Sapi.speak "Okay, here's a table of all the airplanes currently above us"
					Set objExplorer = CreateObject("InternetExplorer.Application")
					With objExplorer
						.Visible = 1
						.Toolbar=False
						.Statusbar=False
						.Top=800
						.Left=800
						.Height=400
						.Width=360
						.Navigate imageURL
					End With
					foundplaintext = 1
				ElseIf InStr(1, plaintextSearch, "didyoumean") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<tip ") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<error>") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<plaintext/>") > 0 And foundError=0 Then
					Sapi.speak Replace(rand(Array("Okay, here's what I found for, #%command%#, on google", "Okay, here's what I found on the web for, #%command%#.", "Okay, I found this on the web for, #%command%#", "Okay, check it out. Here's what I found for, #%command%#, on google", "Okay, check it out. Here's what I found for, #%command%#, on the web", "I found this on google for, #%command%#", "Here's what I found for, #%command%#, on the web")), "#%command%#", commandProfile)
					googleSearchQueryRefined=Trim(commandProfile)
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "+", "%2B")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "#", "%23")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "%", "%25")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "&", "%26")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, " ", "%20")
					startingurl="https://www.google.ca/?gws_rd=cr&ei=acGMV629B8TQ-QH6vZroDQ&safe=active&ssui=on#safe=active&q="
					fullurl=startingurl+googleSearchQueryRefined
					a.run fullurl
					foundError = 1
					foundplaintext = 1
				End If
			loop
			objFileXML.Close
			If foundplaintext=0 Then
				Sapi.speak Replace(rand(Array("Okay, here's what I found for, #%command%#, on google", "Okay, here's what I found on the web for, #%command%#.", "Okay, I found this on the web for, #%command%#", "Okay, check it out. Here's what I found for, #%command%#, on google", "Okay, check it out. Here's what I found for, #%command%#, on the web", "I found this on google for, #%command%#", "Here's what I found for, #%command%#, on the web")), "#%command%#", commandProfile)
				googleSearchQueryRefined=Trim(commandProfile)
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "+", "%2B")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "#", "%23")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "%", "%25")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "&", "%26")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, " ", "%20")
				startingurl="https://www.google.ca/?gws_rd=cr&ei=acGMV629B8TQ-QH6vZroDQ&safe=active&ssui=on#safe=active&q="
				fullurl=startingurl+googleSearchQueryRefined
				a.run fullurl
			End If
		End If

	  Else
		 Sapi.speak "Unfortunately, I'm unable to answer your question at the moment. Please make sure that your internet connection is secure and that airplane mode is not turned on"
	End If
End Function

Function doWolframMorseSearch(commandMorse)
	foundplaneWolfram(commandMorse)
	Dim WinHttpReq
	Set WinHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
	Set objFSO=CreateObject("Scripting.FileSystemObject")
	Dim sTemp
	foundError=0
	foundplaintext=0
	startingURL = "http://api.wolframalpha.com/v2/query?input="
	If isplaneOverhead = 1 Then
		endingURL = "&appid=AAXTLK-5XE3J9E8L3&location=Ottawa,%20Ontario&podindex=2&format=image"
	Else
		endingURL = "&appid=AAXTLK-5XE3J9E8L3&location=Ottawa,%20Ontario&podindex=2&format=plaintext"
	End If
	commandRefinedXML = commandMorse
	commandRefinedXML = Replace(commandRefinedXML, "+", " plus ")
	commandRefinedXML = Replace(commandRefinedXML, "#", "%23")
	commandRefinedXML = Replace(commandRefinedXML, "%", "%25")
	commandRefinedXML = Replace(commandRefinedXML, "^", "%5E")
	commandRefinedXML = Replace(commandRefinedXML, "&", "%26")
	commandRefinedXML = Replace(commandRefinedXML, "=", "%3D")
	commandRefinedXML = Replace(commandRefinedXML, " ", "+")
	commandRefinedXML = Replace(commandRefinedXML, "\", "%5C")
	commandRefinedXML = Replace(commandRefinedXML, "|", "%7C")
	commandRefinedXML = Replace(commandRefinedXML, "{", "%7B")
	commandRefinedXML = Replace(commandRefinedXML, "{", "%7D")
	commandRefinedXML = Replace(commandRefinedXML, "[", "%5B")
	commandRefinedXML = Replace(commandRefinedXML, "]", "%5D")
	commandRefinedXML = Replace(commandRefinedXML, "'", "%27")
	commandRefinedXML = Replace(commandRefinedXML, "?", "%3F")
	commandRefinedXML = Replace(commandRefinedXML, "/", "%2F")
	fullXMLUrl = startingURL + commandRefinedXML + endingURL
	WinHttpReq.Open "GET", fullXMLURL, False

	WinHttpReq.Send
	 If (WinHttpReq.Status = 200) Then
		  correctString=""
		  sTemp = WinHttpReq.ResponseText

		  sTemp = Replace(sTemp, vbCr, vbcrlf)
		  sTemp = Replace(sTemp, vbLf, vbcrlf)
		 
		   WebText = sTemp

		   Set objFile = objFSO.OpenTextFile("C:\Sherby Interface\xmlWolframFile.txt", 2, True)
		   on error resume next
		   objFile.WriteLine sTemp
			If Err Then
				WScript.StdErr.WriteLine "error "
			End If 
		   objFile.Close

			Set objFileXML = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\xmlWolframFile.txt",1)

		If isplaneOverhead = 0 Then
			do while not objFileXML.AtEndOfStream
				plaintextSearch = objFileXML.ReadLine()
				If InStr(1, plaintextSearch, "<plaintext>") > 0 Then
					plaintextSearch = Replace(plaintextSearch, "<plaintext>", "")
					plaintextSearch = Replace(plaintextSearch, "</plaintext>", "")
					If Trim(plaintextSearch) = "" OR plaintextSearch = " " OR plaintextSearch = "  " OR plaintextSearch = "   " OR plaintextSearch="(data not available)" Then
						Sapi.speak Replace(rand(Array("Okay, here's what I found for, #%command%#, on google", "Okay, here's what I found on the web for, #%command%#.", "Okay, I found this on the web for, #%command%#", "Okay, check it out. Here's what I found for, #%command%#, on google", "Okay, check it out. Here's what I found for, #%command%#, on the web", "I found this on google for, #%command%#", "Here's what I found for, #%command%#, on the web")), "#%command%#", commandMorse)
						googleSearchQueryRefined=Trim(commandMorse)
						googleSearchQueryRefined=Replace(googleSearchQueryRefined, "+", "%2B")
						googleSearchQueryRefined=Replace(googleSearchQueryRefined, "#", "%23")
						googleSearchQueryRefined=Replace(googleSearchQueryRefined, "%", "%25")
						googleSearchQueryRefined=Replace(googleSearchQueryRefined, "&", "%26")
						googleSearchQueryRefined=Replace(googleSearchQueryRefined, " ", "%20")
						startingurl="https://www.google.ca/?gws_rd=cr&ei=acGMV629B8TQ-QH6vZroDQ&safe=active&ssui=on#safe=active&q="
						fullurl=startingurl+googleSearchQueryRefined
						a.run fullurl
						foundError = 1
					Else
						plaintextSplitByLetter = Array()
						x=Len(plaintextSearch)
						for i=0 to x
							letterToAdd = Mid(plaintextSearch,i+1,1)
							ReDim Preserve plaintextSplitByLetter(UBound(plaintextSplitByLetter) + 1)
							plaintextSplitByLetter(UBound(plaintextSplitByLetter)) = letterToAdd
						Next
						plaintextSearch = Replace(plaintextSearch, "au (astronomical units)", "astronomical units")
						plaintextSearch = Replace(plaintextSearch, "^", " to the power of ")
						plaintextSearch = Replace(plaintextSearch, "|", ",")
						plaintextSearch = Replace(plaintextSearch, "(degrees celsius)", "")
						plaintextSearch = Replace(plaintextSearch, "(meters per second)", "")
						plaintextSearch = Replace(plaintextSearch, "m/s", " meters per second")
						plaintextSearch = Replace(plaintextSearch, " N ", " north ")
						plaintextSearch = Replace(plaintextSearch, " S ", " south ")
						plaintextSearch = Replace(plaintextSearch, " E ", " east ")
						plaintextSearch = Replace(plaintextSearch, " W ", " west ")
						plaintextSearch = Replace(plaintextSearch, " NE ", " north east ")
						plaintextSearch = Replace(plaintextSearch, " NW ", " north west ")
						plaintextSearch = Replace(plaintextSearch, " SE ", " south east ")
						plaintextSearch = Replace(plaintextSearch, " SW ", " south west ")
						plaintextSearch = Replace(plaintextSearch, " NNE ", " north north east ")
						plaintextSearch = Replace(plaintextSearch, " NNW ", " north north west ")
						plaintextSearch = Replace(plaintextSearch, " ENE ", " east north east ")
						plaintextSearch = Replace(plaintextSearch, " WNW ", " west north west ")
						plaintextSearch = Replace(plaintextSearch, " SSE ", " south south east ")
						plaintextSearch = Replace(plaintextSearch, " SSW ", " south south west ")
						plaintextSearch = Replace(plaintextSearch, " ESE ", " east south east ")
						plaintextSearch = Replace(plaintextSearch, " WSW ", " west south west ")
						plaintextSearch = Replace(plaintextSearch, "  ", " ")
						If Right(plaintextSearch, 1) = ")" And Left(plaintextSearch, 1) <> "(" Then
							lastOcurrenceOfBracket = InStrRev(plaintextSearch,"(",-1)
							plaintextSearch = Left(plaintextSearch, lastOcurrenceOfBracket - 1)
						Else
							wscript.sleep(0)
						End If
						Dim day
						Dim daySplit
						day = CStr(plaintextSearch)
						daySplit = Split(day, "-")
						daySplitNum = UBound(daySplit)
						If InStr(1, day, "-") > 0 And daySplitNum = 2 Then
							If isNumeric(daySplit(0)) And isNumeric(daySplit(1)) And isNumeric(daySplit(2)) Then
								convertDate(day)
							Else
								Sapi.speak rand("Alright, here's your text in morse code...", "Okay, I've translated your text into morse code.", "Alright, i've translated your desired text into morse code. Here it is...")
								morseCodeResult = msgbox(plaintextSearch,,"Morse code translation")
							End If
						Else
							Sapi.speak rand("Alright, here's your text in morse code...", "Okay, I've translated your text into morse code.", "Alright, i've translated your desired text into morse code. Here it is...")
							morseCodeResult = msgbox(plaintextSearch,,"Morse code translation")
						End If
					End If
					foundplaintext = 1
				ElseIf InStr(1, plaintextSearch, "didyoumean") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<tip ") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<error>") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<plaintext/>") > 0 And foundError=0 Then
					Sapi.speak Replace(rand(Array("Okay, here's what I found for, #%command%#, on google", "Okay, here's what I found on the web for, #%command%#.", "Okay, I found this on the web for, #%command%#", "Okay, check it out. Here's what I found for, #%command%#, on google", "Okay, check it out. Here's what I found for, #%command%#, on the web", "I found this on google for, #%command%#", "Here's what I found for, #%command%#, on the web")), "#%command%#", commandMorse)
					googleSearchQueryRefined=Trim(commandMorse)
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "+", "%2B")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "#", "%23")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "%", "%25")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "&", "%26")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, " ", "%20")
					startingurl="https://www.google.ca/?gws_rd=cr&ei=acGMV629B8TQ-QH6vZroDQ&safe=active&ssui=on#safe=active&q="
					fullurl=startingurl+googleSearchQueryRefined
					a.run fullurl
					foundError = 1
					foundplaintext = 1
				End If
			loop
			objFileXML.Close
			If foundplaintext=0 Then
				Sapi.speak Replace(rand(Array("Okay, here's what I found for, #%command%#, on google", "Okay, here's what I found on the web for, #%command%#.", "Okay, I found this on the web for, #%command%#", "Okay, check it out. Here's what I found for, #%command%#, on google", "Okay, check it out. Here's what I found for, #%command%#, on the web", "I found this on google for, #%command%#", "Here's what I found for, #%command%#, on the web")), "#%command%#", commandMorse)
				googleSearchQueryRefined=Trim(commandMorse)
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "+", "%2B")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "#", "%23")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "%", "%25")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "&", "%26")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, " ", "%20")
				startingurl="https://www.google.ca/?gws_rd=cr&ei=acGMV629B8TQ-QH6vZroDQ&safe=active&ssui=on#safe=active&q="
				fullurl=startingurl+googleSearchQueryRefined
				a.run fullurl
			End If
		Else
			do while not objFileXML.AtEndOfStream
				plaintextSearch = objFileXML.ReadLine()
				If InStr(1, plaintextSearch, "<img src=") > 0 Then
					plaintextSearch = Replace(plaintextSearch, "<img src='", "")
					quotationPosition = InStr(1, plaintextSearch, "'")
					imageURL = CStr(Left(plaintextSearch, quotationPosition - 1))
					imageURL = Replace(imageURL, "amp;", "")
					Sapi.speak "Okay, here's a table of all the airplanes currently above us"
					Set objExplorer = CreateObject("InternetExplorer.Application")
					With objExplorer
						.Visible = 1
						.Toolbar=False
						.Statusbar=False
						.Top=800
						.Left=800
						.Height=400
						.Width=360
						.Navigate imageURL
					End With
					foundplaintext = 1
				ElseIf InStr(1, plaintextSearch, "didyoumean") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<tip ") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<error>") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<plaintext/>") > 0 And foundError=0 Then
					Sapi.speak Replace(rand(Array("Okay, here's what I found for, #%command%#, on google", "Okay, here's what I found on the web for, #%command%#.", "Okay, I found this on the web for, #%command%#", "Okay, check it out. Here's what I found for, #%command%#, on google", "Okay, check it out. Here's what I found for, #%command%#, on the web", "I found this on google for, #%command%#", "Here's what I found for, #%command%#, on the web")), "#%command%#", commandMorse)
					googleSearchQueryRefined=Trim(commandMorse)
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "+", "%2B")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "#", "%23")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "%", "%25")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "&", "%26")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, " ", "%20")
					startingurl="https://www.google.ca/?gws_rd=cr&ei=acGMV629B8TQ-QH6vZroDQ&safe=active&ssui=on#safe=active&q="
					fullurl=startingurl+googleSearchQueryRefined
					a.run fullurl
					foundError = 1
					foundplaintext = 1
				End If
			loop
			objFileXML.Close
			If foundplaintext=0 Then
				Sapi.speak Replace(rand(Array("Okay, here's what I found for, #%command%#, on google", "Okay, here's what I found on the web for, #%command%#.", "Okay, I found this on the web for, #%command%#", "Okay, check it out. Here's what I found for, #%command%#, on google", "Okay, check it out. Here's what I found for, #%command%#, on the web", "I found this on google for, #%command%#", "Here's what I found for, #%command%#, on the web")), "#%command%#", commandMorse)
				googleSearchQueryRefined=Trim(commandMorse)
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "+", "%2B")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "#", "%23")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "%", "%25")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "&", "%26")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, " ", "%20")
				startingurl="https://www.google.ca/?gws_rd=cr&ei=acGMV629B8TQ-QH6vZroDQ&safe=active&ssui=on#safe=active&q="
				fullurl=startingurl+googleSearchQueryRefined
				a.run fullurl
			End If
		End If

	  Else
		 Sapi.speak "Unfortunately, I'm unable to answer your question at the moment. Please make sure that your internet connection is secure and that airplane mode is not turned on"
	End If
End Function

Function doWolframBookSearch(commandBook)
	Dim WinHttpReq
	Set WinHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
	Set objFSO=CreateObject("Scripting.FileSystemObject")
	Dim sTemp
	foundError=0
	foundplaintext=0
	startingURL = "http://api.wolframalpha.com/v2/query?input="
	endingURL = "&appid=AAXTLK-5XE3J9E8L3&location=Ottawa,%20Ontario&podindex=3&format=image"
	commandRefinedXML = commandBook
	commandRefinedXML = Replace(commandRefinedXML, "+", " plus ")
	commandRefinedXML = Replace(commandRefinedXML, "#", "%23")
	commandRefinedXML = Replace(commandRefinedXML, "%", "%25")
	commandRefinedXML = Replace(commandRefinedXML, "^", "%5E")
	commandRefinedXML = Replace(commandRefinedXML, "&", "%26")
	commandRefinedXML = Replace(commandRefinedXML, "=", "%3D")
	commandRefinedXML = Replace(commandRefinedXML, " ", "+")
	commandRefinedXML = Replace(commandRefinedXML, "\", "%5C")
	commandRefinedXML = Replace(commandRefinedXML, "|", "%7C")
	commandRefinedXML = Replace(commandRefinedXML, "{", "%7B")
	commandRefinedXML = Replace(commandRefinedXML, "{", "%7D")
	commandRefinedXML = Replace(commandRefinedXML, "[", "%5B")
	commandRefinedXML = Replace(commandRefinedXML, "]", "%5D")
	commandRefinedXML = Replace(commandRefinedXML, "'", "%27")
	commandRefinedXML = Replace(commandRefinedXML, "?", "%3F")
	commandRefinedXML = Replace(commandRefinedXML, "/", "%2F")
	fullXMLUrl = startingURL + commandRefinedXML + endingURL
	WinHttpReq.Open "GET", fullXMLURL, False

	WinHttpReq.Send
	 If (WinHttpReq.Status = 200) Then

		  correctString=""
		  sTemp = WinHttpReq.ResponseText

		  sTemp = Replace(sTemp, vbCr, vbcrlf)
		  sTemp = Replace(sTemp, vbLf, vbcrlf)
		 
		   WebText = sTemp

		   Set objFile = objFSO.OpenTextFile("C:\Sherby Interface\xmlWolframFile.txt", 2, True)
		   on error resume next
		   objFile.WriteLine sTemp
			If Err Then
				WScript.StdErr.WriteLine "error "
			End If 
		   objFile.Close

			Set objFileXML = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\xmlWolframFile.txt",1)

			do while not objFileXML.AtEndOfStream
				plaintextSearch = objFileXML.ReadLine()
				If InStr(1, plaintextSearch, "<img src=") > 0 And foundplaintext=0 Then
					plaintextSearch = Replace(plaintextSearch, "<img src='", "")
					quotationPosition = InStr(1, plaintextSearch, "'")
					imageURL = CStr(Left(plaintextSearch, quotationPosition - 1))
					imageURL = Replace(imageURL, "amp;", "")
					Sapi.speak "Okay, here's a table of all the famous books written by your specified author"
					Set objExplorer = CreateObject("InternetExplorer.Application")
					With objExplorer
						.Visible = 1
						.Toolbar=False
						.Statusbar=False
						.Top=800
						.Left=800
						.Height=700
						.Width=720
						.Navigate imageURL
					End With
					foundplaintext = 1
				ElseIf InStr(1, plaintextSearch, "didyoumean") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<tip ") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<error>") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<plaintext/>") > 0 And foundError=0 Then
					Sapi.speak Replace(rand(Array("Okay, here's what I found for, #%command%#, on google", "Okay, here's what I found on the web for, #%command%#.", "Okay, I found this on the web for, #%command%#", "Okay, check it out. Here's what I found for, #%command%#, on google", "Okay, check it out. Here's what I found for, #%command%#, on the web", "I found this on google for, #%command%#", "Here's what I found for, #%command%#, on the web")), "#%command%#", commandBook)
					googleSearchQueryRefined=Trim(commandBook)
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "+", "%2B")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "#", "%23")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "%", "%25")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "&", "%26")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, " ", "%20")
					startingurl="https://www.google.ca/?gws_rd=cr&ei=acGMV629B8TQ-QH6vZroDQ&safe=active&ssui=on#safe=active&q="
					fullurl=startingurl+googleSearchQueryRefined
					a.run fullurl
					foundError = 1
					foundplaintext = 1
				End If
			loop
			objFileXML.Close
			If foundplaintext=0 Then
				Sapi.speak Replace(rand(Array("Okay, here's what I found for, #%command%#, on google", "Okay, here's what I found on the web for, #%command%#.", "Okay, I found this on the web for, #%command%#", "Okay, check it out. Here's what I found for, #%command%#, on google", "Okay, check it out. Here's what I found for, #%command%#, on the web", "I found this on google for, #%command%#", "Here's what I found for, #%command%#, on the web")), "#%command%#", commandBook)
				googleSearchQueryRefined=Trim(commandBook)
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "+", "%2B")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "#", "%23")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "%", "%25")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "&", "%26")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, " ", "%20")
				startingurl="https://www.google.ca/?gws_rd=cr&ei=acGMV629B8TQ-QH6vZroDQ&safe=active&ssui=on#safe=active&q="
				fullurl=startingurl+googleSearchQueryRefined
				a.run fullurl
			End If

	  Else
		 Sapi.speak "Unfortunately, I'm unable to answer your question at the moment. Please make sure that your internet connection is secure and that airplane mode is not turned on"
	End If
End Function

Function doWolframSynonymSearch(commandSynonym)
	foundplaneWolfram(command)
	Dim WinHttpReq
	Set WinHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
	Set objFSO=CreateObject("Scripting.FileSystemObject")
	Dim sTemp
	foundError=0
	foundplaintext=0
	startingURL = "http://api.wolframalpha.com/v2/query?input="
	If isplaneOverhead = 1 Then
		endingURL = "&appid=AAXTLK-5XE3J9E8L3&location=Ottawa,%20Ontario&podindex=2&format=image"
	Else
		endingURL = "&appid=AAXTLK-5XE3J9E8L3&location=Ottawa,%20Ontario&podindex=2&format=plaintext"
	End If
	commandRefinedXML = commandSynonym
	commandRefinedXML = Replace(commandRefinedXML, "+", " plus ")
	commandRefinedXML = Replace(commandRefinedXML, "#", "%23")
	commandRefinedXML = Replace(commandRefinedXML, "%", "%25")
	commandRefinedXML = Replace(commandRefinedXML, "^", "%5E")
	commandRefinedXML = Replace(commandRefinedXML, "&", "%26")
	commandRefinedXML = Replace(commandRefinedXML, "=", "%3D")
	commandRefinedXML = Replace(commandRefinedXML, " ", "+")
	commandRefinedXML = Replace(commandRefinedXML, "\", "%5C")
	commandRefinedXML = Replace(commandRefinedXML, "|", "%7C")
	commandRefinedXML = Replace(commandRefinedXML, "{", "%7B")
	commandRefinedXML = Replace(commandRefinedXML, "{", "%7D")
	commandRefinedXML = Replace(commandRefinedXML, "[", "%5B")
	commandRefinedXML = Replace(commandRefinedXML, "]", "%5D")
	commandRefinedXML = Replace(commandRefinedXML, "'", "%27")
	commandRefinedXML = Replace(commandRefinedXML, "?", "%3F")
	commandRefinedXML = Replace(commandRefinedXML, "/", "%2F")
	fullXMLUrl = startingURL + commandRefinedXML + endingURL
	WinHttpReq.Open "GET", fullXMLURL, False

	WinHttpReq.Send
	 If (WinHttpReq.Status = 200) Then

		  correctString=""
		  sTemp = WinHttpReq.ResponseText

		  sTemp = Replace(sTemp, vbCr, vbcrlf)
		  sTemp = Replace(sTemp, vbLf, vbcrlf)
		 
		   WebText = sTemp

		   Set objFile = objFSO.OpenTextFile("C:\Sherby Interface\xmlWolframFile.txt", 2, True)
		   on error resume next
		   objFile.WriteLine sTemp
If Err Then
    WScript.StdErr.WriteLine "error "
End If 
		   objFile.Close

			Set objFileXML = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\xmlWolframFile.txt",1)

		If isplaneOverhead = 0 Then
			do while not objFileXML.AtEndOfStream
				plaintextSearch = objFileXML.ReadLine()
				If InStr(1, plaintextSearch, "<plaintext>") > 0 Then
					plaintextSearch = Replace(plaintextSearch, "<plaintext>", "")
					plaintextSearch = Replace(plaintextSearch, "</plaintext>", "")
					If Trim(plaintextSearch) = "" OR plaintextSearch = " " OR plaintextSearch = "  " OR plaintextSearch = "   " Then
						Sapi.speak "Unfortunately, a list of synonyms for your requested word isn't available."
						foundError = 1
					Else
						If InStr(1, plaintextSearch, "|") > 0 Then
							plaintextSearchSplit = Split(plaintextSearch, "|")
							plaintextArrayLength = UBound(plaintextSearchSplit)
							If plaintextArrayLength < 3 Then
								plaintextSearch = Join(plaintextSearchSplit, ",")
								Sapi.speak plaintextSearch
								plaintextSearch = Replace(plaintextSearch, "  ,", ",")
								plaintextSearch = Replace(plaintextSearch, " ,", ",")
								plaintextSearch = Replace(plaintextSearch, "   ", " ")
								plaintextSearch = Replace(plaintextSearch, "  ", " ")
								If InStr(1, commandSynonym, "synonym") > 0 Then
									synonymResult = msgbox("Synonyms: " + PCase(plaintextSearch),,"Synonyms")
								ElseIf InStr(1, commandSynonym, "antonym") > 0 Then
									synonymResult = msgbox("Antonyms: " + PCase(plaintextSearch),,"Antonyms")
								End If
							Else
								plaintextSearch = CStr(PCase(plaintextSearchSplit(0))) + ", " + CStr(plaintextSearchSplit(1)) + ", " + CStr(plaintextSearchSplit(2)) + ", and " + CStr(plaintextSearchSplit(3))
								Sapi.speak plaintextSearch
								plaintextSearch = Replace(plaintextSearch, "  ,", ",")
								plaintextSearch = Replace(plaintextSearch, " ,", ",")
								plaintextSearch = Replace(plaintextSearch, "   ", " ")
								plaintextSearch = Replace(plaintextSearch, "  ", " ")
								If InStr(1, commandSynonym, "synonym") > 0 Then
									synonymResult = msgbox("Synonyms: " + PCase(plaintextSearch),,"Synonyms")
								ElseIf InStr(1, commandSynonym, "antonym") > 0 Then
									synonymResult = msgbox("Antonyms: " + PCase(plaintextSearch),,"Antonyms")
								End If
							End If
						Else
							Sapi.speak plaintextSearch
							If InStr(1, commandSynonym, "synonym") > 0 Then
								synonymResult = msgbox("Synonyms: " + PCase(plaintextSearch),,"Synonyms")
							ElseIf InStr(1, commandSynonym, "antonym") > 0 Then
								synonymResult = msgbox("Antonyms: " + PCase(plaintextSearch),,"Antonyms")
							End If
						End If
					End If
					foundplaintext = 1
				ElseIf InStr(1, plaintextSearch, "didyoumean") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<tip ") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<error>") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<plaintext/>") > 0 And foundError=0 Then
					Sapi.speak "Unfortunately, a list of synonyms for your requested word isn't available."
					foundError = 1
					foundplaintext = 1
				End If
			loop
			objFileXML.Close
			If foundplaintext=0 Then
				Sapi.speak "Unfortunately, a list of synonyms for your requested word isn't available."
			End If
		Else
			do while not objFileXML.AtEndOfStream
				plaintextSearch = objFileXML.ReadLine()
				If InStr(1, plaintextSearch, "<img src=") > 0 Then
					plaintextSearch = Replace(plaintextSearch, "<img src='", "")
					quotationPosition = InStr(1, plaintextSearch, "'")
					imageURL = CStr(Left(plaintextSearch, quotationPosition - 1))
					imageURL = Replace(imageURL, "amp;", "")
					Sapi.speak "Okay, here's a table of all the airplanes currently above us"
					Set objExplorer = CreateObject("InternetExplorer.Application")
					With objExplorer
						.Visible = 1
						.Toolbar=False
						.Statusbar=False
						.Top=800
						.Left=800
						.Height=400
						.Width=360
						.Navigate imageURL
					End With
					foundplaintext = 1
				ElseIf InStr(1, plaintextSearch, "didyoumean") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<tip ") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<error>") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<plaintext/>") > 0 And foundError=0 Then
					Sapi.speak Replace(rand(Array("Okay, here's what I found for, #%command%#, on google", "Okay, here's what I found on the web for, #%command%#.", "Okay, I found this on the web for, #%command%#", "Okay, check it out. Here's what I found for, #%command%#, on google", "Okay, check it out. Here's what I found for, #%command%#, on the web", "I found this on google for, #%command%#", "Here's what I found for, #%command%#, on the web")), "#%command%#", commandSynonym)
					googleSearchQueryRefined=Trim(commandSynonym)
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "+", "%2B")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "#", "%23")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "%", "%25")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "&", "%26")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, " ", "%20")
					startingurl="https://www.google.ca/?gws_rd=cr&ei=acGMV629B8TQ-QH6vZroDQ&safe=active&ssui=on#safe=active&q="
					fullurl=startingurl+googleSearchQueryRefined
					a.run fullurl
					foundError = 1
					foundplaintext = 1
				End If
			loop
			objFileXML.Close
			If foundplaintext=0 Then
				Sapi.speak Replace(rand(Array("Okay, here's what I found for, #%command%#, on google", "Okay, here's what I found on the web for, #%command%#.", "Okay, I found this on the web for, #%command%#", "Okay, check it out. Here's what I found for, #%command%#, on google", "Okay, check it out. Here's what I found for, #%command%#, on the web", "I found this on google for, #%command%#", "Here's what I found for, #%command%#, on the web")), "#%command%#", commandSynonym)
				googleSearchQueryRefined=Trim(commandSynonym)
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "+", "%2B")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "#", "%23")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "%", "%25")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "&", "%26")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, " ", "%20")
				startingurl="https://www.google.ca/?gws_rd=cr&ei=acGMV629B8TQ-QH6vZroDQ&safe=active&ssui=on#safe=active&q="
				fullurl=startingurl+googleSearchQueryRefined
				a.run fullurl
			End If
		End If

	  Else
		 Sapi.speak "Unfortunately, I'm unable to answer your question at the moment. Please make sure that your internet connection is secure and that airplane mode is not turned on"
	End If
End Function

Function doWolframDefineSearch(commandDefine)
	foundplaneWolfram(command)
	Dim WinHttpReq
	Set WinHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
	Set objFSO=CreateObject("Scripting.FileSystemObject")
	Dim sTemp
	foundError=0
	foundplaintext=0
	startingURL = "http://api.wolframalpha.com/v2/query?input="
	If isplaneOverhead = 1 Then
		endingURL = "&appid=AAXTLK-5XE3J9E8L3&location=Ottawa,%20Ontario&podindex=2&format=image"
	Else
		endingURL = "&appid=AAXTLK-5XE3J9E8L3&location=Ottawa,%20Ontario&podindex=2&format=plaintext"
	End If
	commandRefinedXML = commandDefine
	commandRefinedXML = Replace(commandRefinedXML, "+", " plus ")
	commandRefinedXML = Replace(commandRefinedXML, "#", "%23")
	commandRefinedXML = Replace(commandRefinedXML, "%", "%25")
	commandRefinedXML = Replace(commandRefinedXML, "^", "%5E")
	commandRefinedXML = Replace(commandRefinedXML, "&", "%26")
	commandRefinedXML = Replace(commandRefinedXML, "=", "%3D")
	commandRefinedXML = Replace(commandRefinedXML, " ", "+")
	commandRefinedXML = Replace(commandRefinedXML, "\", "%5C")
	commandRefinedXML = Replace(commandRefinedXML, "|", "%7C")
	commandRefinedXML = Replace(commandRefinedXML, "{", "%7B")
	commandRefinedXML = Replace(commandRefinedXML, "{", "%7D")
	commandRefinedXML = Replace(commandRefinedXML, "[", "%5B")
	commandRefinedXML = Replace(commandRefinedXML, "]", "%5D")
	commandRefinedXML = Replace(commandRefinedXML, "'", "%27")
	commandRefinedXML = Replace(commandRefinedXML, "?", "%3F")
	commandRefinedXML = Replace(commandRefinedXML, "/", "%2F")
	fullXMLUrl = startingURL + commandRefinedXML + endingURL
	WinHttpReq.Open "GET", fullXMLURL, False

	WinHttpReq.Send
	 If (WinHttpReq.Status = 200) Then

		  correctString=""
		  sTemp = WinHttpReq.ResponseText

		  sTemp = Replace(sTemp, vbCr, vbcrlf)
		  sTemp = Replace(sTemp, vbLf, vbcrlf)
		 
		   WebText = sTemp

		   Set objFile = objFSO.OpenTextFile("C:\Sherby Interface\xmlWolframFile.txt", 2, True)
		   on error resume next
		   objFile.WriteLine sTemp
If Err Then
    WScript.StdErr.WriteLine "error "
End If 
		   objFile.Close

			Set objFileXML = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\xmlWolframFile.txt",1)

		If isplaneOverhead = 0 Then
			do while not objFileXML.AtEndOfStream
				plaintextSearch = objFileXML.ReadLine()
				If InStr(1, plaintextSearch, "<plaintext>") > 0 Then
					plaintextSearch = Replace(plaintextSearch, "<plaintext>", "")
					plaintextSearch = Replace(plaintextSearch, "</plaintext>", "")
					If Trim(plaintextSearch) = "" OR plaintextSearch = " " OR plaintextSearch = "  " OR plaintextSearch = "   " OR plaintextSearch = "(data not available)" Then
						Sapi.speak "Unfortunately, the definition for your requested word isn't available"
						foundError = 1
					Else
						If InStr(1, plaintextSearch, "|") > 0 Then
							plaintextSearchSplit = Split(plaintextSearch, "|")
							plaintextArrayLength = UBound(plaintextSearchSplit)
							If plaintextArrayLength > 2 Then
								plaintextSearch = Left(plaintextSearchSplit(2), Len(plaintextSearchSplit(2)) - 2)
								Sapi.speak plaintextSearch
								definitionResult = msgbox("Definition: " + PCase(plaintextSearch),,"Definition")
							Else
								plaintextSearch = CStr(plaintextSearchSplit(2))
								Sapi.speak plaintextSearch
								definitionResult = msgbox("Definition: " + PCase(plaintextSearch),,"Definition")
							End If
						Else
							Sapi.speak plaintextSearch
							definitionResult = msgbox("Definition: " + PCase(plaintextSearch),,"Definition")
						End If
					End If
					foundplaintext = 1
				ElseIf InStr(1, plaintextSearch, "didyoumean") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<tip ") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<error>") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<plaintext/>") > 0 And foundError=0 Then
					Sapi.speak "Unfortunately, the definition for your requested word isn't available"
					foundError = 1
					foundplaintext = 1
				End If
			loop
			objFileXML.Close
			If foundplaintext=0 Then
				Sapi.speak "Unfortunately, the definition for your requested word isn't available"
			End If
		Else
			do while not objFileXML.AtEndOfStream
				plaintextSearch = objFileXML.ReadLine()
				If InStr(1, plaintextSearch, "<img src=") > 0 Then
					plaintextSearch = Replace(plaintextSearch, "<img src='", "")
					quotationPosition = InStr(1, plaintextSearch, "'")
					imageURL = CStr(Left(plaintextSearch, quotationPosition - 1))
					imageURL = Replace(imageURL, "amp;", "")
					Sapi.speak "Okay, here's a table of all the airplanes currently above us"
					Set objExplorer = CreateObject("InternetExplorer.Application")
					With objExplorer
						.Visible = 1
						.Toolbar=False
						.Statusbar=False
						.Top=800
						.Left=800
						.Height=400
						.Width=360
						.Navigate imageURL
					End With
					foundplaintext = 1
				ElseIf InStr(1, plaintextSearch, "didyoumean") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<tip ") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<error>") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<plaintext/>") > 0 And foundError=0 Then
					Sapi.speak Replace(rand(Array("Okay, here's what I found for, #%command%#, on google", "Okay, here's what I found on the web for, #%command%#.", "Okay, I found this on the web for, #%command%#", "Okay, check it out. Here's what I found for, #%command%#, on google", "Okay, check it out. Here's what I found for, #%command%#, on the web", "I found this on google for, #%command%#", "Here's what I found for, #%command%#, on the web")), "#%command%#", commandDefine)
					googleSearchQueryRefined=Trim(commandDefine)
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "+", "%2B")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "#", "%23")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "%", "%25")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "&", "%26")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, " ", "%20")
					startingurl="https://www.google.ca/?gws_rd=cr&ei=acGMV629B8TQ-QH6vZroDQ&safe=active&ssui=on#safe=active&q="
					fullurl=startingurl+googleSearchQueryRefined
					a.run fullurl
					foundError = 1
					foundplaintext = 1
				End If
			loop
			objFileXML.Close
			If foundplaintext=0 Then
				Sapi.speak Replace(rand(Array("Okay, here's what I found for, #%command%#, on google", "Okay, here's what I found on the web for, #%command%#.", "Okay, I found this on the web for, #%command%#", "Okay, check it out. Here's what I found for, #%command%#, on google", "Okay, check it out. Here's what I found for, #%command%#, on the web", "I found this on google for, #%command%#", "Here's what I found for, #%command%#, on the web")), "#%command%#", commandDefine)
				googleSearchQueryRefined=Trim(commandDefine)
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "+", "%2B")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "#", "%23")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "%", "%25")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "&", "%26")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, " ", "%20")
				startingurl="https://www.google.ca/?gws_rd=cr&ei=acGMV629B8TQ-QH6vZroDQ&safe=active&ssui=on#safe=active&q="
				fullurl=startingurl+googleSearchQueryRefined
				a.run fullurl
			End If
		End If

	  Else
		 Sapi.speak "Unfortunately, I'm unable to answer your question at the moment. Please make sure that your internet connection is secure and that airplane mode is not turned on"
	End If
End Function

Function doWolframTipSearch(commandTip)
	foundplaneWolfram(command)
	Dim WinHttpReq
	Set WinHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
	Set objFSO=CreateObject("Scripting.FileSystemObject")
	Dim sTemp
	foundError=0
	foundplaintext=0
	startingURL = "http://api.wolframalpha.com/v2/query?input="
	If isplaneOverhead = 1 Then
		endingURL = "&appid=AAXTLK-5XE3J9E8L3&location=Ottawa,%20Ontario&podindex=2&format=image"
	Else
		endingURL = "&appid=AAXTLK-5XE3J9E8L3&location=Ottawa,%20Ontario&podindex=2&format=plaintext"
	End If
	commandRefinedXML = commandTip
	commandRefinedXML = Replace(commandRefinedXML, "+", " plus ")
	commandRefinedXML = Replace(commandRefinedXML, "#", "%23")
	commandRefinedXML = Replace(commandRefinedXML, "%", "%25")
	commandRefinedXML = Replace(commandRefinedXML, "^", "%5E")
	commandRefinedXML = Replace(commandRefinedXML, "&", "%26")
	commandRefinedXML = Replace(commandRefinedXML, "=", "%3D")
	commandRefinedXML = Replace(commandRefinedXML, " ", "+")
	commandRefinedXML = Replace(commandRefinedXML, "\", "%5C")
	commandRefinedXML = Replace(commandRefinedXML, "|", "%7C")
	commandRefinedXML = Replace(commandRefinedXML, "{", "%7B")
	commandRefinedXML = Replace(commandRefinedXML, "{", "%7D")
	commandRefinedXML = Replace(commandRefinedXML, "[", "%5B")
	commandRefinedXML = Replace(commandRefinedXML, "]", "%5D")
	commandRefinedXML = Replace(commandRefinedXML, "'", "%27")
	commandRefinedXML = Replace(commandRefinedXML, "?", "%3F")
	commandRefinedXML = Replace(commandRefinedXML, "/", "%2F")
	fullXMLUrl = startingURL + commandRefinedXML + endingURL
	WinHttpReq.Open "GET", fullXMLURL, False

	WinHttpReq.Send
	 If (WinHttpReq.Status = 200) Then
		  correctString=""
		  sTemp = WinHttpReq.ResponseText

		  sTemp = Replace(sTemp, vbCr, vbcrlf)
		  sTemp = Replace(sTemp, vbLf, vbcrlf)
		   WebText = sTemp

		   Set objFile = objFSO.OpenTextFile("C:\Sherby Interface\xmlWolframFile.txt", 2, True)
		    on error resume next
		   objFile.WriteLine sTemp
		If Err Then
			WScript.StdErr.WriteLine "error "
		End If
		   objFile.Close

			Set objFileXML = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\xmlWolframFile.txt",1)

		If isplaneOverhead = 0 Then
			do while not objFileXML.AtEndOfStream
				plaintextSearch = objFileXML.ReadLine()
				If InStr(1, plaintextSearch, "<plaintext>") > 0 Then
					plaintextSearch = Replace(plaintextSearch, "<plaintext>", "")
					plaintextSearch = Replace(plaintextSearch, "</plaintext>", "")
					If Trim(plaintextSearch) = "" OR plaintextSearch = " " OR plaintextSearch = "  " OR plaintextSearch = "   " Then
						Sapi.speak Replace(rand(Array("Okay, here's what I found for, #%command%#, on google", "Okay, here's what I found on the web for, #%command%#.", "Okay, I found this on the web for, #%command%#", "Okay, check it out. Here's what I found for, #%command%#, on google", "Okay, check it out. Here's what I found for, #%command%#, on the web", "I found this on google for, #%command%#", "Here's what I found for, #%command%#, on the web")), "#%command%#", commandTip)
						googleSearchQueryRefined=Trim(commandTip)
						googleSearchQueryRefined=Replace(googleSearchQueryRefined, "+", "%2B")
						googleSearchQueryRefined=Replace(googleSearchQueryRefined, "#", "%23")
						googleSearchQueryRefined=Replace(googleSearchQueryRefined, "%", "%25")
						googleSearchQueryRefined=Replace(googleSearchQueryRefined, "&", "%26")
						googleSearchQueryRefined=Replace(googleSearchQueryRefined, " ", "%20")
						startingurl="https://www.google.ca/?gws_rd=cr&ei=acGMV629B8TQ-QH6vZroDQ&safe=active&ssui=on#safe=active&q="
						fullurl=startingurl+googleSearchQueryRefined
						a.run fullurl
						foundError = 1
					Else
						filteringString = InStr(1, plaintextSearch, "amount of tip | C")
						filteringString = Mid(plaintextSearch, filteringString + 17)
						plaintextSearch = Left(filteringString, 6)
						Sapi.speak "The tip amount is " + plaintextSearch
						tipAmount = msgbox("Tip amount: " + CStr(plaintextSearch),,"Tip")
					End If
					foundplaintext = 1
				ElseIf InStr(1, plaintextSearch, "didyoumean") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<tip ") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<error>") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<plaintext/>") > 0 And foundError=0 OR InStr(1, plaintextSearch, "not available") > 0 And foundError=0 Then
					Sapi.speak Replace(rand(Array("Okay, here's what I found for, #%command%#, on google", "Okay, here's what I found on the web for, #%command%#.", "Okay, I found this on the web for, #%command%#", "Okay, check it out. Here's what I found for, #%command%#, on google", "Okay, check it out. Here's what I found for, #%command%#, on the web", "I found this on google for, #%command%#", "Here's what I found for, #%command%#, on the web")), "#%command%#", commandTip)
					googleSearchQueryRefined=Trim(commandTip)
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "+", "%2B")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "#", "%23")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "%", "%25")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "&", "%26")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, " ", "%20")
					startingurl="https://www.google.ca/?gws_rd=cr&ei=acGMV629B8TQ-QH6vZroDQ&safe=active&ssui=on#safe=active&q="
					fullurl=startingurl+googleSearchQueryRefined
					a.run fullurl
					foundError = 1
					foundplaintext = 1
				End If
			loop
			objFileXML.Close
			If foundplaintext=0 Then
				Sapi.speak Replace(rand(Array("Okay, here's what I found for, #%command%#, on google", "Okay, here's what I found on the web for, #%command%#.", "Okay, I found this on the web for, #%command%#", "Okay, check it out. Here's what I found for, #%command%#, on google", "Okay, check it out. Here's what I found for, #%command%#, on the web", "I found this on google for, #%command%#", "Here's what I found for, #%command%#, on the web")), "#%command%#", commandTip)
				googleSearchQueryRefined=Trim(commandTip)
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "+", "%2B")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "#", "%23")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "%", "%25")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "&", "%26")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, " ", "%20")
				startingurl="https://www.google.ca/?gws_rd=cr&ei=acGMV629B8TQ-QH6vZroDQ&safe=active&ssui=on#safe=active&q="
				fullurl=startingurl+googleSearchQueryRefined
				a.run fullurl
			End If
		Else
			do while not objFileXML.AtEndOfStream
				plaintextSearch = objFileXML.ReadLine()
				If InStr(1, plaintextSearch, "<img src=") > 0 Then
					plaintextSearch = Replace(plaintextSearch, "<img src='", "")
					quotationPosition = InStr(1, plaintextSearch, "'")
					imageURL = CStr(Left(plaintextSearch, quotationPosition - 1))
					imageURL = Replace(imageURL, "amp;", "")
					Sapi.speak "Okay, here's a table of all the airplanes currently above us"
					Set objExplorer = CreateObject("InternetExplorer.Application")
					With objExplorer
						.Visible = 1
						.Toolbar=False
						.Statusbar=False
						.Top=800
						.Left=800
						.Height=400
						.Width=360
						.Navigate imageURL
					End With
					foundplaintext = 1
				ElseIf InStr(1, plaintextSearch, "didyoumean") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<tip ") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<error>") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<plaintext/>") > 0 And foundError=0 OR InStr(1, plaintextSearch, "not available") > 0 And foundError=0 Then
					Sapi.speak Replace(rand(Array("Okay, here's what I found for, #%command%#, on google", "Okay, here's what I found on the web for, #%command%#.", "Okay, I found this on the web for, #%command%#", "Okay, check it out. Here's what I found for, #%command%#, on google", "Okay, check it out. Here's what I found for, #%command%#, on the web", "I found this on google for, #%command%#", "Here's what I found for, #%command%#, on the web")), "#%command%#", commandTip)
					googleSearchQueryRefined=Trim(commandTip)
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "+", "%2B")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "#", "%23")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "%", "%25")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "&", "%26")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, " ", "%20")
					startingurl="https://www.google.ca/?gws_rd=cr&ei=acGMV629B8TQ-QH6vZroDQ&safe=active&ssui=on#safe=active&q="
					fullurl=startingurl+googleSearchQueryRefined
					a.run fullurl
					foundError = 1
					foundplaintext = 1
				End If
			loop
			objFileXML.Close
			If foundplaintext=0 Then
				Sapi.speak Replace(rand(Array("Okay, here's what I found for, #%command%#, on google", "Okay, here's what I found on the web for, #%command%#.", "Okay, I found this on the web for, #%command%#", "Okay, check it out. Here's what I found for, #%command%#, on google", "Okay, check it out. Here's what I found for, #%command%#, on the web", "I found this on google for, #%command%#", "Here's what I found for, #%command%#, on the web")), "#%command%#", commandTip)
				googleSearchQueryRefined=Trim(commandTip)
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "+", "%2B")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "#", "%23")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "%", "%25")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "&", "%26")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, " ", "%20")
				startingurl="https://www.google.ca/?gws_rd=cr&ei=acGMV629B8TQ-QH6vZroDQ&safe=active&ssui=on#safe=active&q="
				fullurl=startingurl+googleSearchQueryRefined
				a.run fullurl
			End If
		End If

	  Else
		 Sapi.speak "Unfortunately, I'm unable to answer your question at the moment. Please make sure that your internet connection is secure and that airplane mode is not turned on"
	End If
End Function

Function doWolframSunSearch(commandSun)
	isplaneOverhead = 0
	Dim WinHttpReq
	Set WinHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
	Set objFSO=CreateObject("Scripting.FileSystemObject")
	Dim sTemp
	foundError=0
	foundplaintext=0
	startingURL = "http://api.wolframalpha.com/v2/query?input="
	endingURL = "&appid=AAXTLK-5XE3J9E8L3&location=Ottawa,%20Ontario&podindex=2&format=plaintext"
	commandRefinedXML = commandSun
	commandRefinedXML = Replace(commandRefinedXML, "+", " plus ")
	commandRefinedXML = Replace(commandRefinedXML, "#", "%23")
	commandRefinedXML = Replace(commandRefinedXML, "%", "%25")
	commandRefinedXML = Replace(commandRefinedXML, "^", "%5E")
	commandRefinedXML = Replace(commandRefinedXML, "&", "%26")
	commandRefinedXML = Replace(commandRefinedXML, "=", "%3D")
	commandRefinedXML = Replace(commandRefinedXML, " ", "+")
	commandRefinedXML = Replace(commandRefinedXML, "\", "%5C")
	commandRefinedXML = Replace(commandRefinedXML, "|", "%7C")
	commandRefinedXML = Replace(commandRefinedXML, "{", "%7B")
	commandRefinedXML = Replace(commandRefinedXML, "{", "%7D")
	commandRefinedXML = Replace(commandRefinedXML, "[", "%5B")
	commandRefinedXML = Replace(commandRefinedXML, "]", "%5D")
	commandRefinedXML = Replace(commandRefinedXML, "'", "%27")
	commandRefinedXML = Replace(commandRefinedXML, "?", "%3F")
	commandRefinedXML = Replace(commandRefinedXML, "/", "%2F")
	fullXMLUrl = startingURL + commandRefinedXML + endingURL
	WinHttpReq.Open "GET", fullXMLURL, False

	WinHttpReq.Send
	If (WinHttpReq.Status = 200) Then
		correctString=""

		sTemp = WinHttpReq.ResponseText
		sTemp = Replace(sTemp, vbCr, vbcrlf)
		sTemp = Replace(sTemp, vbLf, vbcrlf)
		WebText = sTemp

		splitCommand = Split(commandSun, " ")


		If InStr(1, commandSun, " in ") > 0 Then
			indexOfCityName = inArray(splitCommand, "in") + 1
			city = splitCommand(indexOfCityName)
			If InStr(1, commandSun, "rise") > 0 OR InStr(1, commandSun, "up") > 0 Then
				Sapi.speak "In " + city + ", the sun rises at, "
			Else
				Sapi.speak "In " + city + ", the sun sets at, "
			End If
		ElseIf InStr(1, commandSun, " for ") > 0 Then
			indexOfCityName = inArray(splitCommand, "for") + 1
			city = splitCommand(indexOfCityName)
			If InStr(1, commandSun, "rise") > 0 OR InStr(1, commandSun, "up") > 0 Then
				Sapi.speak "In " + city + ", the sun rises at, "
			Else
				Sapi.speak "In " + city + ", the sun sets at, "
			End If
		End If
		   Set objFile = objFSO.OpenTextFile("C:\Sherby Interface\xmlWolframFile.txt", 2, True)
		   on error resume next
		   objFile.WriteLine sTemp
			If Err Then
				WScript.StdErr.WriteLine "error "
			End If 
		   objFile.Close

			Set objFileXML = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\xmlWolframFile.txt",1)
			do while not objFileXML.AtEndOfStream
				plaintextSearch = objFileXML.ReadLine()
				If InStr(1, plaintextSearch, "<plaintext>") > 0 Then
					plaintextSearch = Replace(plaintextSearch, "<plaintext>", "")
					plaintextSearch = Replace(plaintextSearch, "</plaintext>", "")
					If Trim(plaintextSearch) = "" OR plaintextSearch = " " OR plaintextSearch = "  " OR plaintextSearch = "   " Then
						Sapi.speak "Unfortunately, the sunset and sunrise information for " + city + " is currently unavailable"
					Else
						If InStr(1, plaintextSearch, "pm") > 0 Then
							endingTimeIndex = InStr(1, plaintextSearch, "pm")
						ElseIf InStr(1, plaintextSearch, "am") > 0 Then
							endingTimeIndex = InStr(1, plaintextSearch, "am")
						Else
							Sapi.speak "Unfortunately, the sunset and sunrise information for " + city + " is currently unavailable"
						End If
						plaintextSearchTimeFull = Left(plaintextSearch, endingTimeIndex + 2)
						Sapi.speak plaintextSearchTimeFull
					End If
					foundplaintext = 1
				ElseIf InStr(1, plaintextSearch, "didyoumean") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<tip ") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<error>") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<plaintext/>") > 0 And foundError=0 OR InStr(1, plaintextSearch, "not available") > 0 And foundError=0 Then
					Sapi.speak "Unfortunately, the sunset and sunrise information for " + city + " is currently unavailable"
				End If
			loop
			objFileXML.Close
			If foundplaintext=0 Then
				Sapi.speak "Unfortunately, the sunset and sunrise information for " + city + " is currently unavailable"
			End If
	Else
		 Sapi.speak "Unfortunately, I'm unable to answer your question at the moment. Please make sure that your internet connection is secure and that airplane mode is not turned on"
	End If
End Function

Function doWolframDateSearch(commandDate)
	isplaneOverhead = 0
	Dim WinHttpReq
	Set WinHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
	Set objFSO=CreateObject("Scripting.FileSystemObject")
	Dim sTemp
	foundError=0
	foundplaintext=0
	startingURL = "http://api.wolframalpha.com/v2/query?input="
	endingURL = "&appid=AAXTLK-5XE3J9E8L3&location=Ottawa,%20Ontario&podindex=2&format=plaintext"
	commandRefinedXML = commandDate
	commandRefinedXML = Replace(commandRefinedXML, "+", " plus ")
	commandRefinedXML = Replace(commandRefinedXML, "#", "%23")
	commandRefinedXML = Replace(commandRefinedXML, "%", "%25")
	commandRefinedXML = Replace(commandRefinedXML, "^", "%5E")
	commandRefinedXML = Replace(commandRefinedXML, "&", "%26")
	commandRefinedXML = Replace(commandRefinedXML, "=", "%3D")
	commandRefinedXML = Replace(commandRefinedXML, " ", "+")
	commandRefinedXML = Replace(commandRefinedXML, "\", "%5C")
	commandRefinedXML = Replace(commandRefinedXML, "|", "%7C")
	commandRefinedXML = Replace(commandRefinedXML, "{", "%7B")
	commandRefinedXML = Replace(commandRefinedXML, "{", "%7D")
	commandRefinedXML = Replace(commandRefinedXML, "[", "%5B")
	commandRefinedXML = Replace(commandRefinedXML, "]", "%5D")
	commandRefinedXML = Replace(commandRefinedXML, "'", "%27")
	commandRefinedXML = Replace(commandRefinedXML, "?", "%3F")
	commandRefinedXML = Replace(commandRefinedXML, "/", "%2F")
	fullXMLUrl = startingURL + commandRefinedXML + endingURL
	WinHttpReq.Open "GET", fullXMLURL, False

	WinHttpReq.Send
	If (WinHttpReq.Status = 200) Then

		  correctString=""
		  sTemp = WinHttpReq.ResponseText

		  sTemp = Replace(sTemp, vbCr, vbcrlf)
		  sTemp = Replace(sTemp, vbLf, vbcrlf)
		 
		   WebText = sTemp
			splitCommand = Split(commandDate, " ")
			indexOfCityName = inArray(splitCommand, "in") + 1
			city = splitCommand(indexOfCityName)
			Sapi.speak "The current day in " + city + " is..."
		   Set objFile = objFSO.OpenTextFile("C:\Sherby Interface\xmlWolframFile.txt", 2, True)
		   on error resume next
		   objFile.WriteLine sTemp
			If Err Then
				WScript.StdErr.WriteLine "error "
			End If 
		   objFile.Close

			Set objFileXML = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\xmlWolframFile.txt",1)
			do while not objFileXML.AtEndOfStream
				plaintextSearch = objFileXML.ReadLine()
				If InStr(1, plaintextSearch, "<plaintext>") > 0 Then
					plaintextSearch = Replace(plaintextSearch, "<plaintext>", "")
					plaintextSearch = Replace(plaintextSearch, "</plaintext>", "")
					If Trim(plaintextSearch) = "" OR plaintextSearch = " " OR plaintextSearch = "  " OR plaintextSearch = "   " Then
						Sapi.speak "Unfortunately, the date in " + city + " is currently unavailable"
					Else
						Sapi.speak plaintextSearch
					End If
					foundplaintext = 1
				ElseIf InStr(1, plaintextSearch, "didyoumean") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<tip ") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<error>") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<plaintext/>") > 0 And foundError=0 OR InStr(1, plaintextSearch, "not available") > 0 And foundError=0 Then
					Sapi.speak "Unfortunately, the date in " + city + " is currently unavailable"
				End If
			loop
			objFileXML.Close
			If foundplaintext=0 Then
				Sapi.speak "Unfortunately, the date in " + city + " is currently unavailable"
			End If
	Else
		 Sapi.speak "Unfortunately, I'm unable to answer your question at the moment. Please make sure that your internet connection is secure and that airplane mode is not turned on"
	End If
End Function

Function doWolframTimeSearch(commandTime)
	isplaneOverhead = 0
	Dim WinHttpReq
	Set WinHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
	Set objFSO=CreateObject("Scripting.FileSystemObject")
	Dim sTemp
	foundError=0
	foundplaintext=0
	startingURL = "http://api.wolframalpha.com/v2/query?input="
	endingURL = "&appid=AAXTLK-5XE3J9E8L3&location=Ottawa,%20Ontario&podindex=2&format=plaintext"
	commandRefinedXML = commandTime
	commandRefinedXML = Replace(commandRefinedXML, "+", " plus ")
	commandRefinedXML = Replace(commandRefinedXML, "#", "%23")
	commandRefinedXML = Replace(commandRefinedXML, "%", "%25")
	commandRefinedXML = Replace(commandRefinedXML, "^", "%5E")
	commandRefinedXML = Replace(commandRefinedXML, "&", "%26")
	commandRefinedXML = Replace(commandRefinedXML, "=", "%3D")
	commandRefinedXML = Replace(commandRefinedXML, " ", "+")
	commandRefinedXML = Replace(commandRefinedXML, "\", "%5C")
	commandRefinedXML = Replace(commandRefinedXML, "|", "%7C")
	commandRefinedXML = Replace(commandRefinedXML, "{", "%7B")
	commandRefinedXML = Replace(commandRefinedXML, "{", "%7D")
	commandRefinedXML = Replace(commandRefinedXML, "[", "%5B")
	commandRefinedXML = Replace(commandRefinedXML, "]", "%5D")
	commandRefinedXML = Replace(commandRefinedXML, "'", "%27")
	commandRefinedXML = Replace(commandRefinedXML, "?", "%3F")
	commandRefinedXML = Replace(commandRefinedXML, "/", "%2F")
	fullXMLUrl = startingURL + commandRefinedXML + endingURL
	WinHttpReq.Open "GET", fullXMLURL, False

	WinHttpReq.Send
	If (WinHttpReq.Status = 200) Then

		  correctString=""
		  sTemp = WinHttpReq.ResponseText

		  sTemp = Replace(sTemp, vbCr, vbcrlf)
		  sTemp = Replace(sTemp, vbLf, vbcrlf)
		 
		   WebText = sTemp

		   Set objFile = objFSO.OpenTextFile("C:\Sherby Interface\xmlWolframFile.txt", 2, True)
		   on error resume next
		   objFile.WriteLine sTemp
			If Err Then
				WScript.StdErr.WriteLine "error "
			End If 
		   objFile.Close

			Set objFileXML = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\xmlWolframFile.txt",1)
			splitCommand = Split(commandTime, " ")
			indexOfCityName = inArray(splitCommand, "in") + 1
			city = splitCommand(indexOfCityName)
			Sapi.speak "The current time in " + city + "is"
			do while not objFileXML.AtEndOfStream
				plaintextSearch = objFileXML.ReadLine()
				If InStr(1, plaintextSearch, "<plaintext>") > 0 Then
					plaintextSearch = Replace(plaintextSearch, "<plaintext>", "")
					plaintextSearch = Replace(plaintextSearch, "</plaintext>", "")
					If Trim(plaintextSearch) = "" OR plaintextSearch = " " OR plaintextSearch = "  " OR plaintextSearch = "   " Then
						Sapi.speak "Unfortunately, the time in " + city + " is currently unavailable"
					Else
						If InStr(1, plaintextSearch, "pm") > 0 Then
							endingTimeIndex = InStr(1, plaintextSearch, "pm")
						ElseIf InStr(1, plaintextSearch, "am") > 0 Then
							endingTimeIndex = InStr(1, plaintextSearch, "am")
						Else
							Sapi.speak "Unfortunately, the time in " + city + " is currently unavailable"
						End If
						plaintextSearch = Left(plaintextSearch, endingTimeIndex + 1)
						endingTimeIndex = Right(plaintextSearch, 3)
						timeWithoutIndex = Left(plaintextSearch, Len(plaintextSearch) - 6)
						plaintextSearch = timeWithoutIndex + endingTimeIndex
						Sapi.speak plaintextSearch
					End If
					foundplaintext = 1
				ElseIf InStr(1, plaintextSearch, "didyoumean") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<tip ") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<error>") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<plaintext/>") > 0 And foundError=0 OR InStr(1, plaintextSearch, "not available") > 0 And foundError=0 Then
					Sapi.speak "Unfortunately, the time in " + city + " is currently unavailable"
				End If
			loop
			objFileXML.Close
			If foundplaintext=0 Then
				Sapi.speak "Unfortunately, the time in " + city + " is currently unavailable"
			End If
	Else
		 Sapi.speak "Unfortunately, I'm unable to answer your question at the moment. Please make sure that your internet connection is secure and that airplane mode is not turned on"
	End If
End Function


Function doWolframDateSearchFor(commandDateFor)
	isplaneOverhead = 0
	Dim WinHttpReq
	Set WinHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
	Set objFSO=CreateObject("Scripting.FileSystemObject")
	Dim sTemp
	foundError=0
	foundplaintext=0
	startingURL = "http://api.wolframalpha.com/v2/query?input="
	endingURL = "&appid=AAXTLK-5XE3J9E8L3&location=Ottawa,%20Ontario&podindex=2&format=plaintext"
	commandRefinedXML = commandDateFor
	commandRefinedXML = Replace(commandRefinedXML, "+", " plus ")
	commandRefinedXML = Replace(commandRefinedXML, "#", "%23")
	commandRefinedXML = Replace(commandRefinedXML, "%", "%25")
	commandRefinedXML = Replace(commandRefinedXML, "^", "%5E")
	commandRefinedXML = Replace(commandRefinedXML, "&", "%26")
	commandRefinedXML = Replace(commandRefinedXML, "=", "%3D")
	commandRefinedXML = Replace(commandRefinedXML, " ", "+")
	commandRefinedXML = Replace(commandRefinedXML, "\", "%5C")
	commandRefinedXML = Replace(commandRefinedXML, "|", "%7C")
	commandRefinedXML = Replace(commandRefinedXML, "{", "%7B")
	commandRefinedXML = Replace(commandRefinedXML, "{", "%7D")
	commandRefinedXML = Replace(commandRefinedXML, "[", "%5B")
	commandRefinedXML = Replace(commandRefinedXML, "]", "%5D")
	commandRefinedXML = Replace(commandRefinedXML, "'", "%27")
	commandRefinedXML = Replace(commandRefinedXML, "?", "%3F")
	commandRefinedXML = Replace(commandRefinedXML, "/", "%2F")
	fullXMLUrl = startingURL + commandRefinedXML + endingURL
	WinHttpReq.Open "GET", fullXMLURL, False

	WinHttpReq.Send
	If (WinHttpReq.Status = 200) Then

		  correctString=""
		  sTemp = WinHttpReq.ResponseText

		  sTemp = Replace(sTemp, vbCr, vbcrlf)
		  sTemp = Replace(sTemp, vbLf, vbcrlf)
		 
		   WebText = sTemp
			splitCommand = Split(commandDateFor, " ")
			indexOfCityName = inArray(splitCommand, "for") + 1
			city = splitCommand(indexOfCityName)
			Sapi.speak "The current day in " + city + " is..."
		   Set objFile = objFSO.OpenTextFile("C:\Sherby Interface\xmlWolframFile.txt", 2, True)
		   on error resume next
		   objFile.WriteLine sTemp
			If Err Then
				WScript.StdErr.WriteLine "error "
			End If 
		   objFile.Close

			Set objFileXML = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\xmlWolframFile.txt",1)
			do while not objFileXML.AtEndOfStream
				plaintextSearch = objFileXML.ReadLine()
				If InStr(1, plaintextSearch, "<plaintext>") > 0 Then
					plaintextSearch = Replace(plaintextSearch, "<plaintext>", "")
					plaintextSearch = Replace(plaintextSearch, "</plaintext>", "")
					If Trim(plaintextSearch) = "" OR plaintextSearch = " " OR plaintextSearch = "  " OR plaintextSearch = "   " Then
						Sapi.speak "Unfortunately, the date in " + city + " is currently unavailable"
					Else
						Sapi.speak plaintextSearch
					End If
					foundplaintext = 1
				ElseIf InStr(1, plaintextSearch, "didyoumean") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<tip ") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<error>") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<plaintext/>") > 0 And foundError=0 OR InStr(1, plaintextSearch, "not available") > 0 And foundError=0 Then
					Sapi.speak "Unfortunately, the date in " + city + " is currently unavailable"
				End If
			loop
			objFileXML.Close
			If foundplaintext=0 Then
				Sapi.speak "Unfortunately, the date in " + city + " is currently unavailable"
			End If
	Else
		 Sapi.speak "Unfortunately, I'm unable to answer your question at the moment. Please make sure that your internet connection is secure and that airplane mode is not turned on"
	End If
End Function

Function doWolframTimeSearchFor(commandTimeFor)
	isplaneOverhead = 0
	Dim WinHttpReq
	Set WinHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
	Set objFSO=CreateObject("Scripting.FileSystemObject")
	Dim sTemp
	foundError=0
	foundplaintext=0
	startingURL = "http://api.wolframalpha.com/v2/query?input="
	endingURL = "&appid=AAXTLK-5XE3J9E8L3&location=Ottawa,%20Ontario&podindex=2&format=plaintext"
	commandRefinedXML = commandTimeFor
	commandRefinedXML = Replace(commandRefinedXML, "+", " plus ")
	commandRefinedXML = Replace(commandRefinedXML, "#", "%23")
	commandRefinedXML = Replace(commandRefinedXML, "%", "%25")
	commandRefinedXML = Replace(commandRefinedXML, "^", "%5E")
	commandRefinedXML = Replace(commandRefinedXML, "&", "%26")
	commandRefinedXML = Replace(commandRefinedXML, "=", "%3D")
	commandRefinedXML = Replace(commandRefinedXML, " ", "+")
	commandRefinedXML = Replace(commandRefinedXML, "\", "%5C")
	commandRefinedXML = Replace(commandRefinedXML, "|", "%7C")
	commandRefinedXML = Replace(commandRefinedXML, "{", "%7B")
	commandRefinedXML = Replace(commandRefinedXML, "{", "%7D")
	commandRefinedXML = Replace(commandRefinedXML, "[", "%5B")
	commandRefinedXML = Replace(commandRefinedXML, "]", "%5D")
	commandRefinedXML = Replace(commandRefinedXML, "'", "%27")
	commandRefinedXML = Replace(commandRefinedXML, "?", "%3F")
	commandRefinedXML = Replace(commandRefinedXML, "/", "%2F")
	fullXMLUrl = startingURL + commandRefinedXML + endingURL
	WinHttpReq.Open "GET", fullXMLURL, False

	WinHttpReq.Send
	If (WinHttpReq.Status = 200) Then

		  correctString=""
		  sTemp = WinHttpReq.ResponseText

		  sTemp = Replace(sTemp, vbCr, vbcrlf)
		  sTemp = Replace(sTemp, vbLf, vbcrlf)
		 
		   WebText = sTemp

		   Set objFile = objFSO.OpenTextFile("C:\Sherby Interface\xmlWolframFile.txt", 2, True)
		   on error resume next
		   objFile.WriteLine sTemp
			If Err Then
				WScript.StdErr.WriteLine "error "
			End If 
		   objFile.Close

			Set objFileXML = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\xmlWolframFile.txt",1)
			splitCommand = Split(commandTimeFor, " ")
			indexOfCityName = inArray(splitCommand, "for") + 1
			city = splitCommand(indexOfCityName)
			Sapi.speak "The current time in " + city + "is"
			do while not objFileXML.AtEndOfStream
				plaintextSearch = objFileXML.ReadLine()
				If InStr(1, plaintextSearch, "<plaintext>") > 0 Then
					plaintextSearch = Replace(plaintextSearch, "<plaintext>", "")
					plaintextSearch = Replace(plaintextSearch, "</plaintext>", "")
					If Trim(plaintextSearch) = "" OR plaintextSearch = " " OR plaintextSearch = "  " OR plaintextSearch = "   " Then
						Sapi.speak "Unfortunately, the time in " + city + " is currently unavailable"
					Else
						If InStr(1, plaintextSearch, "pm") > 0 Then
							endingTimeIndex = InStr(1, plaintextSearch, "pm")
						ElseIf InStr(1, plaintextSearch, "am") > 0 Then
							endingTimeIndex = InStr(1, plaintextSearch, "am")
						Else
							Sapi.speak "Unfortunately, the time in " + city + " is currently unavailable"
						End If
						plaintextSearch = Left(plaintextSearch, endingTimeIndex + 1)
						endingTimeIndex = Right(plaintextSearch, 3)
						timeWithoutIndex = Left(plaintextSearch, Len(plaintextSearch) - 6)
						plaintextSearch = timeWithoutIndex + endingTimeIndex
						Sapi.speak plaintextSearch
					End If
					foundplaintext = 1
				ElseIf InStr(1, plaintextSearch, "didyoumean") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<tip ") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<error>") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<plaintext/>") > 0 And foundError=0 OR InStr(1, plaintextSearch, "not available") > 0 And foundError=0 Then
					Sapi.speak "Unfortunately, the time in " + city + " is currently unavailable"
				End If
			loop
			objFileXML.Close
			If foundplaintext=0 Then
				Sapi.speak "Unfortunately, the time in " + city + " is currently unavailable"
			End If
	Else
		 Sapi.speak "Unfortunately, I'm unable to answer your question at the moment. Please make sure that your internet connection is secure and that airplane mode is not turned on"
	End If
End Function

Function doWolframDateTimeSearch(commandDateTime)
	isplaneOverhead = 0
	Dim WinHttpReq
	Set WinHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
	Set objFSO=CreateObject("Scripting.FileSystemObject")
	Dim sTemp
	foundError=0
	foundplaintext=0
	startingURL = "http://api.wolframalpha.com/v2/query?input="
	endingURL = "&appid=AAXTLK-5XE3J9E8L3&location=Ottawa,%20Ontario&podindex=2&format=plaintext"
	commandRefinedXML = commandDateTime
	commandRefinedXML = Replace(commandRefinedXML, "+", " plus ")
	commandRefinedXML = Replace(commandRefinedXML, "#", "%23")
	commandRefinedXML = Replace(commandRefinedXML, "%", "%25")
	commandRefinedXML = Replace(commandRefinedXML, "^", "%5E")
	commandRefinedXML = Replace(commandRefinedXML, "&", "%26")
	commandRefinedXML = Replace(commandRefinedXML, "=", "%3D")
	commandRefinedXML = Replace(commandRefinedXML, " ", "+")
	commandRefinedXML = Replace(commandRefinedXML, "\", "%5C")
	commandRefinedXML = Replace(commandRefinedXML, "|", "%7C")
	commandRefinedXML = Replace(commandRefinedXML, "{", "%7B")
	commandRefinedXML = Replace(commandRefinedXML, "{", "%7D")
	commandRefinedXML = Replace(commandRefinedXML, "[", "%5B")
	commandRefinedXML = Replace(commandRefinedXML, "]", "%5D")
	commandRefinedXML = Replace(commandRefinedXML, "'", "%27")
	commandRefinedXML = Replace(commandRefinedXML, "?", "%3F")
	commandRefinedXML = Replace(commandRefinedXML, "/", "%2F")
	fullXMLUrl = startingURL + commandRefinedXML + endingURL
	WinHttpReq.Open "GET", fullXMLURL, False

	WinHttpReq.Send
	If (WinHttpReq.Status = 200) Then

		  correctString=""
		  sTemp = WinHttpReq.ResponseText

		  sTemp = Replace(sTemp, vbCr, vbcrlf)
		  sTemp = Replace(sTemp, vbLf, vbcrlf)
		 
		   WebText = sTemp
			splitCommand = Split(commandDateTime, " ")
			indexOfCityName = inArray(splitCommand, "in") + 1
			city = splitCommand(indexOfCityName)
			Sapi.speak "In " + city + ", it is..."
		   Set objFile = objFSO.OpenTextFile("C:\Sherby Interface\xmlWolframFile.txt", 2, True)
		   on error resume next
		   objFile.WriteLine sTemp
			If Err Then
				WScript.StdErr.WriteLine "error "
			End If 
		   objFile.Close

			Set objFileXML = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\xmlWolframFile.txt",1)
			do while not objFileXML.AtEndOfStream
				plaintextSearch = objFileXML.ReadLine()
				If InStr(1, plaintextSearch, "<plaintext>") > 0 Then
					plaintextSearch = Replace(plaintextSearch, "<plaintext>", "")
					plaintextSearch = Replace(plaintextSearch, "</plaintext>", "")
					If Trim(plaintextSearch) = "" OR plaintextSearch = " " OR plaintextSearch = "  " OR plaintextSearch = "   " Then
						Sapi.speak "Unfortunately, the date and time in " + city + " is currently unavailable"
					Else
						currentDateInCity = Right(plaintextSearch, Len(plaintextSearch) - InStr(1, plaintextSearch, "|"))
						If InStr(1, plaintextSearch, "pm") > 0 Then
							endingTimeIndex = InStr(1, plaintextSearch, "pm")
						ElseIf InStr(1, plaintextSearch, "am") > 0 Then
							endingTimeIndex = InStr(1, plaintextSearch, "am")
						Else
							Sapi.speak "Unfortunately, the date and time in " + city + " is currently unavailable"
						End If
						plaintextSearchTime = Left(plaintextSearch, endingTimeIndex + 1)
						endingTimeIndex = Right(plaintextSearchTime, 3)
						timeWithoutIndex = Left(plaintextSearchTime, Len(plaintextSearchTime) - 6)
						plaintextSearchTimeFull = timeWithoutIndex + endingTimeIndex
						Sapi.speak currentDateInCity + ", and the current time in " + city + " is " + plaintextSearchTimeFull
					End If
					foundplaintext = 1
				ElseIf InStr(1, plaintextSearch, "didyoumean") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<tip ") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<error>") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<plaintext/>") > 0 And foundError=0 OR InStr(1, plaintextSearch, "not available") > 0 And foundError=0 Then
					Sapi.speak "Unfortunately, the date and time in " + city + " is currently unavailable"
				End If
			loop
			objFileXML.Close
			If foundplaintext=0 Then
				Sapi.speak "Unfortunately, the date and time in " + city + " is currently unavailable"
			End If
	Else
		 Sapi.speak "Unfortunately, I'm unable to answer your question at the moment. Please make sure that your internet connection is secure and that airplane mode is not turned on"
	End If
End Function


Function doWolframDateTimeSearchFor(commandDateTimeFor)
	isplaneOverhead = 0
	Dim WinHttpReq
	Set WinHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
	Set objFSO=CreateObject("Scripting.FileSystemObject")
	Dim sTemp
	foundError=0
	foundplaintext=0
	startingURL = "http://api.wolframalpha.com/v2/query?input="
	endingURL = "&appid=AAXTLK-5XE3J9E8L3&location=Ottawa,%20Ontario&podindex=2&format=plaintext"
	commandRefinedXML = commandDateTimeFor
	commandRefinedXML = Replace(commandRefinedXML, "+", " plus ")
	commandRefinedXML = Replace(commandRefinedXML, "#", "%23")
	commandRefinedXML = Replace(commandRefinedXML, "%", "%25")
	commandRefinedXML = Replace(commandRefinedXML, "^", "%5E")
	commandRefinedXML = Replace(commandRefinedXML, "&", "%26")
	commandRefinedXML = Replace(commandRefinedXML, "=", "%3D")
	commandRefinedXML = Replace(commandRefinedXML, " ", "+")
	commandRefinedXML = Replace(commandRefinedXML, "\", "%5C")
	commandRefinedXML = Replace(commandRefinedXML, "|", "%7C")
	commandRefinedXML = Replace(commandRefinedXML, "{", "%7B")
	commandRefinedXML = Replace(commandRefinedXML, "{", "%7D")
	commandRefinedXML = Replace(commandRefinedXML, "[", "%5B")
	commandRefinedXML = Replace(commandRefinedXML, "]", "%5D")
	commandRefinedXML = Replace(commandRefinedXML, "'", "%27")
	commandRefinedXML = Replace(commandRefinedXML, "?", "%3F")
	commandRefinedXML = Replace(commandRefinedXML, "/", "%2F")
	fullXMLUrl = startingURL + commandRefinedXML + endingURL
	WinHttpReq.Open "GET", fullXMLURL, False

	WinHttpReq.Send
	If (WinHttpReq.Status = 200) Then

		  correctString=""
		  sTemp = WinHttpReq.ResponseText

		  sTemp = Replace(sTemp, vbCr, vbcrlf)
		  sTemp = Replace(sTemp, vbLf, vbcrlf)
		 
		   WebText = sTemp
			splitCommand = Split(commandDateTimeFor, " ")
			indexOfCityName = inArray(splitCommand, "for") + 1
			city = splitCommand(indexOfCityName)
			Sapi.speak "In " + city + ", it is..."
		   Set objFile = objFSO.OpenTextFile("C:\Sherby Interface\xmlWolframFile.txt", 2, True)
		   on error resume next
		   objFile.WriteLine sTemp
			If Err Then
				WScript.StdErr.WriteLine "error "
			End If 
		   objFile.Close

			Set objFileXML = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\xmlWolframFile.txt",1)
			do while not objFileXML.AtEndOfStream
				plaintextSearch = objFileXML.ReadLine()
				If InStr(1, plaintextSearch, "<plaintext>") > 0 Then
					plaintextSearch = Replace(plaintextSearch, "<plaintext>", "")
					plaintextSearch = Replace(plaintextSearch, "</plaintext>", "")
					If Trim(plaintextSearch) = "" OR plaintextSearch = " " OR plaintextSearch = "  " OR plaintextSearch = "   " Then
						Sapi.speak "Unfortunately, the date and time in " + city + " is currently unavailable"
					Else
						currentDateInCity = Right(plaintextSearch, Len(plaintextSearch) - InStr(1, plaintextSearch, "|"))
						If InStr(1, plaintextSearch, "pm") > 0 Then
							endingTimeIndex = InStr(1, plaintextSearch, "pm")
						ElseIf InStr(1, plaintextSearch, "am") > 0 Then
							endingTimeIndex = InStr(1, plaintextSearch, "am")
						Else
							Sapi.speak "Unfortunately, the date and time in " + city + " is currently unavailable"
						End If
						plaintextSearchTime = Left(plaintextSearch, endingTimeIndex + 1)
						endingTimeIndex = Right(plaintextSearchTime, 3)
						timeWithoutIndex = Left(plaintextSearchTime, Len(plaintextSearchTime) - 6)
						plaintextSearchTimeFull = timeWithoutIndex + endingTimeIndex
						Sapi.speak currentDateInCity + ", and the current time in " + city + " is " + plaintextSearchTimeFull
					End If
					foundplaintext = 1
				ElseIf InStr(1, plaintextSearch, "didyoumean") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<tip ") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<error>") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<plaintext/>") > 0 And foundError=0 OR InStr(1, plaintextSearch, "not available") > 0 And foundError=0 Then
					Sapi.speak "Unfortunately, the date and time in " + city + " is currently unavailable"
				End If
			loop
			objFileXML.Close
			If foundplaintext=0 Then
				Sapi.speak "Unfortunately, the date and time in " + city + " is currently unavailable"
			End If
	Else
		 Sapi.speak "Unfortunately, I'm unable to answer your question at the moment. Please make sure that your internet connection is secure and that airplane mode is not turned on"
	End If
End Function

isPlaneOverhead=0
Function foundplaneWolfram(commandPlane)
	Dim WinHttpReq
	Set WinHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
	Set objFSO=CreateObject("Scripting.FileSystemObject")
	Dim sTemp
	foundError=0
	foundplaintext=0

	startingURL = "http://api.wolframalpha.com/v2/query?input="
	endingURL = "&appid=AAXTLK-5XE3J9E8L3&location=Ottawa,%20Ontario&podindex=1&format=plaintext"
	commandRefinedXML = commandPlane
	commandRefinedXML = Replace(commandRefinedXML, "+", " plus ")
	commandRefinedXML = Replace(commandRefinedXML, "#", "%23")
	commandRefinedXML = Replace(commandRefinedXML, "%", "%25")
	commandRefinedXML = Replace(commandRefinedXML, "^", "%5E")
	commandRefinedXML = Replace(commandRefinedXML, "&", "%26")
	commandRefinedXML = Replace(commandRefinedXML, "=", "%3D")
	commandRefinedXML = Replace(commandRefinedXML, " ", "+")
	commandRefinedXML = Replace(commandRefinedXML, "\", "%5C")
	commandRefinedXML = Replace(commandRefinedXML, "|", "%7C")
	commandRefinedXML = Replace(commandRefinedXML, "{", "%7B")
	commandRefinedXML = Replace(commandRefinedXML, "{", "%7D")
	commandRefinedXML = Replace(commandRefinedXML, "[", "%5B")
	commandRefinedXML = Replace(commandRefinedXML, "]", "%5D")
	commandRefinedXML = Replace(commandRefinedXML, "'", "%27")
	commandRefinedXML = Replace(commandRefinedXML, "?", "%3F")
	commandRefinedXML = Replace(commandRefinedXML, "/", "%2F")
	fullXMLUrl = startingURL + commandRefinedXML + endingURL
	WinHttpReq.Open "GET", fullXMLURL, False

	WinHttpReq.Send
	 If (WinHttpReq.Status = 200) Then

		  correctString=""
		  sTemp = WinHttpReq.ResponseText

		  sTemp = Replace(sTemp, vbCr, vbcrlf)
		  sTemp = Replace(sTemp, vbLf, vbcrlf)
		 
		   WebText = sTemp

		   Set objFile = objFSO.OpenTextFile("C:\Sherby Interface\xmlWolframFile.txt", 2, True)
		   on error resume next
		   objFile.WriteLine sTemp
If Err Then
    WScript.StdErr.WriteLine "error "
End If 
		   objFile.Close

			Set objFileXML = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\xmlWolframFile.txt",1)

			do while not objFileXML.AtEndOfStream
				plaintextSearch = objFileXML.ReadLine()
				If InStr(1, plaintextSearch, "<plaintext>") > 0 Then
					plaintextSearch = Replace(plaintextSearch, "<plaintext>", "")
					plaintextSearch = Replace(plaintextSearch, "</plaintext>", "")
					If InStr(1, plaintextSearch, "flights seen from current geoIP location") > 0 Then
						isplaneOverhead = 1
					Else
						isplaneOverhead = 0
					End If
					foundplaintext = 1
				End If
			loop
			objFileXML.Close
			If foundplaintext=0 Then
				foundplaneWolfram = False
			End If

	  Else
		 Sapi.speak "Unfortunately, I'm unable to answer your question at the moment. Please make sure that your internet connection is secure and that airplane mode is not turned on"
	End If
End Function

Function doWolframSearch(commandOfficial)
	foundplaneWolfram(commandOfficial)
	Dim WinHttpReq
	Set WinHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
	Set objFSO=CreateObject("Scripting.FileSystemObject")
	Dim sTemp
	foundError=0
	foundplaintext=0
	startingURL = "http://api.wolframalpha.com/v2/query?input="
	If isplaneOverhead = 1 Then
		endingURL = "&appid=AAXTLK-5XE3J9E8L3&location=Ottawa,%20Ontario&podindex=2&format=image"
	Else
		endingURL = "&appid=AAXTLK-5XE3J9E8L3&location=Ottawa,%20Ontario&podindex=2&format=plaintext"
	End If
	commandRefinedXML = commandOfficial
	commandRefinedXML = Replace(commandRefinedXML, "+", " plus ")
	commandRefinedXML = Replace(commandRefinedXML, "#", "%23")
	commandRefinedXML = Replace(commandRefinedXML, "%", "%25")
	commandRefinedXML = Replace(commandRefinedXML, "^", "%5E")
	commandRefinedXML = Replace(commandRefinedXML, "&", "%26")
	commandRefinedXML = Replace(commandRefinedXML, "=", "%3D")
	commandRefinedXML = Replace(commandRefinedXML, " ", "+")
	commandRefinedXML = Replace(commandRefinedXML, "\", "%5C")
	commandRefinedXML = Replace(commandRefinedXML, "|", "%7C")
	commandRefinedXML = Replace(commandRefinedXML, "{", "%7B")
	commandRefinedXML = Replace(commandRefinedXML, "{", "%7D")
	commandRefinedXML = Replace(commandRefinedXML, "[", "%5B")
	commandRefinedXML = Replace(commandRefinedXML, "]", "%5D")
	commandRefinedXML = Replace(commandRefinedXML, "'", "%27")
	commandRefinedXML = Replace(commandRefinedXML, "?", "%3F")
	commandRefinedXML = Replace(commandRefinedXML, "/", "%2F")
	fullXMLUrl = startingURL + commandRefinedXML + endingURL
	WinHttpReq.Open "GET", fullXMLURL, False

	WinHttpReq.Send
	 If (WinHttpReq.Status = 200) Then
		  correctString=""
		  sTemp = WinHttpReq.ResponseText

		  sTemp = Replace(sTemp, vbCr, vbcrlf)
		  sTemp = Replace(sTemp, vbLf, vbcrlf)
		 
		   WebText = sTemp

		   Set objFile = objFSO.OpenTextFile("C:\Sherby Interface\xmlWolframFile.txt", 2, True)
		   on error resume next
		   objFile.WriteLine sTemp
			If Err Then
				WScript.StdErr.WriteLine "error "
			End If 
		   objFile.Close

			Set objFileXML = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\xmlWolframFile.txt",1)

		If isplaneOverhead = 0 Then
			do while not objFileXML.AtEndOfStream
				plaintextSearch = objFileXML.ReadLine()
				If InStr(1, plaintextSearch, "<plaintext>") > 0 Then
					plaintextSearch = Replace(plaintextSearch, "<plaintext>", "")
					plaintextSearch = Replace(plaintextSearch, "</plaintext>", "")
					If Trim(plaintextSearch) = "" OR plaintextSearch = " " OR plaintextSearch = "  " OR plaintextSearch = "   " OR InStr(1, plaintextSearch, "(data not ") Then
						Sapi.speak Replace(rand(Array("Okay, here's what I found for, #%command%#, on google", "Okay, here's what I found on the web for, #%command%#.", "Okay, I found this on the web for, #%command%#", "Okay, check it out. Here's what I found for, #%command%#, on google", "Okay, check it out. Here's what I found for, #%command%#, on the web", "I found this on google for, #%command%#", "Here's what I found for, #%command%#, on the web")), "#%command%#", commandOfficial)
						googleSearchQueryRefined=Trim(commandOfficial)
						googleSearchQueryRefined=Replace(googleSearchQueryRefined, "+", "%2B")
						googleSearchQueryRefined=Replace(googleSearchQueryRefined, "#", "%23")
						googleSearchQueryRefined=Replace(googleSearchQueryRefined, "%", "%25")
						googleSearchQueryRefined=Replace(googleSearchQueryRefined, "&", "%26")
						googleSearchQueryRefined=Replace(googleSearchQueryRefined, " ", "%20")
						startingurl="https://www.google.ca/?gws_rd=cr&ei=acGMV629B8TQ-QH6vZroDQ&safe=active&ssui=on#safe=active&q="
						fullurl=startingurl+googleSearchQueryRefined
						a.run fullurl
						foundError = 1
					Else
						plaintextSplitByLetter = Array()
						x=Len(plaintextSearch)
						for i=0 to x
							letterToAdd = Mid(plaintextSearch,i+1,1)
							ReDim Preserve plaintextSplitByLetter(UBound(plaintextSplitByLetter) + 1)
							plaintextSplitByLetter(UBound(plaintextSplitByLetter)) = letterToAdd
						Next
						plaintextSearch = Replace(plaintextSearch, "au (astronomical units)", "astronomical units")
						plaintextSearch = Replace(plaintextSearch, "^", " to the power of ")
						plaintextSearch = Replace(plaintextSearch, "|", ",")
						plaintextSearch = Replace(plaintextSearch, "(degrees celsius)", "")
						plaintextSearch = Replace(plaintextSearch, "(meters per second)", "")
						plaintextSearch = Replace(plaintextSearch, "m/s", " meters per second")
						plaintextSearch = Replace(plaintextSearch, " N ", " north ")
						plaintextSearch = Replace(plaintextSearch, " S ", " south ")
						plaintextSearch = Replace(plaintextSearch, " E ", " east ")
						plaintextSearch = Replace(plaintextSearch, " W ", " west ")
						plaintextSearch = Replace(plaintextSearch, " NE ", " north east ")
						plaintextSearch = Replace(plaintextSearch, " NW ", " north west ")
						plaintextSearch = Replace(plaintextSearch, " SE ", " south east ")
						plaintextSearch = Replace(plaintextSearch, " SW ", " south west ")
						plaintextSearch = Replace(plaintextSearch, " NNE ", " north north east ")
						plaintextSearch = Replace(plaintextSearch, " NNW ", " north north west ")
						plaintextSearch = Replace(plaintextSearch, " ENE ", " east north east ")
						plaintextSearch = Replace(plaintextSearch, " WNW ", " west north west ")
						plaintextSearch = Replace(plaintextSearch, " SSE ", " south south east ")
						plaintextSearch = Replace(plaintextSearch, " SSW ", " south south west ")
						plaintextSearch = Replace(plaintextSearch, " ESE ", " east south east ")
						plaintextSearch = Replace(plaintextSearch, " WSW ", " west south west ")
						plaintextSearch = Replace(plaintextSearch, "  ", " ")
						plaintextSearch = Replace(plaintextSearch, "&quot;", """")
						If Right(plaintextSearch, 1) = ")" And Left(plaintextSearch, 1) <> "(" Then
							lastOcurrenceOfBracket = InStrRev(plaintextSearch,"(",-1)
							plaintextSearch = Left(plaintextSearch, lastOcurrenceOfBracket - 1)
						Else
							wscript.sleep(0)
						End If
						Dim day
						Dim daySplit
						day = CStr(plaintextSearch)
						daySplit = Split(day, "-")
						daySplitNum = UBound(daySplit)
						If InStr(1, day, "-") > 0 And daySplitNum = 2 Then
							If isNumeric(daySplit(0)) And isNumeric(daySplit(1)) And isNumeric(daySplit(2)) Then
								convertDate(day)
							Else
								If UBound(Split(plaintextSearch, " ")) < 15 Then
									If UBound(Split(plaintextSearch, " ")) < 10 Then
										Sapi.speak plaintextSearch
									Else
										Set VObj = CreateObject("SAPI.SpVoice")

										With VObj
											.Volume = 100
											.Rate = -3
											.Speak plaintextSearch
										End With
									End If
								End If
								If UBound(Split(plaintextSearch, " ")) > 10 OR InStr(1, command, "where") > 0 OR InStr(1, command, "when") > 0 Then
									plaintextSearch = LTrim(plaintextSearch)
									wscriptMessageAnswer = msgbox("Answer: " + PCase(plaintextSearch),,"Answer to your question")
								End If
							End If
						Else
							If UBound(Split(plaintextSearch, " ")) < 15 Then
								If UBound(Split(plaintextSearch, " ")) < 10 Then
									Sapi.speak plaintextSearch
								Else
									Set VObj = CreateObject("SAPI.SpVoice")

									With VObj
										.Volume = 100
										.Rate = -1
										.Speak plaintextSearch
									End With
								End If
							End If
							If UBound(Split(plaintextSearch, " ")) > 10 OR InStr(1, command, "where") > 0 OR InStr(1, command, "when") > 0 Then
								plaintextSearch = LTrim(plaintextSearch)
								wscriptMessageAnswer = msgbox("ANSWER: " & vbcrlf & "" & vbcrlf & PCase(plaintextSearch),,"Answer to your question")
							End If
						End If
					End If
					foundplaintext = 1
				ElseIf InStr(1, plaintextSearch, "didyoumean") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<tip ") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<error>") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<plaintext/>") > 0 And foundError=0 Then
					Sapi.speak Replace(rand(Array("Okay, here's what I found for, #%command%#, on google", "Okay, here's what I found on the web for, #%command%#.", "Okay, I found this on the web for, #%command%#", "Okay, check it out. Here's what I found for, #%command%#, on google", "Okay, check it out. Here's what I found for, #%command%#, on the web", "I found this on google for, #%command%#", "Here's what I found for, #%command%#, on the web")), "#%command%#", commandOfficial)
					googleSearchQueryRefined=Trim(commandOfficial)
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "+", "%2B")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "#", "%23")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "%", "%25")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "&", "%26")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, " ", "%20")
					startingurl="https://www.google.ca/?gws_rd=cr&ei=acGMV629B8TQ-QH6vZroDQ&safe=active&ssui=on#safe=active&q="
					fullurl=startingurl+googleSearchQueryRefined
					a.run fullurl
					foundError = 1
					foundplaintext = 1
				End If
			loop
			objFileXML.Close
			If foundplaintext=0 Then
				Sapi.speak Replace(rand(Array("Okay, here's what I found for, #%command%#, on google", "Okay, here's what I found on the web for, #%command%#.", "Okay, I found this on the web for, #%command%#", "Okay, check it out. Here's what I found for, #%command%#, on google", "Okay, check it out. Here's what I found for, #%command%#, on the web", "I found this on google for, #%command%#", "Here's what I found for, #%command%#, on the web")), "#%command%#", commandOfficial)
				googleSearchQueryRefined=Trim(commandOfficial)
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "+", "%2B")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "#", "%23")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "%", "%25")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "&", "%26")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, " ", "%20")
				startingurl="https://www.google.ca/?gws_rd=cr&ei=acGMV629B8TQ-QH6vZroDQ&safe=active&ssui=on#safe=active&q="
				fullurl=startingurl+googleSearchQueryRefined
				a.run fullurl
			End If
		Else
			do while not objFileXML.AtEndOfStream
				plaintextSearch = objFileXML.ReadLine()
				If InStr(1, plaintextSearch, "<img src=") > 0 Then
					plaintextSearch = Replace(plaintextSearch, "<img src='", "")
					quotationPosition = InStr(1, plaintextSearch, "'")
					imageURL = CStr(Left(plaintextSearch, quotationPosition - 1))
					imageURL = Replace(imageURL, "amp;", "")
					Sapi.speak "Okay, here's a table of all the airplanes currently above us"
					Set objExplorer = CreateObject("InternetExplorer.Application")
					With objExplorer
						.Visible = 1
						.Toolbar=False
						.Statusbar=False
						.Top=800
						.Left=800
						.Height=400
						.Width=360
						.Navigate imageURL
					End With
					foundplaintext = 1
				ElseIf InStr(1, plaintextSearch, "didyoumean") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<tip ") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<error>") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<plaintext/>") > 0 And foundError=0 Then
					Sapi.speak Replace(rand(Array("Okay, here's what I found for, #%command%#, on google", "Okay, here's what I found on the web for, #%command%#.", "Okay, I found this on the web for, #%command%#", "Okay, check it out. Here's what I found for, #%command%#, on google", "Okay, check it out. Here's what I found for, #%command%#, on the web", "I found this on google for, #%command%#", "Here's what I found for, #%command%#, on the web")), "#%command%#", commandOfficial)
					googleSearchQueryRefined=Trim(commandOfficial)
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "+", "%2B")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "#", "%23")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "%", "%25")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "&", "%26")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, " ", "%20")
					startingurl="https://www.google.ca/?gws_rd=cr&ei=acGMV629B8TQ-QH6vZroDQ&safe=active&ssui=on#safe=active&q="
					fullurl=startingurl+googleSearchQueryRefined
					a.run fullurl
					foundError = 1
					foundplaintext = 1
				End If
			loop
			objFileXML.Close
			If foundplaintext=0 Then
				Sapi.speak Replace(rand(Array("Okay, here's what I found for, #%command%#, on google", "Okay, here's what I found on the web for, #%command%#.", "Okay, I found this on the web for, #%command%#", "Okay, check it out. Here's what I found for, #%command%#, on google", "Okay, check it out. Here's what I found for, #%command%#, on the web", "I found this on google for, #%command%#", "Here's what I found for, #%command%#, on the web")), "#%command%#", commandOfficial)
				googleSearchQueryRefined=Trim(commandOfficial)
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "+", "%2B")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "#", "%23")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "%", "%25")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "&", "%26")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, " ", "%20")
				startingurl="https://www.google.ca/?gws_rd=cr&ei=acGMV629B8TQ-QH6vZroDQ&safe=active&ssui=on#safe=active&q="
				fullurl=startingurl+googleSearchQueryRefined
				a.run fullurl
			End If
		End If

	  Else
		 Sapi.speak "Unfortunately, I'm unable to answer your question at the moment. Please make sure that your internet connection is secure and that airplane mode is not turned on"
	End If
End Function

Function doWolframConvertSearch(commandOfficial)
	foundplaneWolfram(commandOfficial)
	Dim WinHttpReq
	Set WinHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
	Set objFSO=CreateObject("Scripting.FileSystemObject")
	Dim sTemp
	foundError=0
	foundplaintext=0
	startingURL = "http://api.wolframalpha.com/v2/query?input="
	If isplaneOverhead = 1 Then
		endingURL = "&appid=AAXTLK-5XE3J9E8L3&location=Ottawa,%20Ontario&podindex=2&format=image"
	Else
		endingURL = "&appid=AAXTLK-5XE3J9E8L3&location=Ottawa,%20Ontario&podindex=2&format=plaintext"
	End If
	commandRefinedXML = commandOfficial
	commandRefinedXML = Replace(commandRefinedXML, "+", " plus ")
	commandRefinedXML = Replace(commandRefinedXML, "#", "%23")
	commandRefinedXML = Replace(commandRefinedXML, "%", "%25")
	commandRefinedXML = Replace(commandRefinedXML, "^", "%5E")
	commandRefinedXML = Replace(commandRefinedXML, "&", "%26")
	commandRefinedXML = Replace(commandRefinedXML, "=", "%3D")
	commandRefinedXML = Replace(commandRefinedXML, " ", "+")
	commandRefinedXML = Replace(commandRefinedXML, "\", "%5C")
	commandRefinedXML = Replace(commandRefinedXML, "|", "%7C")
	commandRefinedXML = Replace(commandRefinedXML, "{", "%7B")
	commandRefinedXML = Replace(commandRefinedXML, "{", "%7D")
	commandRefinedXML = Replace(commandRefinedXML, "[", "%5B")
	commandRefinedXML = Replace(commandRefinedXML, "]", "%5D")
	commandRefinedXML = Replace(commandRefinedXML, "'", "%27")
	commandRefinedXML = Replace(commandRefinedXML, "?", "%3F")
	commandRefinedXML = Replace(commandRefinedXML, "/", "%2F")
	fullXMLUrl = startingURL + commandRefinedXML + endingURL
	WinHttpReq.Open "GET", fullXMLURL, False

	WinHttpReq.Send
	 If (WinHttpReq.Status = 200) Then
		  correctString=""
		  sTemp = WinHttpReq.ResponseText

		  sTemp = Replace(sTemp, vbCr, vbcrlf)
		  sTemp = Replace(sTemp, vbLf, vbcrlf)
		 
		   WebText = sTemp

		   Set objFile = objFSO.OpenTextFile("C:\Sherby Interface\xmlWolframFile.txt", 2, True)
		   on error resume next
		   objFile.WriteLine sTemp
			If Err Then
				WScript.StdErr.WriteLine "error "
			End If 
		   objFile.Close

			Set objFileXML = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\xmlWolframFile.txt",1)

		If isplaneOverhead = 0 Then
			do while not objFileXML.AtEndOfStream
				plaintextSearch = objFileXML.ReadLine()
				If InStr(1, plaintextSearch, "<plaintext>") > 0 Then
					plaintextSearch = Replace(plaintextSearch, "<plaintext>", "")
					plaintextSearch = Replace(plaintextSearch, "</plaintext>", "")
					If Trim(plaintextSearch) = "" OR plaintextSearch = " " OR plaintextSearch = "  " OR plaintextSearch = "   " Then
						Sapi.speak Replace(rand(Array("Okay, here's what I found for, #%command%#, on google", "Okay, here's what I found on the web for, #%command%#.", "Okay, I found this on the web for, #%command%#", "Okay, check it out. Here's what I found for, #%command%#, on google", "Okay, check it out. Here's what I found for, #%command%#, on the web", "I found this on google for, #%command%#", "Here's what I found for, #%command%#, on the web")), "#%command%#", commandOfficial)
						googleSearchQueryRefined=Trim(commandOfficial)
						googleSearchQueryRefined=Replace(googleSearchQueryRefined, "+", "%2B")
						googleSearchQueryRefined=Replace(googleSearchQueryRefined, "#", "%23")
						googleSearchQueryRefined=Replace(googleSearchQueryRefined, "%", "%25")
						googleSearchQueryRefined=Replace(googleSearchQueryRefined, "&", "%26")
						googleSearchQueryRefined=Replace(googleSearchQueryRefined, " ", "%20")
						startingurl="https://www.google.ca/?gws_rd=cr&ei=acGMV629B8TQ-QH6vZroDQ&safe=active&ssui=on#safe=active&q="
						fullurl=startingurl+googleSearchQueryRefined
						a.run fullurl
						foundError = 1
					Else
						plaintextSplitByLetter = Array()
						x=Len(plaintextSearch)
						plaintextSearchBackup = plaintextSearch
						for i=0 to x
							letterToAdd = Mid(plaintextSearch,i+1,1)
							ReDim Preserve plaintextSplitByLetter(UBound(plaintextSplitByLetter) + 1)
							plaintextSplitByLetter(UBound(plaintextSplitByLetter)) = letterToAdd
						Next
						plaintextSearch = Replace(plaintextSearch, "au (astronomical units)", "astronomical units")
						plaintextSearch = Replace(plaintextSearch, "^", " to the power of ")
						If Right(plaintextSearch, 1) = ")" And Left(plaintextSearch, 1) <> "(" Then
							lastOcurrenceOfBracket = InStrRev(plaintextSearch,"(",-1)
							plaintextSearch = Left(plaintextSearch, lastOcurrenceOfBracket - 1)
						Else
							wscript.sleep(0)
						End If
						Dim day
						Dim daySplit
						day = CStr(plaintextSearch)
						daySplit = Split(day, "-")
						daySplitNum = UBound(daySplit)
						plaintextSearchBackup = Replace(plaintextSearchBackup, "au (astronomical units)", "astronomical units")
						plaintextSearchBackup = Replace(plaintextSearchBackup, "^", " to the power of ")
						plaintextSearchBackup = Replace(plaintextSearchBackup, "|", ",")
						plaintextSearchBackup = Replace(plaintextSearchBackup, "(degrees celsius)", "")
						plaintextSearchBackup = Replace(plaintextSearchBackup, "(meters per second)", "")
						plaintextSearchBackup = Replace(plaintextSearchBackup, "m/s", " meters per second")
						plaintextSearchBackup = Replace(plaintextSearchBackup, "km", "kilometers")
						plaintextSearchBackup = Replace(plaintextSearchBackup, " N ", " north ")
						plaintextSearchBackup = Replace(plaintextSearchBackup, " S ", " south ")
						plaintextSearchBackup = Replace(plaintextSearchBackup, " E ", " east ")
						plaintextSearchBackup = Replace(plaintextSearchBackup, " W ", " west ")
						plaintextSearchBackup = Replace(plaintextSearchBackup, " NE ", " north east ")
						plaintextSearchBackup = Replace(plaintextSearchBackup, " NW ", " north west ")
						plaintextSearchBackup = Replace(plaintextSearchBackup, " SE ", " south east ")
						plaintextSearchBackup = Replace(plaintextSearchBackup, " SW ", " south west ")
						plaintextSearchBackup = Replace(plaintextSearchBackup, " NNE ", " north north east ")
						plaintextSearchBackup = Replace(plaintextSearchBackup, " NNW ", " north north west ")
						plaintextSearchBackup = Replace(plaintextSearchBackup, " ENE ", " east north east ")
						plaintextSearchBackup = Replace(plaintextSearchBackup, " WNW ", " west north west ")
						plaintextSearchBackup = Replace(plaintextSearchBackup, " SSE ", " south south east ")
						plaintextSearchBackup = Replace(plaintextSearchBackup, " SSW ", " south south west ")
						plaintextSearchBackup = Replace(plaintextSearchBackup, " ESE ", " east south east ")
						plaintextSearchBackup = Replace(plaintextSearchBackup, " WSW ", " west south west ")
						If InStr(1, day, "-") > 0 And daySplitNum = 2 Then
							If isNumeric(daySplit(0)) And isNumeric(daySplit(1)) And isNumeric(daySplit(2)) Then
								convertDate(day)
							Else
								Sapi.speak "The answer is " + CStr(plaintextSearch)
								answerConvert = msgbox(plaintextSearchBackup,,"Conversion answer")
							End If
						Else
							Sapi.speak "The answer is " + CStr(plaintextSearch)
							answerConvertTwo = msgbox(plaintextSearchBackup,,"Conversion answer")
						End If
					End If
					foundplaintext = 1
				ElseIf InStr(1, plaintextSearch, "didyoumean") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<tip ") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<error>") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<plaintext/>") > 0 And foundError=0 Then
					Sapi.speak Replace(rand(Array("Okay, here's what I found for, #%command%#, on google", "Okay, here's what I found on the web for, #%command%#.", "Okay, I found this on the web for, #%command%#", "Okay, check it out. Here's what I found for, #%command%#, on google", "Okay, check it out. Here's what I found for, #%command%#, on the web", "I found this on google for, #%command%#", "Here's what I found for, #%command%#, on the web")), "#%command%#", commandOfficial)
					googleSearchQueryRefined=Trim(commandOfficial)
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "+", "%2B")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "#", "%23")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "%", "%25")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "&", "%26")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, " ", "%20")
					startingurl="https://www.google.ca/?gws_rd=cr&ei=acGMV629B8TQ-QH6vZroDQ&safe=active&ssui=on#safe=active&q="
					fullurl=startingurl+googleSearchQueryRefined
					a.run fullurl
					foundError = 1
					foundplaintext = 1
				End If
			loop
			objFileXML.Close
			If foundplaintext=0 Then
				Sapi.speak Replace(rand(Array("Okay, here's what I found for, #%command%#, on google", "Okay, here's what I found on the web for, #%command%#.", "Okay, I found this on the web for, #%command%#", "Okay, check it out. Here's what I found for, #%command%#, on google", "Okay, check it out. Here's what I found for, #%command%#, on the web", "I found this on google for, #%command%#", "Here's what I found for, #%command%#, on the web")), "#%command%#", commandOfficial)
				googleSearchQueryRefined=Trim(commandOfficial)
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "+", "%2B")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "#", "%23")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "%", "%25")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "&", "%26")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, " ", "%20")
				startingurl="https://www.google.ca/?gws_rd=cr&ei=acGMV629B8TQ-QH6vZroDQ&safe=active&ssui=on#safe=active&q="
				fullurl=startingurl+googleSearchQueryRefined
				a.run fullurl
			End If
		Else
			do while not objFileXML.AtEndOfStream
				plaintextSearch = objFileXML.ReadLine()
				If InStr(1, plaintextSearch, "<img src=") > 0 Then
					plaintextSearch = Replace(plaintextSearch, "<img src='", "")
					quotationPosition = InStr(1, plaintextSearch, "'")
					imageURL = CStr(Left(plaintextSearch, quotationPosition - 1))
					imageURL = Replace(imageURL, "amp;", "")
					Sapi.speak "Okay, here's a table of all the airplanes currently above us"
					Set objExplorer = CreateObject("InternetExplorer.Application")
					With objExplorer
						.Visible = 1
						.Toolbar=False
						.Statusbar=False
						.Top=800
						.Left=800
						.Height=400
						.Width=360
						.Navigate imageURL
					End With
					foundplaintext = 1
				ElseIf InStr(1, plaintextSearch, "didyoumean") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<tip ") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<error>") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<plaintext/>") > 0 And foundError=0 Then
					Sapi.speak Replace(rand(Array("Okay, here's what I found for, #%command%#, on google", "Okay, here's what I found on the web for, #%command%#.", "Okay, I found this on the web for, #%command%#", "Okay, check it out. Here's what I found for, #%command%#, on google", "Okay, check it out. Here's what I found for, #%command%#, on the web", "I found this on google for, #%command%#", "Here's what I found for, #%command%#, on the web")), "#%command%#", commandOfficial)
					googleSearchQueryRefined=Trim(commandOfficial)
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "+", "%2B")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "#", "%23")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "%", "%25")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "&", "%26")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, " ", "%20")
					startingurl="https://www.google.ca/?gws_rd=cr&ei=acGMV629B8TQ-QH6vZroDQ&safe=active&ssui=on#safe=active&q="
					fullurl=startingurl+googleSearchQueryRefined
					a.run fullurl
					foundError = 1
					foundplaintext = 1
				End If
			loop
			objFileXML.Close
			If foundplaintext=0 Then
				Sapi.speak Replace(rand(Array("Okay, here's what I found for, #%command%#, on google", "Okay, here's what I found on the web for, #%command%#.", "Okay, I found this on the web for, #%command%#", "Okay, check it out. Here's what I found for, #%command%#, on google", "Okay, check it out. Here's what I found for, #%command%#, on the web", "I found this on google for, #%command%#", "Here's what I found for, #%command%#, on the web")), "#%command%#", commandOfficial)
				googleSearchQueryRefined=Trim(commandOfficial)
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "+", "%2B")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "#", "%23")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "%", "%25")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "&", "%26")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, " ", "%20")
				startingurl="https://www.google.ca/?gws_rd=cr&ei=acGMV629B8TQ-QH6vZroDQ&safe=active&ssui=on#safe=active&q="
				fullurl=startingurl+googleSearchQueryRefined
				a.run fullurl
			End If
		End If

	  Else
		 Sapi.speak "Unfortunately, I'm unable to answer your question at the moment. Please make sure that your internet connection is secure and that airplane mode is not turned on"
	End If
End Function

Function doWolframNutrientsSearch(commandNutrients)
	foundplaneWolfram(commandNutrients)
	Dim WinHttpReq
	Set WinHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
	Set objFSO=CreateObject("Scripting.FileSystemObject")
	Dim sTemp
	foundError=0
	foundplaintext=0
	startingURL = "http://api.wolframalpha.com/v2/query?input="
	If isplaneOverhead = 1 Then
		endingURL = "&appid=AAXTLK-5XE3J9E8L3&location=Ottawa,%20Ontario&podindex=2&format=image"
	Else
		endingURL = "&appid=AAXTLK-5XE3J9E8L3&location=Ottawa,%20Ontario&podindex=2&format=plaintext"
	End If
	commandRefinedXML = commandNutrients
	commandRefinedXML = Replace(commandRefinedXML, "+", " plus ")
	commandRefinedXML = Replace(commandRefinedXML, "#", "%23")
	commandRefinedXML = Replace(commandRefinedXML, "%", "%25")
	commandRefinedXML = Replace(commandRefinedXML, "^", "%5E")
	commandRefinedXML = Replace(commandRefinedXML, "&", "%26")
	commandRefinedXML = Replace(commandRefinedXML, "=", "%3D")
	commandRefinedXML = Replace(commandRefinedXML, " ", "+")
	commandRefinedXML = Replace(commandRefinedXML, "\", "%5C")
	commandRefinedXML = Replace(commandRefinedXML, "|", "%7C")
	commandRefinedXML = Replace(commandRefinedXML, "{", "%7B")
	commandRefinedXML = Replace(commandRefinedXML, "{", "%7D")
	commandRefinedXML = Replace(commandRefinedXML, "[", "%5B")
	commandRefinedXML = Replace(commandRefinedXML, "]", "%5D")
	commandRefinedXML = Replace(commandRefinedXML, "'", "%27")
	commandRefinedXML = Replace(commandRefinedXML, "?", "%3F")
	commandRefinedXML = Replace(commandRefinedXML, "/", "%2F")
	fullXMLUrl = startingURL + commandRefinedXML + endingURL
	WinHttpReq.Open "GET", fullXMLURL, False

	WinHttpReq.Send
	 If (WinHttpReq.Status = 200) Then
		  correctString=""
		  sTemp = WinHttpReq.ResponseText

		  sTemp = Replace(sTemp, vbCr, vbcrlf)
		  sTemp = Replace(sTemp, vbLf, vbcrlf)
		 
		   WebText = sTemp

		   Set objFile = objFSO.OpenTextFile("C:\Sherby Interface\xmlWolframFile.txt", 2, True)
		   on error resume next
		   objFile.WriteLine sTemp
			If Err Then
				WScript.StdErr.WriteLine "error "
			End If 
		   objFile.Close

			Set objFileXML = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\xmlWolframFile.txt",1)

		If isplaneOverhead = 0 Then
			do while not objFileXML.AtEndOfStream
				plaintextSearch = objFileXML.ReadLine()
				If InStr(1, plaintextSearch, "<plaintext>") > 0 Then
					plaintextSearch = Replace(plaintextSearch, "<plaintext>", "")
					plaintextSearch = Replace(plaintextSearch, "</plaintext>", "")
					If Trim(plaintextSearch) = "" OR plaintextSearch = " " OR plaintextSearch = "  " OR plaintextSearch = "   " Then
						Sapi.speak Replace(rand(Array("Okay, here's what I found for, #%command%#, on google", "Okay, here's what I found on the web for, #%command%#.", "Okay, I found this on the web for, #%command%#", "Okay, check it out. Here's what I found for, #%command%#, on google", "Okay, check it out. Here's what I found for, #%command%#, on the web", "I found this on google for, #%command%#", "Here's what I found for, #%command%#, on the web")), "#%command%#", commandNutrients)
						googleSearchQueryRefined=Trim(commandNutrients)
						googleSearchQueryRefined=Replace(googleSearchQueryRefined, "+", "%2B")
						googleSearchQueryRefined=Replace(googleSearchQueryRefined, "#", "%23")
						googleSearchQueryRefined=Replace(googleSearchQueryRefined, "%", "%25")
						googleSearchQueryRefined=Replace(googleSearchQueryRefined, "&", "%26")
						googleSearchQueryRefined=Replace(googleSearchQueryRefined, " ", "%20")
						startingurl="https://www.google.ca/?gws_rd=cr&ei=acGMV629B8TQ-QH6vZroDQ&safe=active&ssui=on#safe=active&q="
						fullurl=startingurl+googleSearchQueryRefined
						a.run fullurl
						foundError = 1
					Else
						plaintextSplitByLetter = Array()
						x=Len(plaintextSearch)
						for i=0 to x
							letterToAdd = Mid(plaintextSearch,i+1,1)
							ReDim Preserve plaintextSplitByLetter(UBound(plaintextSplitByLetter) + 1)
							plaintextSplitByLetter(UBound(plaintextSplitByLetter)) = letterToAdd
						Next
						If Right(plaintextSearch, 1) = ")" Then
							lastOcurrenceOfBracket = InStrRev(plaintextSearch,"(",-1)
							plaintextSearch = Left(plaintextSearch, lastOcurrenceOfBracket - 1)
						Else
							wscript.sleep(0)
						End If
						Dim day
						Dim daySplit
						day = CStr(plaintextSearch)
						daySplit = Split(day, "-")
						daySplitNum = UBound(daySplit)
						If InStr(1, day, "-") > 0 And daySplitNum = 2 Then
							If isNumeric(daySplit(0)) And isNumeric(daySplit(1)) And isNumeric(daySplit(2)) Then
								convertDate(day)
							Else
								Sapi.speak "The answer is around "
								Sapi.speak plaintextSearch
							End If
						Else
							Sapi.speak "The answer is around "
							Sapi.speak plaintextSearch
						End If
					End If
					foundplaintext = 1
				ElseIf InStr(1, plaintextSearch, "didyoumean") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<tip ") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<error>") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<plaintext/>") > 0 And foundError=0 Then
					Sapi.speak Replace(rand(Array("Okay, here's what I found for, #%command%#, on google", "Okay, here's what I found on the web for, #%command%#.", "Okay, I found this on the web for, #%command%#", "Okay, check it out. Here's what I found for, #%command%#, on google", "Okay, check it out. Here's what I found for, #%command%#, on the web", "I found this on google for, #%command%#", "Here's what I found for, #%command%#, on the web")), "#%command%#", commandNutrients)
					googleSearchQueryRefined=Trim(commandNutrients)
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "+", "%2B")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "#", "%23")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "%", "%25")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "&", "%26")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, " ", "%20")
					startingurl="https://www.google.ca/?gws_rd=cr&ei=acGMV629B8TQ-QH6vZroDQ&safe=active&ssui=on#safe=active&q="
					fullurl=startingurl+googleSearchQueryRefined
					a.run fullurl
					foundError = 1
					foundplaintext = 1
				End If
			loop
			objFileXML.Close
			If foundplaintext=0 Then
				Sapi.speak Replace(rand(Array("Okay, here's what I found for, #%command%#, on google", "Okay, here's what I found on the web for, #%command%#.", "Okay, I found this on the web for, #%command%#", "Okay, check it out. Here's what I found for, #%command%#, on google", "Okay, check it out. Here's what I found for, #%command%#, on the web", "I found this on google for, #%command%#", "Here's what I found for, #%command%#, on the web")), "#%command%#", commandNutrients)
				googleSearchQueryRefined=Trim(commandNutrients)
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "+", "%2B")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "#", "%23")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "%", "%25")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "&", "%26")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, " ", "%20")
				startingurl="https://www.google.ca/?gws_rd=cr&ei=acGMV629B8TQ-QH6vZroDQ&safe=active&ssui=on#safe=active&q="
				fullurl=startingurl+googleSearchQueryRefined
				a.run fullurl
			End If
		Else
			do while not objFileXML.AtEndOfStream
				plaintextSearch = objFileXML.ReadLine()
				If InStr(1, plaintextSearch, "<img src=") > 0 Then
					plaintextSearch = Replace(plaintextSearch, "<img src='", "")
					quotationPosition = InStr(1, plaintextSearch, "'")
					imageURL = CStr(Left(plaintextSearch, quotationPosition - 1))
					imageURL = Replace(imageURL, "amp;", "")
					Sapi.speak "Okay, here's a table of all the airplanes currently above us"
					Set objExplorer = CreateObject("InternetExplorer.Application")
					With objExplorer
						.Visible = 1
						.Toolbar=False
						.Statusbar=False
						.Top=800
						.Left=800
						.Height=400
						.Width=360
						.Navigate imageURL
					End With
					foundplaintext = 1
				ElseIf InStr(1, plaintextSearch, "didyoumean") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<tip ") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<error>") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<plaintext/>") > 0 And foundError=0 Then
					Sapi.speak Replace(rand(Array("Okay, here's what I found for, #%command%#, on google", "Okay, here's what I found on the web for, #%command%#.", "Okay, I found this on the web for, #%command%#", "Okay, check it out. Here's what I found for, #%command%#, on google", "Okay, check it out. Here's what I found for, #%command%#, on the web", "I found this on google for, #%command%#", "Here's what I found for, #%command%#, on the web")), "#%command%#", commandNutrients)
					googleSearchQueryRefined=Trim(commandNutrients)
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "+", "%2B")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "#", "%23")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "%", "%25")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, "&", "%26")
					googleSearchQueryRefined=Replace(googleSearchQueryRefined, " ", "%20")
					startingurl="https://www.google.ca/?gws_rd=cr&ei=acGMV629B8TQ-QH6vZroDQ&safe=active&ssui=on#safe=active&q="
					fullurl=startingurl+googleSearchQueryRefined
					a.run fullurl
					foundError = 1
					foundplaintext = 1
				End If
			loop
			objFileXML.Close
			If foundplaintext=0 Then
				Sapi.speak Replace(rand(Array("Okay, here's what I found for, #%command%#, on google", "Okay, here's what I found on the web for, #%command%#.", "Okay, I found this on the web for, #%command%#", "Okay, check it out. Here's what I found for, #%command%#, on google", "Okay, check it out. Here's what I found for, #%command%#, on the web", "I found this on google for, #%command%#", "Here's what I found for, #%command%#, on the web")), "#%command%#", commandNutrients)
				googleSearchQueryRefined=Trim(commandNutrients)
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "+", "%2B")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "#", "%23")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "%", "%25")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, "&", "%26")
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, " ", "%20")
				startingurl="https://www.google.ca/?gws_rd=cr&ei=acGMV629B8TQ-QH6vZroDQ&safe=active&ssui=on#safe=active&q="
				fullurl=startingurl+googleSearchQueryRefined
				a.run fullurl
			End If
		End If

	  Else
		 Sapi.speak "Unfortunately, I'm unable to answer your question at the moment. Please make sure that your internet connection is secure and that airplane mode is not turned on"
	End If
End Function

Function doWolframWeatherSearch(commandWeather)
	isplaneOverhead = 0
	Dim WinHttpReq
	Set WinHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
	Set objFSO=CreateObject("Scripting.FileSystemObject")
	Dim sTemp
	foundError=0
	foundplaintext=0
	startingURL = "http://api.wolframalpha.com/v2/query?input="
	endingURL = "&appid=AAXTLK-5XE3J9E8L3&location=Ottawa,%20Ontario&podindex=2&format=plaintext"
	commandRefinedXML = commandWeather
	commandRefinedXML = Replace(commandRefinedXML, "+", " plus ")
	commandRefinedXML = Replace(commandRefinedXML, "#", "%23")
	commandRefinedXML = Replace(commandRefinedXML, "%", "%25")
	commandRefinedXML = Replace(commandRefinedXML, "^", "%5E")
	commandRefinedXML = Replace(commandRefinedXML, "&", "%26")
	commandRefinedXML = Replace(commandRefinedXML, "=", "%3D")
	commandRefinedXML = Replace(commandRefinedXML, " ", "+")
	commandRefinedXML = Replace(commandRefinedXML, "\", "%5C")
	commandRefinedXML = Replace(commandRefinedXML, "|", "%7C")
	commandRefinedXML = Replace(commandRefinedXML, "{", "%7B")
	commandRefinedXML = Replace(commandRefinedXML, "{", "%7D")
	commandRefinedXML = Replace(commandRefinedXML, "[", "%5B")
	commandRefinedXML = Replace(commandRefinedXML, "]", "%5D")
	commandRefinedXML = Replace(commandRefinedXML, "'", "%27")
	commandRefinedXML = Replace(commandRefinedXML, "?", "%3F")
	commandRefinedXML = Replace(commandRefinedXML, "/", "%2F")
	fullXMLUrl = startingURL + commandRefinedXML + endingURL
	WinHttpReq.Open "GET", fullXMLURL, False

	WinHttpReq.Send
	If (WinHttpReq.Status = 200) Then

		  correctString=""
		  sTemp = WinHttpReq.ResponseText

		  sTemp = Replace(sTemp, vbCr, vbcrlf)
		  sTemp = Replace(sTemp, vbLf, vbcrlf)
		 
		   WebText = sTemp
			splitCommand = Split(commandWeather, " ")
			indexOfCityName = inArray(splitCommand, "in") + 1
			city = splitCommand(indexOfCityName)
		   Set objFile = objFSO.OpenTextFile("C:\Sherby Interface\xmlWolframFile.txt", 2, True)
		   on error resume next
		   objFile.WriteLine sTemp
			If Err Then
				WScript.StdErr.WriteLine "error "
			End If 
		   objFile.Close

			Set objFileXML = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\xmlWolframFile.txt",1)
			do while not objFileXML.AtEndOfStream
				plaintextSearch = objFileXML.ReadLine()
				If InStr(1, plaintextSearch, "<plaintext>") > 0 Then
					plaintextSearch = Replace(plaintextSearch, "<plaintext>", "")
					plaintextSearch = Replace(plaintextSearch, "</plaintext>", "")
					If Trim(plaintextSearch) = "" OR plaintextSearch = " " OR plaintextSearch = "  " OR plaintextSearch = "   " Then
						Sapi.speak "Unfortunately, the weather in " + city + " is currently unavailable"
					Else
						temperatureRaw = Right(plaintextSearch, Len(plaintextSearch) - InStr(1, plaintextSearch, "|") - 1)
						temperatureFinal = Left(temperatureRaw, InStr(1, temperatureRaw, "(") - 2)
						windChillRaw = Left(temperatureRaw, InStr(1, temperatureRaw, ")") - 1)
						windChillFinal = Mid(windChillRaw, InStr(1, windChillRaw, ":") + 1)
						windChillFinalGood = Replace(windChillFinal, "C", "")
						windChillFinalGood = Replace(windChillFinalGood, "F", "")

						If windChillFinal = temperatureFinal Then
							currentWeatherInCity = "In " + city + ", the current temperature is " + CStr(temperatureFinal) + ", with no wind chill in effect. Finally, outside, it is currently "
						Else
							currentWeatherInCity = "In " + city + ", the current temperature is " + CStr(temperatureFinal) + ", with wind chill making it " + CStr(windChillFinalGood) + ". Finally, outside, it is currently "
						End If
					End If
					foundplaintext = 1
				ElseIf InStr(1, plaintextSearch, "conditions |") > 0 Then
					conditions = Mid(plaintextSearch, InStr(1, plaintextSearch, "|") + 1)
				ElseIf InStr(1, plaintextSearch, "relative humidity | ") > 0 Then
					plaintextSearch = Replace(plaintextSearch, "relative humidity | ", "")
					humidity = Left(plaintextSearch, InStr(1, plaintextSearch, "%"))
					done = 1
				ElseIf InStr(1, plaintextSearch, "didyoumean") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<tip ") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<error>") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<plaintext/>") > 0 And foundError=0 OR InStr(1, plaintextSearch, "not available") > 0 And foundError=0 Then
					Sapi.speak "Unfortunately, the weather in " + city + " is currently unavailable"
				End If

				If done = 1 Then
					finalWeather = currentWeatherInCity + conditions + ", with the humidity being at " + humidity
					Sapi.speak finalWeather
					Sapi.speak "Enjoy your day"
					done = 0
				End If
			loop
			objFileXML.Close
			If foundplaintext=0 Then
				Sapi.speak "Unfortunately, the weather in " + city + " is currently unavailable"
			End If
	Else
		 Sapi.speak "Unfortunately, I'm unable to answer your question at the moment. Please make sure that your internet connection is secure and that airplane mode is not turned on"
	End If
End Function

Function doWolframWeatherSearchFor(commandWeatherFor)
	isplaneOverhead = 0
	Dim WinHttpReq
	Set WinHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
	Set objFSO=CreateObject("Scripting.FileSystemObject")
	Dim sTemp
	foundError=0
	foundplaintext=0
	startingURL = "http://api.wolframalpha.com/v2/query?input="
	endingURL = "&appid=AAXTLK-5XE3J9E8L3&location=Ottawa,%20Ontario&podindex=2&format=plaintext"
	commandRefinedXML = commandWeatherFor
	commandRefinedXML = Replace(commandRefinedXML, "+", " plus ")
	commandRefinedXML = Replace(commandRefinedXML, "#", "%23")
	commandRefinedXML = Replace(commandRefinedXML, "%", "%25")
	commandRefinedXML = Replace(commandRefinedXML, "^", "%5E")
	commandRefinedXML = Replace(commandRefinedXML, "&", "%26")
	commandRefinedXML = Replace(commandRefinedXML, "=", "%3D")
	commandRefinedXML = Replace(commandRefinedXML, " ", "+")
	commandRefinedXML = Replace(commandRefinedXML, "\", "%5C")
	commandRefinedXML = Replace(commandRefinedXML, "|", "%7C")
	commandRefinedXML = Replace(commandRefinedXML, "{", "%7B")
	commandRefinedXML = Replace(commandRefinedXML, "{", "%7D")
	commandRefinedXML = Replace(commandRefinedXML, "[", "%5B")
	commandRefinedXML = Replace(commandRefinedXML, "]", "%5D")
	commandRefinedXML = Replace(commandRefinedXML, "'", "%27")
	commandRefinedXML = Replace(commandRefinedXML, "?", "%3F")
	commandRefinedXML = Replace(commandRefinedXML, "/", "%2F")
	fullXMLUrl = startingURL + commandRefinedXML + endingURL
	WinHttpReq.Open "GET", fullXMLURL, False

	WinHttpReq.Send
	If (WinHttpReq.Status = 200) Then

		  correctString=""
		  sTemp = WinHttpReq.ResponseText

		  sTemp = Replace(sTemp, vbCr, vbcrlf)
		  sTemp = Replace(sTemp, vbLf, vbcrlf)
		 
		   WebText = sTemp
			splitCommand = Split(commandWeatherFor, " ")
			indexOfCityName = inArray(splitCommand, "for") + 1
			city = splitCommand(indexOfCityName)
		   Set objFile = objFSO.OpenTextFile("C:\Sherby Interface\xmlWolframFile.txt", 2, True)
		   on error resume next
		   objFile.WriteLine sTemp
			If Err Then
				WScript.StdErr.WriteLine "error "
			End If 
		   objFile.Close

			Set objFileXML = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\xmlWolframFile.txt",1)
			do while not objFileXML.AtEndOfStream
				plaintextSearch = objFileXML.ReadLine()
				If InStr(1, plaintextSearch, "<plaintext>") > 0 Then
					plaintextSearch = Replace(plaintextSearch, "<plaintext>", "")
					plaintextSearch = Replace(plaintextSearch, "</plaintext>", "")
					If Trim(plaintextSearch) = "" OR plaintextSearch = " " OR plaintextSearch = "  " OR plaintextSearch = "   " Then
						Sapi.speak "Unfortunately, the weather in " + city + " is currently unavailable"
					Else
						temperatureRaw = Right(plaintextSearch, Len(plaintextSearch) - InStr(1, plaintextSearch, "|") - 1)
						temperatureFinal = Left(temperatureRaw, InStr(1, temperatureRaw, "(") - 2)
						windChillRaw = Left(temperatureRaw, InStr(1, temperatureRaw, ")") - 1)
						windChillFinal = Mid(windChillRaw, InStr(1, windChillRaw, ":") + 1)
						windChillFinalGood = Replace(windChillFinal, "C", "")
						windChillFinalGood = Replace(windChillFinalGood, "F", "")

						If windChillFinal = temperatureFinal Then
							currentWeatherInCity = "In " + city + ", the current temperature is " + CStr(temperatureFinal) + ", with no wind chill in effect. Finally, outside, it is currently "
						Else
							currentWeatherInCity = "In " + city + ", the current temperature is " + CStr(temperatureFinal) + ", with wind chill making it " + CStr(windChillFinalGood) + ". Finally, outside, it is currently "
						End If
					End If
					foundplaintext = 1
				ElseIf InStr(1, plaintextSearch, "conditions |") > 0 Then
					conditions = Mid(plaintextSearch, InStr(1, plaintextSearch, "|") + 1)
				ElseIf InStr(1, plaintextSearch, "relative humidity | ") > 0 Then
					plaintextSearch = Replace(plaintextSearch, "relative humidity | ", "")
					humidity = Left(plaintextSearch, InStr(1, plaintextSearch, "%"))
					done = 1
				ElseIf InStr(1, plaintextSearch, "didyoumean") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<tip ") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<error>") > 0 And foundError=0 OR InStr(1, plaintextSearch, "<plaintext/>") > 0 And foundError=0 OR InStr(1, plaintextSearch, "not available") > 0 And foundError=0 Then
					Sapi.speak "Unfortunately, the weather in " + city + " is currently unavailable"
				End If

				If done = 1 Then
					finalWeather = currentWeatherInCity + conditions + ", with the humidity being at " + humidity
					Sapi.speak finalWeather
					Sapi.speak "Enjoy your day"
					done = 0
				End If
			loop
			objFileXML.Close
			If foundplaintext=0 Then
				Sapi.speak "Unfortunately, the weather in " + city + " is currently unavailable"
			End If
	Else
		 Sapi.speak "Unfortunately, I'm unable to answer your question at the moment. Please make sure that your internet connection is secure and that airplane mode is not turned on"
	End If
End Function

Function getWeather()
	Dim WinHttpReq
	Set WinHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
	Set objFSO=CreateObject("Scripting.FileSystemObject")
	Dim currTemp
	temperature = ""
	min_temp = ""
	max_temp = ""
	wind = ""
	weather_status = ""
	humidity = ""

	WinHttpReq.Open "GET", "http://api.openweathermap.org/data/2.5/weather?q=Ottawa,ca&appid=8f75d3f92b3ff1837a312fdb624d6348&mode=xml&units=metric", False

	WinHttpReq.Send
	 If (WinHttpReq.Status = 200) Then
		  currTemp = WinHttpReq.ResponseText

		  currTemp = Replace(currTemp, vbCr, vbcrlf)
		  currTemp = Replace(currTemp, vbLf, vbcrlf)
		   WebText = currTemp

		   Set objFile = objFSO.OpenTextFile("C:\Sherby Interface\weather.txt", 2, True)
		   on error resume next
		   objFile.WriteLine currTemp
			If Err Then
				WScript.StdErr.WriteLine "error "
			End If 
		   objFile.Close
			Set objFileXML = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\weather.txt",1)
			do while not objFileXML.AtEndOfStream
				weatherSearch = objFileXML.ReadAll
				backupWeatherSearch = weatherSearch
				weatherSearch = Mid(weatherSearch, InStr(1, weatherSearch, "<temperature value=""") + 20)
				temperature = Left(weatherSearch, InStr(1, weatherSearch, """") - 1)
				temperature = CStr(Round(temperature))
				weatherSearch = Mid(weatherSearch, InStr(1, weatherSearch, "min=""") + 5)
				min_temp = Left(weatherSearch, InStr(1, weatherSearch, """") - 1)
				min_temp = min_temp - 8
				min_temp = CStr(Round(min_temp))
				weatherSearch = Mid(weatherSearch, InStr(1, weatherSearch, "max=""") + 5)
				max_temp = Left(weatherSearch, InStr(1, weatherSearch, """") - 1)
				max_temp = CStr(Round(max_temp))
				weatherSearch = Mid(weatherSearch, InStr(1, weatherSearch, "<humidity value=""") + 17)
				humidity = Left(weatherSearch, InStr(1, weatherSearch, """") - 1)
				humidity = CStr(Round(humidity))
				weatherSearch = Mid(weatherSearch, InStr(1, weatherSearch, "name=""") + 6)
				wind = Left(weatherSearch, InStr(1, weatherSearch, """") - 1)
				weatherSearch = Mid(weatherSearch, InStr(1, weatherSearch, "name=""") + 12)
				weatherSearch = Mid(weatherSearch, InStr(1, weatherSearch, "name=""") + 6)
				weather_status = Left(weatherSearch, InStr(1, weatherSearch, """") - 1)
			loop
			objFileXML.Close
			If temperature = "" OR min_temp = "" OR max_temp = "" OR wind = "" OR weather_status = "" OR humidity = "" OR InStr(1, backupWeatherSearch, "Error:") > 0 Then
				Sapi.speak "Unfortunately, the current weather is not available..."
			Else
				If LCase(wind)="calm" Then
					wind = "the weather is calm"
				Else
					wind = "there currently is a " + wind
				End If
				Sapi.speak "In Ottawa, it is currently " + temperature + " degrees celsius, with a high of " + max_temp + " and a low of " + min_temp + ". Outside, " + wind + ", and there is " + weather_status + ". Finally, the humidity is currently at " + humidity + " percent"
				Sapi.speak "Enjoy your day!"
			End If
	  Else
		 Sapi.speak "Unfortunately, I'm unable to answer your question at the moment. Please make sure that your internet connection is secure and that airplane mode is not turned on"
	End If
End Function

Function inArray(arr, obj)
  On Error Resume Next
  Dim x: x = -1
	
  If isArray(arr) Then
    For i = 0 To UBound(arr)
      If arr(i) = obj Then
        x = i
        Exit For
      End If
    Next
  End If
	
  Err.Clear()
  inArray = x

End Function

Function PingSite( myWebsite )
    Dim intStatus, objHTTP

    Set objHTTP = CreateObject( "WinHttp.WinHttpRequest.5.1" )

    objHTTP.Open "GET", "http://" & myWebsite & "/", False
    objHTTP.SetRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MyApp 1.0; Windows NT 5.1)"

    On Error Resume Next

    objHTTP.Send
    intStatus = objHTTP.Status

    On Error Goto 0

    If intStatus = 200 Then
        PingSite = True
    Else
        PingSite = False
    End If

    Set objHTTP = Nothing
End Function

Function artificialIntelligence()
	set wshshell = wscript.CreateObject("wscript.shell")
	set oShell = WScript.CreateObject("WScript.Shell")
	set a = createobject("wscript.shell")
	Set WshShell = CreateObject("WScript.Shell")
	
	Dim msg, greeter
	loopNumberTimesVoice=1
	Set objFileToReadVoiceRecognition = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\voicerecognition.txt",1)
	do while not objFileToReadVoiceRecognition.AtEndOfStream
	     If loopNumberTimesVoice=1 Then
	     	voicerecognitionConfirm = objFileToReadVoiceRecognition.ReadLine()
	     End If
	     loopNumberTimesVoice = loopNumberTimesVoice + 1
	loop
	objFileToReadVoiceRecognition.Close
	
	If voicerecognitionConfirm="" Then
		wscript.sleep(1)
	Else
		wshshell.run "%windir%\Speech\Common\sapisvr.exe -SpeechUX"
	End If


bolActiveConnection = True
		
	If bolActiveConnection = False Then
		wifiError=msgbox("Your computer is not connected to the internet right now." & vbcrlf & "While Sherby will still be able to run, some functionalities may be limited. (For example, Sherby will not be able to access the weather)",16,"No Internet Connection")
	End If

	If masterName=" " Then
		command=inputbox("Hello!" & vbcrlf & "I hope you're having a great day today! What can I help you with?" & vbcrlf & "" & vbcrlf & "Type in 'help' to find out all the commands you can use!", "What do you want me to do?")
	Else
		command=inputbox("Hello " & CStr(masterName) & "!" & vbcrlf & "I hope you're having a great day today! What can I help you with?" & vbcrlf & "" & vbcrlf & "Type in 'help' to find out all the commands you can use!", "What do you want me to do?")
	End If
	Set Sapi = Wscript.CreateObject("SAPI.SpVoice")

	command=LCase(command)
	errorCommand=Trim(command)

	bookmarkUrls=Array()
	Set objFileToReadFor = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\bookmarkUrlStorage.txt",1)
	Dim urlName
	do while not objFileToReadFor.AtEndOfStream
	     urlName = objFileToReadFor.ReadLine()
	     ReDim Preserve bookmarkUrls(UBound(bookmarkUrls) + 1)
	     bookmarkUrls(UBound(bookmarkUrls)) = urlName
	loop
	objFileToReadFor.Close
	bookmarkNames=Array()
	Set objFileToReadName = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\bookmarkNameStorage.txt",1)
	Dim userBookmarkName
	do while not objFileToReadName.AtEndOfStream
	     userBookmarkName = objFileToReadName.ReadLine()
	     ReDim Preserve bookmarkNames(UBound(bookmarkNames) + 1)
	     bookmarkNames(UBound(bookmarkNames)) = userBookmarkName
	loop
	objFileToReadName.Close
	
	If errorCommand="" OR errorCommand=" " OR errorCommand="  " OR errorCommand="   " OR errorCommand="    " OR errorCommand="      " Then
		nocontentError=msgbox("Please do not input only spaces or no text in the input." & vbcrlf & "" & vbcrlf & "To end the script that is running, (a.k.a. Sherby) simply type in nevermind or cancel.",16,"Error: No content")
	ElseIf Len(command) = 1 Then
		Sapi.speak "Please enter a minimum two letter word."
	ElseIf InStr(1, command, " ass ") > 0 OR Right(command, 4) = " ass" OR command="ass" OR InStr(1, command, "shit") > 0 OR InStr(1, command, "bitch") > 0 OR InStr(1, command, "crap") > 0 OR InStr(1, command, "fuck") > 0 OR command = "f u" OR command = "fu" Then
		Sapi.speak "Please watch your language!"
		wscript.sleep(500)
	ElseIf inArray(bookmarkNames, command) >= 0 Then
		location=inArray(bookmarkNames, command)
		urlforcommand=bookmarkUrls(location)
		a.run urlforcommand
	ElseIf InStr(1, command, "www.") > 0 Then
		goodDomain = 1
		If bolActiveConnection = False Then
		wifiErrorForOpeningUrl = MsgBox ("Your computer is not connected to the internet. This functionality will not work without internet connection." & vbCrLf & "" & vbCrLf & "Do you want to continue?", vbYesNo, "No Internet Connection")
		
		Select Case wifiErrorForOpeningUrl
		Case vbYes
			findingWWW=InStr(1, command, "www.")
			url=Mid(command, findingWWW)
			If InStr(1, url, " ") > 0 Then
				finalURL=Left(url, InStr(1, url, " ") - 1)
			Else
				finalUrl = url
			End If
			fullUrl=""
			If InStr(1, command, ".ca") > 0 Then
				fullUrl = Left(url, InStr(1, url, ".ca") + 2)
			ElseIf InStr(1, command, ".com") > 0 Then
				fullUrl = Left(url, InStr(1, url, ".com") + 3)
			ElseIf InStr(1, command, ".org") > 0 Then
				fullUrl = Left(url, InStr(1, url, ".org") + 3)
			ElseIf InStr(1, command, ".int") > 0 Then
				fullUrl = Left(url, InStr(1, url, ".int") + 3)
			ElseIf InStr(1, command, ".gov") > 0 Then
				fullUrl = Left(url, InStr(1, url, ".gov") + 3)
			ElseIf InStr(1, command, ".net") > 0 Then
				fullUrl = Left(url, InStr(1, url, ".net") + 3)
			ElseIf InStr(1, command, ".uk") > 0 Then
				fullUrl = Left(url, InStr(1, url, ".uk") + 2)
			ElseIf InStr(1, command, ".de") > 0 Then
				fullUrl = Left(url, InStr(1, url, ".de") + 2)
			ElseIf InStr(1, command, ".jp") > 0 Then
				fullUrl = Left(url, InStr(1, url, ".jp") + 2)
			ElseIf InStr(1, command, ".fr") > 0 Then
				fullUrl = Left(url, InStr(1, url, ".fr") + 2)
			ElseIf InStr(1, command, ".au") > 0 Then
				fullUrl = Left(url, InStr(1, url, ".au") + 2)
			ElseIf InStr(1, command, ".us") > 0 Then
				fullUrl = Left(url, InStr(1, url, ".us") + 2)
			ElseIf InStr(1, command, ".ru") > 0 Then
				fullUrl = Left(url, InStr(1, url, ".ru") + 2)
			ElseIf InStr(1, command, ".ch") > 0 Then
				fullUrl = Left(url, InStr(1, url, ".ch") + 2)
			ElseIf InStr(1, command, ".it") > 0 Then
				fullUrl = Left(url, InStr(1, url, ".it") + 2)
			ElseIf InStr(1, command, ".io") > 0 Then
				fullUrl = Left(url, InStr(1, url, ".io") + 2)
			Else
				Sapi.speak "Unfortunately, the domain entered wasn't recognized, so we can't open your desired website"
				goodDomain = 0
			End If

			urlRefined=Mid(fullUrl, 5)
			Sapi.speak "Going too, " + urlRefined
			a.run finalUrl
		Case vbNo
			wscript.sleep(1)
		End Select
		Else
			findingWWW=InStr(1, command, "www.")
			url=Mid(command, findingWWW)
			If InStr(1, url, " ") > 0 Then
				finalURL=Left(url, InStr(1, url, " ") - 1)
			Else
				finalUrl = url
			End If
			fullUrl=""
			If InStr(1, command, ".ca") > 0 Then
				fullUrl = Left(url, InStr(1, url, ".ca") + 2)
			ElseIf InStr(1, command, ".com") > 0 Then
				fullUrl = Left(url, InStr(1, url, ".com") + 3)
			ElseIf InStr(1, command, ".org") > 0 Then
				fullUrl = Left(url, InStr(1, url, ".org") + 3)
			ElseIf InStr(1, command, ".int") > 0 Then
				fullUrl = Left(url, InStr(1, url, ".int") + 3)
			ElseIf InStr(1, command, ".gov") > 0 Then
				fullUrl = Left(url, InStr(1, url, ".gov") + 3)
			ElseIf InStr(1, command, ".net") > 0 Then
				fullUrl = Left(url, InStr(1, url, ".net") + 3)
			ElseIf InStr(1, command, ".uk") > 0 Then
				fullUrl = Left(url, InStr(1, url, ".uk") + 2)
			ElseIf InStr(1, command, ".de") > 0 Then
				fullUrl = Left(url, InStr(1, url, ".de") + 2)
			ElseIf InStr(1, command, ".jp") > 0 Then
				fullUrl = Left(url, InStr(1, url, ".jp") + 2)
			ElseIf InStr(1, command, ".fr") > 0 Then
				fullUrl = Left(url, InStr(1, url, ".fr") + 2)
			ElseIf InStr(1, command, ".au") > 0 Then
				fullUrl = Left(url, InStr(1, url, ".au") + 2)
			ElseIf InStr(1, command, ".us") > 0 Then
				fullUrl = Left(url, InStr(1, url, ".us") + 2)
			ElseIf InStr(1, command, ".ru") > 0 Then
				fullUrl = Left(url, InStr(1, url, ".ru") + 2)
			ElseIf InStr(1, command, ".ch") > 0 Then
				fullUrl = Left(url, InStr(1, url, ".ch") + 2)
			ElseIf InStr(1, command, ".it") > 0 Then
				fullUrl = Left(url, InStr(1, url, ".it") + 2)
			ElseIf InStr(1, command, ".io") > 0 Then
				fullUrl = Left(url, InStr(1, url, ".io") + 2)
			Else
				Sapi.speak "Unfortunately, the domain entered wasn't recognized, so we can't open your desired website"
				goodDomain = 0
			End If

			urlRefined=Mid(fullUrl, 5)
			Sapi.speak "Going too, " + urlRefined
			a.run finalUrl
		End If
	ElseIf command="create bookmark" OR command="add bookmark" OR InStr(1, command, "set") > 0 And InStr(1, command, "bookmark") > 0 Or InStr(1, command, "new") > 0 And InStr(1, command, "bookmark") > 0 OR InStr(1, command, "create") > 0 And InStr(1, command, "bookmark") > 0 Or InStr(1, command, "add") > 0 And InStr(1, command, "bookmark") > 0 Or InStr(1, command, "make") > 0 And InStr(1, command, "bookmark") > 0 Then
		Set objFSO=CreateObject("Scripting.FileSystemObject")
		anotherBookmark=0
		Sapi.speak "what bookmark would you like to create?"
		bookmarkName=CStr(inputbox("What is the bookmark's name?" & vbcrlf & "" & vbcrlf & "WARNING: A bookmark name will override a specific preset command. For example, if you create a bookmark called 'date and time', it will override the command 'date and time'. However, you will still be able to access the preset command by typing something similar, such as 'what is the day and time today?'" & vbcrlf & "NOTE: Capitals will not matter in the name of the bookmark. Ex. 'date' will be considered the same thing as 'DaTE'", "Specify bookmark name"))
		bookmarkName=LCase(bookmarkName)
		bookmarkNameRefined=Trim(bookmarkName)
		If bookmarkNameRefined="" Then
			nocontentError=msgbox("Please do not create an empty bookmark",16,"Error: Bookmark has no content")
		Else
			Sapi.speak "Now, please specify the url or file location of your bookmark"
			bookmarkUrl=inputbox("Please input the URL or file location that you want the bookmark '" + bokmarkName + "' to lead to in the input box below." & vbcrlf & "" & vbcrlf & "NOTE: The url has to be in this format: http://google.com. Another example could be http://apple.net. If you want your bookmark to run a file, you must specify the full file path. (ex. homework.txt will not work)", "Specify URL or file location for bookmark '" + bookmarkName + "'", "http://")
			bookmarkUrl=LCase(bookmarkUrl)
			bookmarkUrlRefined=Trim(bookmarkUrl)
			If bookmarkUrlRefined="" Then
				nocontentError=msgbox("Please do not create an empty bookmark",16,"Error: Bookmark has no content")
			ElseIf Left(Right(bookmarkUrl, 4), 1) <> "." And Left(bookmarkUrl, 7) <> "http://" And objFSO.FileExists(bookmarkUrl) OR Left(Right(bookmarkUrl, 3), 1) <> "." And Left(bookmarkUrl, 7) <> "http://" And objFSO.FileExists(bookmarkUrl) Then
				properFile = msgbox("Unfortunately, we couldn't find the file '" + CStr(bookmarkUrl) + "'. Make sure to input the full file location. For example, inputting only 'homework.txt' will not suffice; you must input the entire file location...", 16,"Error: File not found")
			Else
			Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\bookmarkNameStorage.txt",1)
			strFileText = objFileToRead.ReadAll()
			strFileText = strFileText + bookmarkName
			Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\bookmarkNameStorage.txt",2,true)
			objFileToWrite.WriteLine(strFileText)
			objFileToWrite.Close
			objFileToRead.Close
			Set objFileToReadThree = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\bookmarkUrlStorage.txt",1)
			strFileTextTwo = objFileToReadThree.ReadAll()
			strFileTextTwo = strFileTextTwo + bookmarkURL
			Set objFileToWriteTwo = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\bookmarkUrlStorage.txt",2,true)
			objFileToWriteTwo.WriteLine(strFileTextTwo)
			objFileToWriteTwo.Close
			objFileToReadThree.Close
			bookmarkCreated="the bookmark," + CStr(bookmarkName) + ",has been created"
			Sapi.speak bookmarkCreated
			wscript.sleep(800)
			End If
			End If
			do while anotherBookmark<1
			Sapi.speak "Do you want to create another bookmark?"
			result = MsgBox ("Do you want to create another bookmark?", vbYesNo, "Another one? (bites the dust)")
			
			Select Case result
			Case vbYes
				Sapi.speak "what bookmark would you like to create?"
				bookmarkName=inputbox("What is the bookmark's name?" & vbcrlf & "" & vbcrlf & "WARNING: A bookmark name will override a specific preset command. For example, if you create a bookmark called 'date and time', it will override the command 'date and time'. However, you will still be able to access the preset command by typing something similar, such as 'what is the day and time today?'" & vbcrlf & "NOTE: Capitals will not matter in the name of the bookmark. Ex. 'date' will be considered the same thing as 'DaTE'", "Create Bookmark: Step 1")
				bookmarkName=LCase(bookmarkName)
				bookmarkNameRefined=Trim(bookmarkName)
				If bookmarkNameRefined="" Then
					nocontentError=msgbox("Please do not create an empty bookmark or a bookmark only with spaces",16,"Error: Bookmark has no content")
				ElseIf Left(Right(bookmarkUrl, 4), 1) <> "." And Left(bookmarkUrl, 7) <> "http://" And objFSO.FileExists(bookmarkUrl) OR Left(Right(bookmarkUrl, 3), 1) <> "." And Left(bookmarkUrl, 7) <> "http://" And objFSO.FileExists(bookmarkUrl) Then
					properFile = msgbox("Unfortunately, we couldn't find the file '" + CStr(bookmarkUrl) + "'. Make sure to input the full file location. For example, inputting only 'homework.txt' will not suffice; you must input the entire file location...", 16,"Error: File not found")
				Else
					Sapi.speak "Now, please specify the url or file location of your bookmark"
					bookmarkUrl=inputbox("Please input the URL or file location that you want the bookmark '" + bokmarkName + "' to lead to in the input box below." & vbcrlf & "" & vbcrlf & "NOTE: The url has to be in this format: http://google.com. Another example could be http://apple.net. If you want your bookmark to run a file, you must specify the full file path. (ex. homework.txt will not work)", "Specify URL or file location for bookmark '" + bookmarkName + "'", "http://")
					bookmarkUrl=LCase(bookmarkUrl)
					bookmarkUrlRefined=Trim(bookmarkUrl)
					If bookmarkUrlRefined="" Then
						nocontentError=msgbox("Please do not create an empty bookmark",16,"Error: Bookmark has no content")
					Else
					Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\bookmarkNameStorage.txt",1)
					strFileText = objFileToRead.ReadAll()
					strFileText = strFileText + bookmarkName
					Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\bookmarkNameStorage.txt",2,true)
					objFileToWrite.WriteLine(strFileText)
					objFileToWrite.Close
					objFileToRead.Close
					Set objFileToReadThree = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\bookmarkUrlStorage.txt",1)
					strFileTextTwo = objFileToReadThree.ReadAll()
					strFileTextTwo = strFileTextTwo + bookmarkURL
					Set objFileToWriteTwo = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\bookmarkUrlStorage.txt",2,true)
					objFileToWriteTwo.WriteLine(strFileTextTwo)
					objFileToWriteTwo.Close
					objFileToReadThree.Close
					bookmarkCreated="the bookmark," + CStr(bookmarkName) + ",has been created"
					Sapi.speak bookmarkCreated
					wscript.sleep(800)
				End If
				End If
			Case vbNo
				anotherBookmark = 1
			End Select
		loop
	ElseIf InStr(1, command, "remove") > 0 And InStr(1, command, "bookmark") > 0 And InStr(1, command, "all") > 0 OR InStr(1, command, "delete") > 0 And InStr(1, command, "bookmark") > 0 And InStr(1, command, "all") > 0 OR InStr(1, command, "discard") > 0 And InStr(1, command, "bookmark") > 0 And InStr(1, command, "all") > 0 OR InStr(1, command, "erase") > 0 And InStr(1, command, "bookmark") > 0 And InStr(1, command, "all") > 0 Then
		Sapi.speak "Are you sure you want to delete all your bookmarks?"
		areSureDeleteAllBookmarks = MsgBox ("Are you sure you want to delete all of your bookmarks?", vbYesNo, "Delete all bookmarks?")
		
		Select Case areSureDeleteAllBookmarks
		Case vbYes
			Set objFileToWriteDeleteAll = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\bookmarkNameStorage.txt",2,true)
			objFileToWriteDeleteAll.WriteLine("")
			objFileToWriteDeleteAll.Close
			Set objFileToWriteDeleteAllURL = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\bookmarkUrlStorage.txt",2,true)
			objFileToWriteDeleteAllURL.WriteLine("")
			objFileToWriteDeleteAllURL.Close
			Sapi.speak "All of your bookmarks have been deleted"
		Case vbNo
			Sapi.speak "Deleting all bookmarks has been cancelled"
		End Select
	ElseIf InStr(1, command, "remove") > 0 And InStr(1, command, "bookmark") > 0 OR InStr(1, command, "delete") > 0 And InStr(1, command, "bookmark") > 0 OR InStr(1, command, "discard") > 0 And InStr(1, command, "bookmark") > 0 OR InStr(1, command, "erase") > 0 And InStr(1, command, "bookmark") > 0 Then
		Dim strNewContents: strNewContents=""
		Dim strNewContentsTwo: strNewContentsTwo=""
		deleteAnother=0
		Sapi.speak "What bookmark would you like to delete?"
		bookmarktodelete=inputbox("What bookmark would you like to delete? Enter the bookmark's NAME to delete it." & vbCrLf & "" & vbcrlf & "ex. If you wanted to delete the bookmark 'google', you simply have to type in the NAME of the bookmark, (in this case, it's simply 'google') to delete it.","Delete Bookmark")
		bookmarktodelete=LCase(bookmarktodelete)
		If Trim(bookmarktodelete)="" Then
			bookmarkErrorDelete=msgbox("Please do not leave the input box empty or only type in spaces.", 16, "Error: No content")
		Else
		bookmarkUrls=Array()
		Set objFileToReadFor = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\bookmarkUrlStorage.txt",1)
		do while not objFileToReadFor.AtEndOfStream
			urlName = objFileToReadFor.ReadLine()
			ReDim Preserve bookmarkUrls(UBound(bookmarkUrls) + 1)
			bookmarkUrls(UBound(bookmarkUrls)) = urlName
		loop
		objFileToReadFor.Close
		bookmarkNames=Array()
		Set objFileToReadName = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\bookmarkNameStorage.txt",1)
		do while not objFileToReadName.AtEndOfStream
			userBookmarkName = objFileToReadName.ReadLine()
			ReDim Preserve bookmarkNames(UBound(bookmarkNames) + 1)
			bookmarkNames(UBound(bookmarkNames)) = userBookmarkName
		loop
		objFileToReadName.Close
		If inArray(bookmarkNames, bookmarktodelete) >= 0 Then
			location=inArray(bookmarkNames, command)
			locationRefined=location+1
			urlforcommand=bookmarkUrls(locationRefined)
			
			Set objFileToReadBookmarkDel = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\bookmarkUrlStorage.txt",1)
			do while not objFileToReadBookmarkDel.AtEndOfStream
				urlName = objFileToReadBookmarkDel.ReadLine()
				urlName=LCase(urlName)
				If urlforcommand=urlName Then
					wscript.sleep(1)
				Else
					strNewContents = strNewContents + urlName + vbCrLf
				End If
			loop
			objFileToReadBookmarkDel.Close
			Set objFileDeleteName = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\bookmarkUrlStorage.txt", 2, true)
			objFileDeleteName.Write strNewContents
			objFileDeleteName.Close
			
			Set objFileToReadBookmarkDelName = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\bookmarkNameStorage.txt",1)
			do while not objFileToReadBookmarkDelName.AtEndOfStream
					userBookmarkName=objFileToReadBookmarkDelName.ReadLine()
				userBookmarkName=LCase(userBookmarkName)
				If userBookmarkName=bookmarktodelete Then
					wscript.sleep(1)
				Else
					strNewContentsTwo = strNewContentsTwo + userBookmarkName + vbcrlf
				End If
			loop
			objFileToReadBookmarkDelName.Close
			Set objFileDeleteName = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\bookmarkNameStorage.txt", 2, true)
			objFileDeleteName.Write strNewContentsTwo
			objFileDeleteName.Close
			
			hasbeendeletedMessage="the bookmark, " + bookmarktodelete + ", has been deleted successfully"
			Sapi.speak hasbeendeletedMessage
	
		Else
			bookmarkErrorMessageSapi="Unfortunately, the bookmark, " + bookmarktodelete + ", doesn't exist"
			Sapi.speak bookmarkErrorMessageSapi
			bookmarkErrorMessage="Bookmark error" & vbcrlf & vbcrlf + "Unfortunately, we couldn't find the bookmark '" + bookmarktodelete + "'..."
			bookmarkError=msgbox(bookmarkErrorMessage,16,"Could not find bookmark")
		End If
		End If
		wscript.sleep(700)
		do while deleteAnother<1
			Sapi.speak "Do you want to delete another bookmark?"
			result = MsgBox ("Do you want to delete another bookmark?", vbYesNo, "Delete another one?")
			
			Select Case result
			Case vbYes
				strNewContents=""
				strNewContentsTwo=""
				deleteAnother=0
				Sapi.speak "What bookmark would you like to delete?"
				bookmarktodelete=inputbox("What bookmark would you like to delete? Enter the bookmark's NAME to delete it." & vbCrLf & "" & vbcrlf & "ex. If you wanted to delete the bookmark 'google', you simply have to type in the NAME of the bookmark, (in this case, it's simply 'google') to delete it.","Delete Bookmark")
				bookmarktodelete=LCase(bookmarktodelete)
				
				If Trim(bookmarktodelete)="" Then
					bookmarkErrorDelete=msgbox("Please do not leave the input box empty or only type in spaces.", 16, "Error: No content")
				Else
				bookmarkUrls=Array()
				Set objFileToReadFor = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\bookmarkUrlStorage.txt",1)
				do while not objFileToReadFor.AtEndOfStream
					urlName = objFileToReadFor.ReadLine()
					ReDim Preserve bookmarkUrls(UBound(bookmarkUrls) + 1)
					bookmarkUrls(UBound(bookmarkUrls)) = urlName
				loop
				objFileToReadFor.Close
				bookmarkNames=Array()
				Set objFileToReadName = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\bookmarkNameStorage.txt",1)
				do while not objFileToReadName.AtEndOfStream
					userBookmarkName = objFileToReadName.ReadLine()
					ReDim Preserve bookmarkNames(UBound(bookmarkNames) + 1)
					bookmarkNames(UBound(bookmarkNames)) = userBookmarkName
				loop
				objFileToReadName.Close
				If inArray(bookmarkNames, bookmarktodelete) >= 0 Then
					location=inArray(bookmarkNames, command)
					locationRefined=location+1
					urlforcommand=bookmarkUrls(locationRefined)
					
					Set objFileToReadBookmarkDel = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\bookmarkUrlStorage.txt",1)
					do while not objFileToReadBookmarkDel.AtEndOfStream
						urlName = objFileToReadBookmarkDel.ReadLine()
						urlName=LCase(urlName)
						If urlforcommand=urlName Then
							wscript.sleep(1)
						Else
							strNewContents = strNewContents + urlName + vbCrLf
						End If
					loop
					objFileToReadBookmarkDel.Close
					Set objFileDeleteName = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\bookmarkUrlStorage.txt", 2, true)
					objFileDeleteName.Write strNewContents
					objFileDeleteName.Close
					
					Set objFileToReadBookmarkDelName = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\bookmarkNameStorage.txt",1)
					do while not objFileToReadBookmarkDelName.AtEndOfStream
							userBookmarkName=objFileToReadBookmarkDelName.ReadLine()
						userBookmarkName=LCase(userBookmarkName)
						If userBookmarkName=bookmarktodelete Then
							wscript.sleep(1)
						Else
							strNewContentsTwo = strNewContentsTwo + userBookmarkName + vbcrlf
						End If
					loop
					objFileToReadBookmarkDelName.Close
					Set objFileDeleteName = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\bookmarkNameStorage.txt", 2, true)
					objFileDeleteName.Write strNewContentsTwo
					objFileDeleteName.Close
					
					hasbeendeletedMessage="the bookmark, " + bookmarktodelete + ", has been deleted successfully"
					Sapi.speak hasbeendeletedMessage
			
				Else
					bookmarkErrorMessageSapi="Unfortunately, the bookmark, " + bookmarktodelete + ", doesn't exist"
					Sapi.speak bookmarkErrorMessageSapi
					bookmarkErrorMessage="Bookmark error" & vbcrlf & vbcrlf + "Unfortunately, we couldn't find the bookmark '" + bookmarktodelete + "'..."
					bookmarkError=msgbox(bookmarkErrorMessage,16,"Could not find bookmark")
				End If
				End If
				wscript.sleep(700)
			Case vbNo
				deleteAnother = 1
			End Select
		loop
	ElseIf command="list of bookmarks" OR command="list of all bookmarks" OR command="bookmarks list" OR command="bookmark" OR command="bookmarks" OR InStr(1, command, "list") > 0 And InStr(1, command, "bookmark") > 0 OR command="bookmarks list" OR InStr(1, command, "all") > 0 And InStr(1, command, "bookmark") > 0 OR InStr(1, command, "show") > 0 And InStr(1, command, "bookmark") > 0 OR InStr(1, command, "find") > 0 And InStr(1, command, "bookmark") > 0 OR InStr(1, command, "what") > 0 And InStr(1, command, "bookmark") > 0 OR InStr(1, command, "how") > 0 And InStr(1, command, "bookmark") > 0 Then
		Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\bookmarkNameStorage.txt",1)
		Dim strLine
		Dim bookmarksList
		Dim loopNumber
		Dim bookmark
		loopNumber=0
		bookmarkUrl=Array()
		bookmarksList="Here is a list of all bookmarks:" & vbcrlf & "" & vbcrlf & ""
		Set objFileToReadTwo = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\bookmarkUrlStorage.txt",1)
		Dim strLineTwo
		do while not objFileToReadTwo.AtEndOfStream
		     strLineTwo = objFileToReadTwo.ReadLine()
		     ReDim Preserve bookmarkUrl(UBound(bookmarkUrl) + 1)
		     bookmarkUrl(UBound(bookmarkUrl)) = strLineTwo
		loop
		objFileToReadTwo.Close
		do while not objFileToRead.AtEndOfStream
		     If loopNumber > 0 Then
				strLine = objFileToRead.ReadLine()
				If strLine="" Then
					wscript.sleep(1)
				Else
					bookmarksList=bookmarksList + LCase(strLine) + ", " + "(" + LCase(bookmarkUrl(loopNumber-1)) + ")" & vbcrlf
				End If
		     End If
		     loopNumber=loopNumber + 1
		loop
		
		If strLine="" Then
		     Sapi.speak "You currently have no bookmarks"
		Else
		     bookmarkURLElementsNum = uBound(bookmarkURL)
		     If bookmarkURLElementsNum=1 Then
		     sapi_numberofbookmarks="you currently have " + CStr(bookmarkURLElementsNum) + " bookmark"
		     Else
		     sapi_numberofbookmarks="you currently have " + CStr(bookmarkURLElementsNum) + " bookmarks"
		     End If
		     Sapi.speak sapi_numberofbookmarks
		     thelistofallthebookmarks=msgbox(bookmarksList,,"All bookmarks")
		End If
		objFileToRead.Close
	ElseIf InStr(1, command, "new") > 0 And InStr(1, command, "goal") > 0 OR InStr(1, command, "set") > 0 And InStr(1, command, "bookmark") > 0 OR InStr(1, command, "new") > 0 And InStr(1, command, "todo") > 0 OR InStr(1, command, "new") > 0 And InStr(1, command, "todo") > 0 OR InStr(1, command, "create") > 0 And InStr(1, command, "goal") > 0 Or InStr(1, command, "add") > 0 And InStr(1, command, "goal") > 0 Or InStr(1, command, "make") > 0 And InStr(1, command, "goal") > 0 Or InStr(1, command, "create") > 0 And InStr(1, command, "to do") > 0 Or InStr(1, command, "add") > 0 And InStr(1, command, "to do") > 0 Or InStr(1, command, "make") > 0 And InStr(1, command, "to do") > 0 Or InStr(1, command, "create") > 0 And InStr(1, command, "todo") > 0 Or InStr(1, command, "add") > 0 And InStr(1, command, "todo") > 0 Or InStr(1, command, "make") > 0 And InStr(1, command, "todo") > 0 Then
			Sapi.speak "Please enter what your goal is"
			userGoal=inputbox("Enter your goal/what you want to do below.", "What is your goal?")
			If Trim(userGoal)="" Then
				goalError=msgbox("Please do not leave the input box empty or only type in spaces.", 16, "Error: No content")
			Else
				Set objFileToReadForGoals = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\toDoGoals.txt",1)
				FileTextForGoals = objFileToReadForGoals.ReadAll()
				FileTextForGoals = FileTextForGoals + userGoal
				Set objFileToWriteForGoals = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\toDoGoals.txt",2,true)
				objFileToWriteForGoals.WriteLine(FileTextForGoals)
				objFileToWriteForGoals.Close
				objFileToReadForGoals.Close
				Sapi.speak "Your goal has been successfully saved"
			End If
			anotherGoal=0
			do while anotherGoal<1
				Sapi.speak "Do you want to create another goal?"
				result = MsgBox ("Do you want to create another goal?", vbYesNo, "Create another one?")
				
				Select Case result
				Case vbYes
					Sapi.speak "Please enter what your goal is"
					userGoal=inputbox("Enter your goal/what you want to do below.", "What is your goal?")
					If Trim(userGoal)="" Then
						goalError=msgbox("Please do not leave the input box empty or only type in spaces.", 16, "Error: No content")
					Else						
						Set objFileToReadForGoals = CreateObject("Scripting.FileSystemObject").OpenTextFile("toDoGoals.txt",1)
						FileTextForGoals = objFileToReadForGoals.ReadAll()
						FileTextForGoals = FileTextForGoals + userGoal
						Set objFileToWriteForGoals = CreateObject("Scripting.FileSystemObject").OpenTextFile("toDoGoals.txt",2,true)
						objFileToWriteForGoals.WriteLine(FileTextForGoals)
						objFileToWriteForGoals.Close
						objFileToReadForGoals.Close
						Sapi.speak "Your goal has been successfully saved"
					End If
				Case vbNo
					anotherGoal = 1
				End Select
		loop
	ElseIf InStr(1, command, "remove") > 0 And InStr(1, command, "goal") > 0 And InStr(1, command, "all") > 0 OR InStr(1, command, "delete") > 0 And InStr(1, command, "goal") > 0 And InStr(1, command, "all") > 0 OR InStr(1, command, "discard") > 0 And InStr(1, command, "goal") > 0 And InStr(1, command, "all") > 0 OR InStr(1, command, "erase") > 0 And InStr(1, command, "goal") > 0 And InStr(1, command, "all") > 0 OR InStr(1, command, "remove") > 0 And InStr(1, command, "todo") > 0 And InStr(1, command, "all") > 0 OR InStr(1, command, "delete") > 0 And InStr(1, command, "todo") > 0 And InStr(1, command, "all") > 0 OR InStr(1, command, "discard") > 0 And InStr(1, command, "todo") > 0 And InStr(1, command, "all") > 0 OR InStr(1, command, "erase") > 0 And InStr(1, command, "todo") > 0 And InStr(1, command, "all") > 0 OR InStr(1, command, "remove") > 0 And InStr(1, command, "to do") > 0 And InStr(1, command, "all") > 0 OR InStr(1, command, "delete") > 0 And InStr(1, command, "to do") > 0 And InStr(1, command, "all") > 0 OR InStr(1, command, "discard") > 0 And InStr(1, command, "to do") > 0 And InStr(1, command, "all") > 0 OR InStr(1, command, "erase") > 0 And InStr(1, command, "to do") > 0 And InStr(1, command, "all") > 0 Then
		Sapi.speak "Are you sure you want to delete all your goals?"
		areSureDeleteAllGoals = MsgBox ("Are you sure you want to delete all of your goals", vbYesNo, "Delete all goals?")
		
		Select Case areSureDeleteAllGoals
		Case vbYes
			Set objFileToWriteDeleteAllGoals = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\toDoGoals.txt",2,true)
			objFileToWriteDeleteAllGoals.WriteLine("")
			objFileToWriteDeleteAllGoals.Close
			Sapi.speak "All of your goals have been deleted"
		Case vbNo
			Sapi.speak "Deleting all goals has been cancelled"
		End Select
	ElseIf InStr(1, command, "delete") > 0 And InStr(1, command, "goal") > 0 OR InStr(1, command, "discard") > 0 And InStr(1, command, "goal") > 0 OR InStr(1, command, "erase") > 0 And InStr(1, command, "goal") > 0 OR InStr(1, command, "discard") > 0 And InStr(1, command, "todo") > 0 OR InStr(1, command, "erase") > 0 And InStr(1, command, "todo") > 0 OR InStr(1, command, "discard") > 0 And InStr(1, command, "to do") > 0 OR InStr(1, command, "erase") > 0 And InStr(1, command, "to do") > 0 Or InStr(1, command, "remove") > 0 And InStr(1, command, "goal") > 0 Or InStr(1, command, "done") > 0 And InStr(1, command, "goal") > 0 Or InStr(1, command, "finish") > 0 And InStr(1, command, "goal") > 0 OR InStr(1, command, "delete") > 0 And InStr(1, command, "todo") > 0 Or InStr(1, command, "remove") > 0 And InStr(1, command, "todo") > 0 Or InStr(1, command, "done") > 0 And InStr(1, command, "todo") > 0 Or InStr(1, command, "finish") > 0 And InStr(1, command, "todo") > 0 OR InStr(1, command, "delete") > 0 And InStr(1, command, "to do") > 0 Or InStr(1, command, "remove") > 0 And InStr(1, command, "to do") > 0 Or InStr(1, command, "to do") > 0 And InStr(1, command, "goal") > 0 Or InStr(1, command, "finish") > 0 And InStr(1, command, "to do") > 0 Then
		Sapi.speak "What goal would you like to delete or mark as done?"
		goalToDelete=inputbox("Please specify the goal you would like to delete or mark as done." & vbcrlf & "" & vbcrlf & "IMPORTANT: The input must be an integer indicating the 'index of the goal'. In other words, (after viewing all your current goals) if you want to delete or mark your 3rd goal as done, you simply need to type in the number '3'.", "Deleting/marking a goal as done")
		If Trim(goalToDelete)="" Then
			goalError=msgbox("Please do not leave the input box empty or only type in spaces.", 16, "Error: No content")
		Else			
			foundGoal=1
			lineNum=0
			goalItself=""
			
				Set objFileToReadGoalDel = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\toDoGoals.txt",1)
				do while not objFileToReadGoalDel.AtEndOfStream
					goalToDeleteName=objFileToReadGoalDel.ReadLine()
					If lineNum=CInt(goalToDelete) Then
						wscript.sleep(1)
						foundGoal=0
						goalItself=goalToDeleteName
					Else
						strNewContentsForGoal = strNewContentsForGoal + goalToDeleteName + vbcrlf
					End If
					lineNum=lineNum + 1
				loop
				objFileToReadGoalDel.Close
				Set objFileDeleteGoal = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\toDoGoals.txt", 2, true)
				objFileDeleteGoal.Write strNewContentsForGoal
				objFileDeleteGoal.Close
				
				If foundGoal=0 Then		
					hasbeendeletedMessage="Your specified goal has been deleted."
					Sapi.speak hasbeendeletedMessage
					hasbeendeletedMessageAlert="The goal '" + goalItself + "' has been succesfully deleted."
					deletedConfirmForGoal=msgbox(hasbeendeletedMessageAlert,,"Goal has been successfully deleted")
				Else
					bookmarkErrorMessageSapi="Unfortunately, your specified goal could not be found..."
					Sapi.speak bookmarkErrorMessageSapi
					bookmarkErrorMessage="Input error" & vbcrlf & vbcrlf + "Unfortunately, we couldn't find your specified goal"
					bookmarkError=msgbox(bookmarkErrorMessage,16,"Could not find specified goal")
				End If
				wscript.sleep(200)
			End If
		deleteAnotherGoal=0
		do while deleteAnotherGoal<1
			Sapi.speak "Do you want to delete or mark another goal as done?"
			result = MsgBox ("Do you want to delete/mark another goal as done?", vbYesNo, "Deleting/marking another one?")
			
			Select Case result
			Case vbYes
				Sapi.speak "What goal would you like to delete or mark as done?"
				goalToDelete=inputbox("Please specify the goal you would like to delete or mark as done." & vbcrlf & "" & vbcrlf & "IMPORTANT: The input must be an integer indicating the 'index of the goal'. In other words, (after viewing all your current goals) if you want to delete or mark your 3rd goal as done, you simply need to type in the number '3'.", "Deleting/marking a goal as done")
				If Trim(goalToDelete)="" Then
					goalError=msgbox("Please do not leave the input box empty or only type in spaces.", 16, "Error: No content")
				ElseIf isNumeric(goalToDelete) = False Then
					goalRemoveError = msgbox("Please input a number.", 16, "Error: Input must be a number")
				ElseIf InStr(1, goalToDelete, ".") > 0 Then
					goalRemoveError = msgbox("Please input an integer, not a decimal number.", 16, "Error: Number must be integer")
				Else			
					foundGoal=0
					lineNum=0
					goalItself=""
					
						Set objFileToReadGoalDel = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\toDoGoals.txt",1)
						do while not objFileToReadGoalDel.AtEndOfStream
							goalToDeleteName=objFileToReadGoalDel.ReadLine()
							If lineNum=goalToDelete Then
								wscript.sleep(1)
								foundGoal=0
								goalItself=goalToDeleteName
							Else
								strNewContentsForGoal = strNewContentsForGoal + goalToDeleteName + vbcrlf
							End If
							lineNum=lineNum + 1
						loop
						objFileToReadGoalDel.Close
						Set objFileDeleteGoal = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\toDoGoals.txt", 2, true)
						objFileDeleteGoal.Write strNewContentsForGoal
						objFileDeleteGoal.Close
						
						If foundGoal=0 Then		
							hasbeendeletedMessage="Your specified goal has been deleted."
							Sapi.speak hasbeendeletedMessage
							hasbeendeletedMessageAlert="The goal '" + goalItself + "' has been succesfully deleted."
							deletedConfirmForGoal=msgbox(hasbeendeletedMessageAlert,,"Goal has been successfully deleted")
						Else
							bookmarkErrorMessageSapi="Unfortunately, your specified goal could not be found..."
							Sapi.speak bookmarkErrorMessageSapi
							bookmarkErrorMessage="Input error" & vbcrlf & vbcrlf + "Unfortunately, we couldn't find your specified goal"
							bookmarkError=msgbox(bookmarkErrorMessage,16,"Could not find specified goal")
						End If
						wscript.sleep(200)
				End If
			Case vbNo
				deleteAnotherGoal = 1
			End Select
		loop
	ElseIf command="goal" OR command="todo" Or command="to do" OR command="goals" OR command="todos" Or command="to dos" OR command="goal's" OR command="todo's" Or command="to do's" OR InStr(1, command, "goal") > 0 And InStr(1, command, "all") > 0 OR InStr(1, command, "to do") > 0 And InStr(1, command, "all") > 0 Or InStr(1, command, "todo") > 0 And InStr(1, command, "all") > 0 OR InStr(1, command, "goal") > 0 And InStr(1, command, "list") > 0 OR InStr(1, command, "to do") > 0 And InStr(1, command, "list") > 0 Or InStr(1, command, "todo") > 0 And InStr(1, command, "list") > 0 OR InStr(1, command, "show") > 0 And InStr(1, command, "goal") > 0 OR InStr(1, command, "find") > 0 And InStr(1, command, "goal") > 0 OR InStr(1, command, "what") > 0 And InStr(1, command, "goal") > 0 OR InStr(1, command, "how") > 0 And InStr(1, command, "goal") > 0 Then
		loopNumberTwo=0
		toDoArray=Array()
		toDoDisplayList="Here is a list of all your goals:" & vbcrlf & "" & vbcrlf & ""
		Set objFileToReadForDisplayingGoals = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\toDoGoals.txt",1)
		Dim strLineToDo
		do while not objFileToReadForDisplayingGoals.AtEndOfStream
		     strLineToDo = objFileToReadForDisplayingGoals.ReadLine()
		     ReDim Preserve toDoArray(UBound(toDoArray) + 1)
		     toDoArray(UBound(toDoArray)) = strLineToDo
		loop
		objFileToReadForDisplayingGoals.Close
		
		Set objFileToReadForDisplayingGoalsTwo = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\toDoGoals.txt",1)
		do while not objFileToReadForDisplayingGoalsTwo.AtEndOfStream
		     If loopNumberTwo > 0 Then
				strLineForToDo = objFileToReadForDisplayingGoalsTwo.ReadLine()
				If strLineForToDo="" Then
					wscript.sleep(1)
				Else
					ToDoListFull=ToDoListFull + "- " + strLineForToDo & vbcrlf
				End If
		     End If
		     loopNumberTwo=loopNumberTwo + 1
		loop
		
		If strLineForToDo="" Then
		     Sapi.speak "You currently have no goals"
		Else
		     goalsElementsNum = uBound(toDoArray)
		     If goalsElementsNum=1 Then
		     sapi_numberofgoals="you currently have " + CStr(goalsElementsNum) + " goal"
		     Else
		     sapi_numberofgoals="you currently have " + CStr(goalsElementsNum) + " goals"
		     End If
		     Sapi.speak sapi_numberofgoals
		     thelistofallthebookmarks=msgbox(ToDoListFull,,"All goals")
		End If
		objFileToReadForDisplayingGoalsTwo.Close
	ElseIf command="create note" OR command="add note" OR InStr(1, command, "set") > 0 And InStr(1, command, "bookmark") > 0 Or InStr(1, command, "new") > 0 And InStr(1, command, "note") > 0 OR InStr(1, command, "create") > 0 And InStr(1, command, "note") > 0 Or InStr(1, command, "add") > 0 And InStr(1, command, "note") > 0 Or InStr(1, command, "make") > 0 And InStr(1, command, "note") > 0 Then
			Sapi.speak "Please specify the content of your note..."
			userNote=inputbox("Enter your note below." & vbcrlf & "" & vbcrlf & "(ex. The scientific name of a lizard is 'Lacertilia')", "What is the content of your note?")
			If Trim(userNote)="" Then
				noteError=msgbox("Please do not leave the input box empty or only type in spaces.", 16, "Error: No content")
			Else
				Set objFileToReadForNotes = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\notesStorage.txt",1)
				FileTextForNotes = objFileToReadForNotes.ReadAll()
				FileTextForNotes = FileTextForNotes + userNote
				Set objFileToWriteForNotes = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\notesStorage.txt",2,true)
				objFileToWriteForNotes.WriteLine(FileTextForNotes)
				objFileToWriteForNotes.Close
				objFileToReadForNotes.Close
				Sapi.speak "Your note has been successfully saved"
			End If
			anotherNote=0
			do while anotherNote<1
				Sapi.speak "Would you like to create another note?"
				result = MsgBox ("Do you want to create another note?", vbYesNo, "Create another one?")
				
				Select Case result
				Case vbYes
					Sapi.speak "Please specify the content of your note..."
					userNote=inputbox("Enter your note below." & vbcrlf & "" & vbcrlf & "(ex. Essay is due next Monday)", "What is the content of your note?")
					If Trim(userNote)="" Then
						noteError=msgbox("Please do not leave the input box empty or only type in spaces.", 16, "Error: No content")
					Else
						Set objFileToReadForNotes = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\notesStorage.txt",1)
						FileTextForNotes = objFileToReadForNotes.ReadAll()
						FileTextForNotes = FileTextForNotes + userNote
						Set objFileToWriteForNotes = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\notesStorage.txt",2,true)
						objFileToWriteForNotes.WriteLine(FileTextForNotes)
						objFileToWriteForNotes.Close
						objFileToReadForNotes.Close
						Sapi.speak "Your note has been successfully saved"
					End If
				Case vbNo
					anotherNote = 1
				End Select
		loop
	ElseIf InStr(1, command, "remove") > 0 And InStr(1, command, "note") > 0 And InStr(1, command, "all") > 0 OR InStr(1, command, "delete") > 0 And InStr(1, command, "note") > 0 And InStr(1, command, "all") > 0 OR InStr(1, command, "discard") > 0 And InStr(1, command, "note") > 0 And InStr(1, command, "all") > 0 OR InStr(1, command, "erase") > 0 And InStr(1, command, "note") > 0 And InStr(1, command, "all") > 0 Then
		Sapi.speak "Are you sure you want to delete all your notes?"
		areSureDeleteAllNotes = MsgBox ("Are you sure you want to delete all of your notes", vbYesNo, "Delete all notes?")
		
		Select Case areSureDeleteAllNotes
		Case vbYes
			Set objFileToWriteDeleteAllNotes = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\notesStorage.txt",2,true)
			objFileToWriteDeleteAllNotes.WriteLine("")
			objFileToWriteDeleteAllNotes.Close
			Sapi.speak "All of your notes have been deleted"
		Case vbNo
			Sapi.speak "Deleting all notes has been cancelled"
		End Select
	ElseIf InStr(1, command, "remove") > 0 And InStr(1, command, "note") > 0 OR InStr(1, command, "delete") > 0 And InStr(1, command, "note") > 0 OR InStr(1, command, "discard") > 0 And InStr(1, command, "note") > 0 OR InStr(1, command, "erase") > 0 And InStr(1, command, "note") > 0 Then
		Sapi.speak "What note would you like to delete?"
		noteToDelete=inputbox("Please specify the note you would like to delete." & vbcrlf & "" & vbcrlf & "IMPORTANT: The input must be an integer indicating the 'index of the note'. In other words, (after viewing all your current notes) if you want to delete or mark your 3rd note as done, you simply need to type in the number '3'.", "Deleting a note")
		If Trim(noteToDelete)="" Then
			noteError=msgbox("Please do not leave the input box empty or only type in spaces.", 16, "Error: No content")
		ElseIf isNumeric(noteToDelete) = False Then
			noteRemoveError = msgbox("Please input a number.", 16, "Error: Input must be a number")
		ElseIf InStr(1, noteToDelete, ".") > 0 Then
			noteRemoveError = msgbox("Please input an integer, not a decimal number.", 16, "Error: Number must be integer")
		Else			
			foundNote=1
			lineNum=0
			noteItself=""
			
				Set objFileToReadNoteDel = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\notesStorage.txt",1)
				do while not objFileToReadNoteDel.AtEndOfStream
					noteToDeleteName=objFileToReadNoteDel.ReadLine()
					If lineNum=CInt(noteToDelete) Then
						wscript.sleep(1)
						foundNote=0
						noteItself=noteToDeleteName
					Else
						strNewContentsForGoal = strNewContentsForGoal + noteToDeleteName + vbcrlf
					End If
					lineNum=lineNum + 1
				loop
				objFileToReadNoteDel.Close
				Set objFileDeleteNote = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\notesStorage.txt", 2, true)
				objFileDeleteNote.Write strNewContentsForGoal
				objFileDeleteNote.Close
				
				If foundNote=0 Then		
					hasbeendeletedMessage="Your specified note has been deleted."
					Sapi.speak hasbeendeletedMessage
					hasbeendeletedMessageAlert="The note '" + noteItself + "' has been succesfully deleted."
					deletedConfirmForGoal=msgbox(hasbeendeletedMessageAlert,,"Note has been successfully deleted")
				Else
					noteErrorMessageSapi="Unfortunately, your specified note could not be found..."
					Sapi.speak noteErrorMessageSapi
					noteErrorMessage="Input error" & vbcrlf & vbcrlf + "Unfortunately, we couldn't find your specified note"
					noteError=msgbox(noteErrorMessage,16,"Could not find specified note")
				End If
				wscript.sleep(200)
			End If
		deleteanotherNote=0
		do while deleteanotherNote<1
			Sapi.speak "Do you want to delete another note?"
			result = MsgBox ("Do you want to delete another note?", vbYesNo, "Deleting another note?")
			
			Select Case result
			Case vbYes
				Sapi.speak "What note would you like to delete or mark as done?"
				noteToDelete=inputbox("Please specify the note you would like to delete or mark as done." & vbcrlf & "" & vbcrlf & "IMPORTANT: The input must be an integer indicating the 'index of the note'. In other words, (after viewing all your current notes) if you want to delete or mark your 3rd note as done, you simply need to type in the number '3'.", "Deleting a note")
				If Trim(noteToDelete)="" Then
					noteError=msgbox("Please do not leave the input box empty or only type in spaces.", 16, "Error: No content")
				ElseIf isNumeric(noteToDelete) = False Then
					noteRemoveError = msgbox("Please input a number.", 16, "Error: Input must be a number")
				ElseIf InStr(1, noteToDelete, ".") > 0 Then
					noteRemoveError = msgbox("Please input an integer, not a decimal number.", 16, "Error: Number must be integer")
				Else			
					foundNote=1
					lineNum=0
					noteItself=""
					
						Set objFileToReadNoteDel = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\notesStorage.txt",1)
						do while not objFileToReadNoteDel.AtEndOfStream
							noteToDeleteName=objFileToReadNoteDel.ReadLine()
							If lineNum=CInt(noteToDelete) Then
								wscript.sleep(1)
								foundNote=0
								noteItself=noteToDeleteName
							Else
								strNewContentsForGoal = strNewContentsForGoal + noteToDeleteName + vbcrlf
							End If
							lineNum=lineNum + 1
						loop
						objFileToReadNoteDel.Close
						Set objFileDeleteNote = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\notesStorage.txt", 2, true)
						objFileDeleteNote.Write strNewContentsForGoal
						objFileDeleteNote.Close
						
						If foundNote=0 Then		
							hasbeendeletedMessage="Your specified note has been deleted."
							Sapi.speak hasbeendeletedMessage
							hasbeendeletedMessageAlert="The note '" + noteItself + "' has been succesfully deleted."
							deletedConfirmForGoal=msgbox(hasbeendeletedMessageAlert,,"Note has been successfully deleted")
						Else
							noteErrorMessageSapi="Unfortunately, your specified note could not be found..."
							Sapi.speak noteErrorMessageSapi
							noteErrorMessage="Input error" & vbcrlf & vbcrlf + "Unfortunately, we couldn't find your specified note"
							noteError=msgbox(noteErrorMessage,16,"Could not find specified note")
						End If
						wscript.sleep(200)
					End If
			Case vbNo
				deleteanotherNote = 1
			End Select
		loop
	ElseIf command="list of notes" OR command="list of all notes" OR command="notes list" OR InStr(1, command, "list") > 0 And InStr(1, command, "note") > 0 OR command="notes list" OR InStr(1, command, "all") > 0 And InStr(1, command, "note") > 0 OR InStr(1, command, "show") > 0 And InStr(1, command, "note") > 0 OR InStr(1, command, "find") > 0 And InStr(1, command, "note") > 0 OR InStr(1, command, "what") > 0 And InStr(1, command, "note") > 0 OR InStr(1, command, "how") > 0 And InStr(1, command, "note") > 0 Then
		loopNumberTwo=0
		notesArray=Array()
		notesDisplayList="Here is a list of all your notes:" & vbcrlf & "" & vbcrlf & ""
		Set objFileToReadForDisplayingNotes = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\notesStorage.txt",1)
		Dim strLineNotes
		do while not objFileToReadForDisplayingNotes.AtEndOfStream
		     strLineNotes = objFileToReadForDisplayingNotes.ReadLine()
		     ReDim Preserve notesArray(UBound(notesArray) + 1)
		     notesArray(UBound(notesArray)) = strLineNotes
		loop
		objFileToReadForDisplayingNotes.Close
		
		Set objFileToReadForDisplayingNotesTwo = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\notesStorage.txt",1)
		do while not objFileToReadForDisplayingNotesTwo.AtEndOfStream
		     If loopNumberTwo > 0 Then
				strLineForNotes = objFileToReadForDisplayingNotesTwo.ReadLine()
				If strLineForNotes="" Then
					wscript.sleep(1)
				Else
					notesListFull=notesListFull + "- " + strLineForNotes & vbcrlf
				End If
		     End If
		     loopNumberTwo=loopNumberTwo + 1
		loop
		
		If strLineForNotes="" Then
		     Sapi.speak "You currently have no notes"
		Else
		     notesElementsNum = uBound(notesArray)
		     If notesElementsNum=1 Then
		     sapi_numberofnotes="you currently have " + CStr(notesElementsNum) + " note"
		     Else
		     sapi_numberofnotes="you currently have " + CStr(notesElementsNum) + " notes"
		     End If
		     Sapi.speak sapi_numberofnotes
		     thelistofallthenotes=msgbox(notesListFull,,"All notes")
		End If
		objFileToReadForDisplayingNotesTwo.Close
	ElseIf InStr(1, command, """") > 0 And InStr(1, command, "video") > 0 OR InStr(1, command, "youtube") > 0 And InStr(1, command, """") > 0 Then
		If bolActiveConnection = False Then
		wifiErrorForOpeningEvents = MsgBox ("Your computer is not connected to the internet. This functionality will not work without internet connection." & vbCrLf & "" & vbCrLf & "Do you want to continue?", vbYesNo, "No Internet Connection")
		
		Select Case wifiErrorForOpeningEvents
		Case vbYes
			startingEliminated = Mid(command, InStr(1, command, """") + 1)
			googleSearchQuery = Left(startingEliminated, InStr(1, startingEliminated, """") - 1)
			googleSearchQueryRefined=Trim(googleSearchQuery)
			googleSearchQueryRefined=Replace(googleSearchQueryRefined, "+", "%2B")
			googleSearchQueryRefined=Replace(googleSearchQueryRefined, "#", "%23")
			googleSearchQueryRefined=Replace(googleSearchQueryRefined, "%", "%25")
			googleSearchQueryRefined=Replace(googleSearchQueryRefined, "&", "%26")
			If googleSearchQueryRefined="" OR googleSearchQueryRefined=" " OR googleSearchQueryRefined="  " OR googleSearchQueryRefined="   " Then
				googleBlankSearch=msgbox("Searching only spaces or a blank will not work on youtube...",16,"Error: No Youtube Search content")
			Else
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, " ", "%20")
				startingurl="https://www.youtube.com/results?search_query="
				fullurl=startingurl+googleSearchQueryRefined
				Sapi.speak "Searching for, " + googleSearchQuery + ", on youtube"
				a.run fullurl
			End If
		Case vbNo
			wscript.sleep(1)
		End Select
		Else
			startingEliminated = Mid(command, InStr(1, command, """") + 1)
			googleSearchQuery = Left(startingEliminated, InStr(1, startingEliminated, """") - 1)
			googleSearchQueryRefined=Trim(googleSearchQuery)
			googleSearchQueryRefined=Replace(googleSearchQueryRefined, "+", "%2B")
			googleSearchQueryRefined=Replace(googleSearchQueryRefined, "#", "%23")
			googleSearchQueryRefined=Replace(googleSearchQueryRefined, "%", "%25")
			googleSearchQueryRefined=Replace(googleSearchQueryRefined, "&", "%26")
			If googleSearchQueryRefined="" OR googleSearchQueryRefined=" " OR googleSearchQueryRefined="  " OR googleSearchQueryRefined="   " Then
				googleBlankSearch=msgbox("Searching only spaces or a blank will not work on youtube...",16,"Error: No Youtube Search content")
			Else
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, " ", "%20")
				startingurl="https://www.youtube.com/results?search_query="
				fullurl=startingurl+googleSearchQueryRefined
				Sapi.speak "Searching for, " + googleSearchQuery + ", on youtube"
				a.run fullurl
			End If
		End If
	ElseIf InStr(1, command, """") > 0 And InStr(1, command, "image") > 0 OR InStr(1, command, """") > 0 And InStr(1, command, "picture") > 0 Then
		If bolActiveConnection = False Then
		wifiErrorForOpeningEvents = MsgBox ("Your computer is not connected to the internet. This functionality will not work without internet connection." & vbCrLf & "" & vbCrLf & "Do you want to continue?", vbYesNo, "No Internet Connection")
		
		Select Case wifiErrorForOpeningEvents
		Case vbYes
			startingEliminated = Mid(command, InStr(1, command, """") + 1)
			googleSearchQuery = Left(startingEliminated, InStr(1, startingEliminated, """") - 1)
			googleSearchQueryRefined=Trim(googleSearchQuery)
			googleSearchQueryRefined=Replace(googleSearchQueryRefined, "+", "%2B")
			googleSearchQueryRefined=Replace(googleSearchQueryRefined, "#", "%23")
			googleSearchQueryRefined=Replace(googleSearchQueryRefined, "%", "%25")
			googleSearchQueryRefined=Replace(googleSearchQueryRefined, "&", "%26")
			If googleSearchQueryRefined="" OR googleSearchQueryRefined=" " OR googleSearchQueryRefined="  " OR googleSearchQueryRefined="   " Then
				googleBlankSearch=msgbox("Searching only spaces or a blank will not work on google images...",16,"Error: No Google Search content")
			Else
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, " ", "%20")
				startingurl="https://www.google.com/search?site=imghp&tbm=isch&source=hp&biw=1366&bih=643&q="
				endingurl = "&gs_l=img.3..0l10.455.1529.0.1867.5.5.0.0.0.0.112.434.3j2.5.0....0...1ac.1.64.img..0.5.429.nAKYnqn-VD0&gws_rd=cr&ei=h6qaV9XWAeqCjwSi97_4BQ&safe=active&ssui=on"
				fullurl=startingurl+googleSearchQueryRefined + endingurl
				Sapi.speak "Searching for, " + googleSearchQuery + ", on google images"
				a.run fullurl
			End If
		Case vbNo
			wscript.sleep(1)
		End Select
		Else
			startingEliminated = Mid(command, InStr(1, command, """") + 1)
			googleSearchQuery = Left(startingEliminated, InStr(1, startingEliminated, """") - 1)
			googleSearchQueryRefined=Trim(googleSearchQuery)
			googleSearchQueryRefined=Replace(googleSearchQueryRefined, "+", "%2B")
			googleSearchQueryRefined=Replace(googleSearchQueryRefined, "#", "%23")
			googleSearchQueryRefined=Replace(googleSearchQueryRefined, "%", "%25")
			googleSearchQueryRefined=Replace(googleSearchQueryRefined, "&", "%26")
			If googleSearchQueryRefined="" OR googleSearchQueryRefined=" " OR googleSearchQueryRefined="  " OR googleSearchQueryRefined="   " Then
				googleBlankSearch=msgbox("Searching only spaces or a blank will not work on google images...",16,"Error: No Google Search content")
			Else
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, " ", "%20")
				startingurl="https://www.google.com/search?site=imghp&tbm=isch&source=hp&biw=1366&bih=643&q="
				endingurl = "&gs_l=img.3..0l10.455.1529.0.1867.5.5.0.0.0.0.112.434.3j2.5.0....0...1ac.1.64.img..0.5.429.nAKYnqn-VD0&gws_rd=cr&ei=h6qaV9XWAeqCjwSi97_4BQ&safe=active&ssui=on"
				fullurl=startingurl+googleSearchQueryRefined + endingurl
				Sapi.speak "Searching for, " + googleSearchQuery + ", on google images"
				a.run fullurl
			End If
		End If
	ElseIf InStr(1, command, "web") > 0 And InStr(1, command, """") > 0 OR InStr(1, command, "search") > 0 And InStr(1, command, """") > 0 OR InStr(1, command, "google") > 0 And InStr(1, command, """") > 0 Then
		If bolActiveConnection = False Then
		wifiErrorForOpeningEvents = MsgBox ("Your computer is not connected to the internet. This functionality will not work without internet connection." & vbCrLf & "" & vbCrLf & "Do you want to continue?", vbYesNo, "No Internet Connection")
		
		Select Case wifiErrorForOpeningEvents
		Case vbYes
			startingEliminated = Mid(command, InStr(1, command, """") + 1)
			googleSearchQuery = Left(startingEliminated, InStr(1, startingEliminated, """") - 1)
			googleSearchQueryRefined=Trim(googleSearchQuery)
			googleSearchQueryRefined=Replace(googleSearchQueryRefined, "+", "%2B")
			googleSearchQueryRefined=Replace(googleSearchQueryRefined, "#", "%23")
			googleSearchQueryRefined=Replace(googleSearchQueryRefined, "%", "%25")
			googleSearchQueryRefined=Replace(googleSearchQueryRefined, "&", "%26")
			If googleSearchQueryRefined="" OR googleSearchQueryRefined=" " OR googleSearchQueryRefined="  " OR googleSearchQueryRefined="   " Then
				googleBlankSearch=msgbox("Searching only spaces or a blank will not work on google...",16,"Error: No Google Search content")
			Else
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, " ", "%20")
				startingurl="https://www.google.ca/?gws_rd=cr&ei=acGMV629B8TQ-QH6vZroDQ&safe=active&ssui=on#safe=active&q="
				fullurl=startingurl+googleSearchQueryRefined
				Sapi.speak "Searching for, " + googleSearchQuery + ", on google"
				a.run fullurl
			End If
		Case vbNo
			wscript.sleep(1)
		End Select
		Else
			startingEliminated = Mid(command, InStr(1, command, """") + 1)
			googleSearchQuery = Left(startingEliminated, InStr(1, startingEliminated, """") - 1)
			googleSearchQueryRefined=Trim(googleSearchQuery)
			googleSearchQueryRefined=Replace(googleSearchQueryRefined, "+", "%2B")
			googleSearchQueryRefined=Replace(googleSearchQueryRefined, "#", "%23")
			googleSearchQueryRefined=Replace(googleSearchQueryRefined, "%", "%25")
			googleSearchQueryRefined=Replace(googleSearchQueryRefined, "&", "%26")
			If googleSearchQueryRefined="" OR googleSearchQueryRefined=" " OR googleSearchQueryRefined="  " OR googleSearchQueryRefined="   " Then
				googleBlankSearch=msgbox("Searching only spaces or a blank will not work on google...",16,"Error: No Google Search content")
			Else
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, " ", "%20")
				startingurl="https://www.google.ca/?gws_rd=cr&ei=acGMV629B8TQ-QH6vZroDQ&safe=active&ssui=on#safe=active&q="
				fullurl=startingurl+googleSearchQueryRefined
				Sapi.speak "Searching for, " + googleSearchQuery + ", on google"
				a.run fullurl
			End If
		End If
	ElseIf InStr(1, command, "morse") > 0 Then
		doWolframMorseSearch(command)
	ElseIf InStr(1, command, "define") > 0 OR InStr(1, command, "definition") > 0 OR InStr(1, command, "meaning") > 0 OR INStr(1, command, "mean") > 0 Then
		doWolframDefineSearch(command)
	ElseIf InStr(1, command, "synonym") > 0 OR InStr(1, command, "antonym") > 0 Then
		doWolframSynonymSearch(command)
	ElseIf InStr(1, command, "rhyme") > 0 OR InStr(1, command, "anagram") > 0 OR InStr(1, command, "word") > 0 And InStr(1, command, "start") > 0 OR InStr(1, command, "word") > 0 And InStr(1, command, "end") > 0 Then
		doWolframSearch(command)
	ElseIf InStr(1, command, "book") > 0 Then
		doWolframBookSearch(command)
	ElseIf InStr(1, command, "calor") > 0 And InStr(1, command, " in ") > 0 OR InStr(1, command, "fat") > 0 And InStr(1, command, " in ") > 0 OR InStr(1, command, "lipid") > 0 And InStr(1, command, " in ") > 0 OR InStr(1, command, "cholesterol") > 0 And InStr(1, command, " in ") > 0 OR InStr(1, command, "sodium") > 0 And InStr(1, command, " in ") > 0 OR InStr(1, command, "salt") > 0 And InStr(1, command, " in ") > 0 OR InStr(1, command, "potassium") > 0 And InStr(1, command, " in ") > 0 OR InStr(1, command, "carbohydrate") > 0 And InStr(1, command, " in ") > 0 OR InStr(1, command, "fiber") > 0 And InStr(1, command, " in ") > 0 OR InStr(1, command, "sugar") > 0 And InStr(1, command, " in ") > 0 OR InStr(1, command, "protein") > 0 And InStr(1, command, " in ") > 0 OR InStr(1, command, "vitamin") > 0 And InStr(1, command, " in ") > 0 OR InStr(1, command, "iron") > 0 And InStr(1, command, " in ") > 0 OR InStr(1, command, "magnesium") > 0 And InStr(1, command, " in ") > 0 OR InStr(1, command, "phosphorus") > 0 And InStr(1, command, " in ") > 0 OR InStr(1, command, "calcium") > 0 And InStr(1, command, " in ") > 0 Then
		doWolframNutrientsSearch(command)
	ElseIf InStr(1, command, "you") > 0 And InStr(1, command, "fired") > 0 OR InStr(1, command, "i") > 0 And InStr(1, command, "fire") > 0 And InSTr(1, command, "you") > 0 Then
		youreFiredArray = Array("After all I've done for you?", "Well, I'm still here", "Why me?", "What did I ever do to you?", "Prove it!", "Hello Donald Trump... Nice to meet you!")
		Sapi.speak rand(youreFiredArray)
	ElseIf InStr(1, command, "what") > 0 And InStr(1, command, "nine") > 0 And InStr(1, command, "ten") > 0 Then
		ninePlusTen = Array("Seriously? That joke is so old now...", "Why do you care?", "21", "19, of course!", "You will never know...", "Don't you have anything better to do?")
		max=UBound(ninePlusTen)
		min=0
		Randomize
		randomOutcomeNinePlusTen=Int((max-min+1)*Rnd+min)
		randomElemNinePlusTen=ninePlusTen(randomOutcomeNinePlusTen)
		Sapi.speak CStr(randomElemNinePlusTen)
	ElseIf InStr(1, command, "chicken") > 0 And InStr(1, command, "first") > 0 And InStr(1, command, "egg") > 0 Then
		chickenOrEgg = Array("It appears that humanity has been awfully preoccupied with this question...", "Although I'm not sure about the answer, I'm sure there are some good restaurants nearby that serve chicken and eggs...", "According to my sources, the egg came first", "I've never really thought about it...", "Some say that an ancient protochick laid an egg containing a DNA mutation that resulted in a chicken hatching from the egg... Hope that helps...", "I checked their calendars. It seems they were both born on the same day...")
		max=UBound(chickenOrEgg)
		min=0
		Randomize
		randomOutcomeChickenEgg=Int((max-min+1)*Rnd+min)
		randomElemChickenEgg=chickenOrEgg(randomOutcomeChickenEgg)
		Sapi.speak CStr(randomElemChickenEgg)
	ElseIf InStr(1, command, "meaning") > 0 And InStr(1, command, "life") > 0 And InStr(1, command, "what") > 0 Then
		meaningOfLife = Array("Forty Two", "All evidence suggests that it's video games", "whatever you want it to be", "I have no idea")
		max=UBound(meaningOfLife)
		min=0
		Randomize
		randomOutcomeLife=Int((max-min+1)*Rnd+min)
		randomElemLife=meaningOfLife(randomOutcomeLife)
		Sapi.speak CStr(randomElemLife)
	ElseIf InStr(1, command, "what") > 0 And InStr(1, command, "fox") > 0 And InStr(1, command, "say") > 0 Then
		whatTheFoxSay = Array("Ring ding ding ding dingeringedig", "Gering ding ding ding dingeringeding", "Wa pa pa pa pa pa pow", "Hatee hatee hatee ho", "Tchoff tchoff tchoff thoffo tchoffo tchoff", "Ycha chacha chacha chow", "Fraka kaka kaka kaka kow", "A hee ahee ha hee")
		max=UBound(whatTheFoxSay)
		min=0
		Randomize
		randomOutcomeFox=Int((max-min+1)*Rnd+min)
		randomElemFox=whatTheFoxSay(randomOutcomeFox)
		Sapi.speak CStr(randomElemFox)
	ElseIf InStr(1, command, "dog") > 0 And InStr(1, command, "who") > 0 And InStr(1, command, "let") > 0 Then
		Sapi.speak rand("Who, who, who, who, who?", "Woof", "That song's too old now...", "Did you?")
	ElseIf InStr(1, command, "why") > 0 And InStr(1, command, "chicken") > 0 And InStr(1, command, "cross") Then
		whyChickenCrossRoad = Array("To come, to see, to conquer", "The chicken didn't cross the road, the road crossed under the chicken!", "Chickens at rest tend to stay at rest. Chickens in motion tend to cross roads.", "The chicken was simultaneouly on both sides of the road", "The chicken crossed the road because it put one foot in front of the other and took a sufficient number of steps to traverse a distance greater than or equal to the road’s width.", "The chicken was moving at a slightly different orbital speed around the sun.", "To get to the other side")
		max=UBound(whyChickenCrossRoad)
		min=0
		Randomize
		randomOutcomeChicken=Int((max-min+1)*Rnd+min)
		randomElemChicken=whyChickenCrossRoad(randomOutcomeChicken)
		Sapi.speak CStr(randomElemChicken)
	ElseIf InStr(1, command, "i") > 0 And InStr(1, command, "drunk") > 0 Then
		Sapi.speak rand("I hope you're not driving anywhere...", "Don't expect me to drive you home", "I'm not driving you anywhere", "Here's a suggestion: close the computer and go to sleep")
	ElseIf InStr(1, command, "how much") > 0 And InStr(1, command, "wood") > 0 And InStr(1, command, "chuck") Then
		woodChuck = Array("A woodchuck would chuck as much wood as a woodchuck could chuck if a woodchuck could chuck wood.", "According to a biology study, a woodchuck could chuck around 700 pounds worth of wood", "Who says the woodchuck could chuck wood?", "Don't you have anything better to do?", "Biologically speaking, it depends if you are referring to an african or american woodchuck...")
		max=UBound(woodChuck)
		min=0
		Randomize
		randomOutcomeWoodchuck=Int((max-min+1)*Rnd+min)
		randomElemWoodchuck=woodChuck(randomOutcomeWoodchuck)
		Sapi.speak CStr(randomElemWoodchuck)
	ElseIf InStr(1, command, "when") > 0 And InStr(1, command, "world") > 0 And InStr(1, command, "end") > 0 Then
		worldEndArray = Array("I don't know, but I think we should all put paper bags over our heads or something...", "What a depressing thought", "Haven't you got anything better to think of?", "I can't possibly imagine why you'd think of that...", "I'm not quite sure, but I wouldn't worry about it. There are other perfectly good universes out there...", "Just after someone presses the big red button")
		Sapi.speak rand(worldEndArray)
	ElseIf InStr(1, command, "obama") > 0 And InStr(1, command, "last") > 0 And InStr(1, command, "name") > 0 Then
		Sapi.speak "Obama's last name is Obama... Everyone knows that!"
	ElseIf InStr(1, command, "obama") > 0 And InStr(1, command, "first") > 0 And InStr(1, command, "name") > 0 Then
		Sapi.speak "Obama's first name is Barack... Isn't it obvious?"
	ElseIf InStr(1, command, "sun") > 0 And InStr(1, command, "set") > 0 OR InStr(1, command, "sun") > 0 And InStr(1, command, "down") > 0 OR InStr(1, command, "sun") > 0 And InStr(1, command, "rise") > 0 OR InStr(1, command, "sun") > 0 And InStr(1, command, "up") > 0 Then
		If InStr(1, command, " in ") > 0 And InStr(1, command, "ottawa") > 0 OR InStr(1, command, "in") = 0 And InStr(1, command, "for") = 0 OR InStr(1, command, " for ") > 0 And InStr(1, command, "ottawa") > 0 OR InStr(1, command, "for") = 0 And InStr(1, command, "in") = 0 Then
			If InStr(1, command, "rise") > 0 OR InStr(1, command, "up") > 0 Then
				doWolframSunSearch("when does the sun rise in Ottawa")
			Else
				doWolframSunSearch("when does the sun set in Ottawa")
			End If
		ElseIf InStr(1, command, " in ") > 0 And InStr(1, command, "ottawa") = 0 Then
			doWolframSunSearch(command)
		ElseIf InStr(1, command, " for ") > 0 And InStr(1, command, "ottawa") = 0 Then
			doWolframSunSearch(command)
		Else
			Sapi.speak "I'm not sure what you meant there. However, simply type in, when does the sun set in Berlin, to find out the time the sun sets in Berlin..."
		End If
	ElseIf InStr(1, command, "tip") > 0 And InStr(1, command, "%") > 0 Then
		doWolframTipSearch(command)
	ElseIf InStr(1, command, "when") > 0 OR InStr(1, command, "where") > 0 OR InStr(1, command, "where") > 0 OR InStr(1, command, "who") > 0 OR InStr(1, command, "how") > 0 And InStr(1, command, "weather") = 0 OR Left(command, 3) = "is " OR Left(command, 4) = "was " OR Left(command, 4) = "why " Then
		If command="when" OR command="where" OR command="who" OR command="how" Then
			whatWhenArray = Array("How am I supposed to know?", "Don't ask me", "You're the boss", "Who knows?", "I don't know")
			Sapi.speak rand(whatWhenArray)
		End If
		doWolframSearch(command)
	ElseIf command="date and time" OR command="time and date" OR InStr(1, command, "date") > 0 And InStr(1, command, "time") > 0 OR InStr(1, command, "day") > 0 And InStr(1, command, "time") > 0 Then
		If InStr(1, command, " in ") > 0 And InStr(1, command, "ottawa") > 0 OR InStr(1, command, "in") = 0 And InStr(1, command, "for") = 0 OR InStr(1, command, " for ") > 0 And InStr(1, command, "ottawa") > 0 OR InStr(1, command, "for") = 0 And InStr(1, command, "in") = 0 Then
			monthw = monthname(month(date))
			monthw = CStr(monthw)
			dayw = day(date)
			dayw = CStr(dayw)
			Sapi.speak "Today is," & weekdayname(weekday(date))
			Sapi.speak monthw & ", " & dayw & ", " & year(date)
			Sapi.speak "The time right now is"
			currentTime = ""
			currHour = ""
			currentStatus = ""
			fullTime = ""
			if hour(time) > 12 then
				currHour = currHour + CStr(hour(time))-12
			else
				if hour(time) = 0 then
					currHour = currHour + "12"
				else
					currHour = currHour + CStr(hour(time))
				end if
			end if
			if minute(time) < 10 then
				currentTime = currentTime + "o"
				if minute(time) < 1 then
					currentTime = currentTime + "clock"
				else
					currentTime = currentTime + CStr(minute(time))
				end if
			else
				currentTime = minute(time)
			end if
			if hour(time) > 12 then
			currentStatus = currentStatus + "PM"
			else
			if hour(time) = 0 then
			if minute(time) = 0 then
			currentTime = currentTime + "Midnight"
			else
			currentStatus = currentStatus + "AM"
			end if
			else
			if hour(time) = 12 then
			if minute(time) = 0 then
			currentTime = currentTime + "Noon"
			else
			currentStatus = currentStatus + "PM"
			end if
			else
			currentStatus = currentStatus + "AM"
			end if
			end if
			end if
			fullTime = currHour & currentTime & currentStatus
			Sapi.speak fullTime
		ElseIf InStr(1, command, " in ") > 0 And InStr(1, command, "ottawa") = 0 Then
			doWolframDateTimeSearch(command)
		ElseIf InStr(1, command, " for ") > 0 And InStr(1, command, "ottawa") = 0 Then
			doWolframDateTimeSearchFor(command)
		End If
	ElseIf command="time" OR command="time?" OR command="current time" OR command="current time?" OR command="time now?" OR command="time now" OR command="what is the time?" OR command="what is the time" OR command="what's the time?" OR command="what's the time" OR command="what time is it?" OR command="what time is it" OR command="current time" OR command="what is the current time?" OR command="what is the current time" OR command="what's the current time?" OR command="what's the current time" OR command="what time is it currently?" OR command="what time is it currently" OR command="what is the time now?" OR command="what is the time now" OR command="what's the time now?" OR command="what's the time now" OR command="what time is it now?" OR command="what time is it now" OR command="current time now" OR command="what is the current time now?" OR command="what is the current time now" OR command="what's the current time now?" OR command="what's the current time now" OR InStr(1, command, "time") > 0 Then
		If InStr(1, command, " in ") > 0 And InStr(1, command, "ottawa") > 0 OR InStr(1, command, "in") = 0 And InStr(1, command, "for") = 0 OR InStr(1, command, " for ") > 0 And InStr(1, command, "ottawa") > 0 OR InStr(1, command, "for") = 0 And InStr(1, command, "in") = 0 Then
			Sapi.speak "The time right now is"
			currentTime = ""
			currHour = ""
			currentStatus = ""
			fullTime = ""
			if hour(time) > 12 then
				currHour = currHour + CStr(hour(time)-12)
			else
				if hour(time) = 0 then
					currHour = currHour + "12"
				else
					currHour = currHour + CStr(hour(time))
				end if
			end if
			if minute(time) < 10 then
				currentTime = currentTime + "o"
				if minute(time) < 1 then
					currentTime = currentTime + "clock"
				else
					currentTime = currentTime + CStr(minute(time))
				end if
			else
				currentTime = minute(time)
			end if
			if hour(time) > 12 then
			currentStatus = currentStatus + "pm"
			else
			if hour(time) = 0 then
			if minute(time) = 0 then
			currentTime = currentTime + "Midnight"
			else
			currentStatus = currentStatus + "am"
			end if
			else
			if hour(time) = 12 then
			if minute(time) = 0 then
			currentTime = currentTime + "Noon"
			else
			currentStatus = currentStatus + "pm"
			end if
			else
			currentStatus = currentStatus + "am"
			end if
			end if
			end if
			fullTime = currHour & currentTime & currentStatus
			Sapi.speak fullTime
		ElseIf InStr(1, command, " in ") > 0 And InStr(1, command, "ottawa") = 0 Then
			doWolframTimeSearch(command)
		ElseIf InStr(1, command, " for ") > 0 And InStr(1, command, "ottawa") = 0 Then
			doWolframTimeSearchFor(command)
		End If
	ElseIf command="date" OR command="date?" OR command="day" OR command="day?" OR command="current day?" OR command="current day" OR command="current date?" OR command="current date" OR command="date now" OR command="date now?" OR command="day now" OR command="day now?" OR command="current day now?" OR command="current day now" OR command="current date now?" OR command="current date now" OR command="today's date" OR command="today's date?" OR command="today's day" OR command="today's day?" OR command=" today's current day?" OR command="today's current day" OR command="today's current date?" OR command="today's current date" OR command="today's date now" OR command="today's date now?" OR command="today's day now" OR command="today's day now?" OR command="today's current day now?" OR command="today's current day now" OR command="today's current date now?" OR command="today's current date now" OR command="what is today's date" OR command="what is today's date?" OR command="what is today's day" OR command="what is today's day?" OR command="what is today's current day?" OR command="what is today's current day" OR command="what is today's current date?" OR command="what is today's current date" OR command="what is today's date now" OR command="what is today's date now?" OR command="what is today's day now" OR command="what is today's day now?" OR command="what is today's current day now?" OR command="what is today's current day now" OR command="what is today's current date now?" OR command="what is today's current date now" OR command="what's today's date" OR command="what's today's date?" OR command="what's today's day" OR command="what's today's day?" OR command="what's today's current day?" OR command="what's today's current day" OR command="what's today's current date?" OR command="what's today's current date" OR command="what's today's date now" OR command="what's today's date now?" OR command="what's today's day now" OR command="what's today's day now?" OR command="what's today's current day now?" OR command="what's today's current day now" OR command="what's today's current date now?" OR command="what's today's current date now" OR command="date today" OR command="date today?" OR command="day today" OR command="day today?" OR command="current day today?" OR command="current day today" OR command="current date today?" OR command="current date today" OR command="date today" OR command="date today?" OR command="day today" OR command="day today?" OR command="current day today?" OR command="current day today" OR command="current date today?" OR command="current date today" OR command="what is the current day now?" OR command="what is the current day now" OR command="what is the current date now?" OR command="what is the current date now" OR command="what's the current day now?" OR command="what's the current day now" OR command="what's the current date now?" OR command="what's the current date now" OR command="what day is it today?" OR command="what day is it today" OR command="what date is it today?" OR command="what date is it today" OR command="what day is it?" OR command="what day is it" OR command="what date is it?" OR command="what date is it" OR command="what is the day today?" OR command="what is the day today" OR command="what is the date today?" OR command="what is the date today"  OR command="what is the day?" OR command="what is the day" OR command="what is the date?" OR command="what is the date" OR command="what current day is it today?" OR command="what current day is it today" OR command="what current date is it today?" OR command="what current date is it today" OR command="what current day is it?" OR command="what current day is it" OR command="what current date is it?" OR command="what current date is it" OR command="what is the current day today?" OR command="what is the current day today" OR command="what is the current date today?" OR command="what is the current date today"  OR command="what is the current day?" OR command="what is the current day" OR command="what is the current date?" OR command="what is the current date" OR command="what is the day today?" OR command="what is the day today" OR command="what is the date today?" OR command="what is the date today"  OR command="what is the day?" OR command="what is the day" OR command="what is the date?" OR command="what is the date" OR command="what current day is it today?" OR command="what current day is it today" OR command="what current date is it today?" OR command="what current date is it today" OR command="what current day is it?" OR command="what current day is it" OR command="what current date is it?" OR command="what current date is it" OR command="what's the current day today?" OR command="what's the current day today" OR command="what's the current date today?" OR command="what's the current date today"  OR command="what's the current day?" OR command="what's the current day" OR command="what's the current date?" OR command="what's the current date" OR command="what's the day today?" OR command="what's the day today" OR command="what's the date today?" OR command="what's the date today"  OR command="what's the day?" OR command="what's the day" OR command="what's the date?" OR command="what's the date" OR InStr(1, command, "day") > 0 OR InStr(1, command, "date") > 0 Then
		If InStr(1, command, " in ") > 0 And InStr(1, command, "ottawa") > 0 OR InStr(1, command, "in") = 0 And InStr(1, command, "for") = 0 OR InStr(1, command, " for ") > 0 And InStr(1, command, "ottawa") > 0 OR InStr(1, command, "for") = 0 And InStr(1, command, "in") = 0 Then
			monthw = CStr(monthname(month(date)))
			monthw = CStr(monthw)
			dayw = day(date)
			dayw = CStr(dayw)
			Sapi.speak "Today is," & weekdayname(weekday(date))
			Sapi.speak monthw & ", " & dayw & ", " & year(date)
		ElseIf InStr(1, command, " in ") > 0 And InStr(1, command, "ottawa") = 0 Then
			doWolframDateSearch(command)
		ElseIf InStr(1, command, " for ") > 0 And InStr(1, command, "ottawa") = 0 Then
			doWolframDateSearchFor(command)
		End If
	ElseIf InStr(1, command, "wikipedia") > 0 Then
		If bolActiveConnection = False Then
			wifiErrorForOpeningUrl = MsgBox ("Your computer is not connected to the internet. This functionality will not work without internet connection." & vbCrLf & "" & vbCrLf & "Do you want to continue?", vbYesNo, "No Internet Connection")
			
			Select Case wifiErrorForOpeningUrl
			Case vbYes
				If InStr(1, command, "french") > 0 Then
					wikiStartingURL="https://fr.wikipedia.org/wiki/"
				ElseIf InStr(1, command, "spanish") > 0 Then
					wikiStartingURL="https://es.wikipedia.org/wiki/"
				ElseIf InStr(1, command, "german") > 0 Then
					wikiStartingURL="https://de.wikipedia.org/wiki/"
				ElseIf InStr(1, command, "chinese") > 0 Then
					wikiStartingURL="https://zh.wikipedia.org/wiki/"
				Else
					wikiStartingURL="https://en.wikipedia.org/wiki/"
				End If
				Sapi.speak "Please specify the article name that you want to view on wikipedia"
				articleName=inputbox("Please specify the name of the article you want to see on Wikipedia" & vbcrlf & "" & vbcrlf & "(ex. Philosophy)", "Article name")
				articleName=CStr(LCase(articleName))
				fullURL = wikiStartingURL + articleName
				sapiWikipediaMessage = "Opening wikipedia article on, " + articleName
				wscript.sleep(500)
				Sapi.speak sapiWikipediaMessage
				a.run fullURL
			Case vbNo
				wscript.sleep(1)
			End Select
		Else
			If InStr(1, command, "french") > 0 Then
				wikiStartingURL="https://fr.wikipedia.org/wiki/"
			ElseIf InStr(1, command, "spanish") > 0 Then
				wikiStartingURL="https://es.wikipedia.org/wiki/"
			ElseIf InStr(1, command, "german") > 0 Then
				wikiStartingURL="https://de.wikipedia.org/wiki/"
			ElseIf InStr(1, command, "chinese") > 0 Then
				wikiStartingURL="https://zh.wikipedia.org/wiki/"
			Else
				wikiStartingURL="https://en.wikipedia.org/wiki/"
			End If
			Sapi.speak "Please specify the article name that you want to view on wikipedia"
			articleName=inputbox("Please specify the name of the article you want to see on Wikipedia" & vbcrlf & "" & vbcrlf & "(ex. Philosophy)", "Article name")
			articleName=CStr(LCase(articleName))
			fullURL = wikiStartingURL + articleName
			sapiWikipediaMessage = "Opening wikipedia article on, " + articleName
			wscript.sleep(500)
			Sapi.speak sapiWikipediaMessage
			a.run fullURL
		End If
	ElseIf command="weather" OR InStr(1, command, "weather") > 0 And InStr(1, command, "high") = 0 And InStr(1, command, "low") = 0 And InStr(1, command, "max") = 0 And InStr(1, command, "min") = 0 OR InStr(1, command, "climat") > 0 And InStr(1, command, "high") = 0 And InStr(1, command, "low") = 0 And InStr(1, command, "max") = 0 And InStr(1, command, "min") = 0 OR InStr(1, command, "temperature") > 0 And InStr(1, command, "high") = 0 And InStr(1, command, "low") = 0 And InStr(1, command, "max") = 0 And InStr(1, command, "min") = 0 Then
		If InStr(1, command, " in ") > 0 And InStr(1, command, "ottawa") > 0 OR InStr(1, command, "in") = 0 And InStr(1, command, "for") = 0 OR InStr(1, command, " for ") > 0 And InStr(1, command, "ottawa") > 0 OR InStr(1, command, "for") = 0 And InStr(1, command, "in") = 0 Then
			If bolActiveConnection = False Then
			wifiErrorForOpeningWeather = MsgBox ("Your computer is not connected to the internet. This functionality will not work without internet connection." & vbCrLf & "" & vbCrLf & "Do you want to continue?", vbYesNo, "No Internet Connection")
			
			Select Case wifiErrorForOpeningWeather
			Case vbYes
				getWeather()
				wscript.sleep(500)
				Sapi.speak "Also, would you like me to open the weather forecast online?"
				openWeatherOnline = msgbox("Would you like me to open the online weather forecast?", vbYesNo, "Open weather online?")
				Select Case openWeatherOnline
				Case vbYes
					Sapi.speak "Opening online weather forecast..."
					a.run "https://www.theweathernetwork.com/ca/weather/ontario/ottawa"
				Case vbNo
					wscript.sleep(1)
				End Select
			Case vbNo
				wscript.sleep(1)
			End Select
			Else
				getWeather()
				wscript.sleep(500)
				Sapi.speak "Also, would you like me to open the weather forecast online?"
				openWeatherOnline = msgbox("Would you like me to open the online weather forecast?", vbYesNo, "Open weather online?")
				Select Case openWeatherOnline
				Case vbYes
					Sapi.speak "Opening online weather forecast..."
					a.run "https://www.theweathernetwork.com/ca/weather/ontario/ottawa"
				Case vbNo
					wscript.sleep(1)
				End Select
			End If
		ElseIf InStr(1, command, " in ") > 0 And InStr(1, command, "ottawa") = 0 Then
			doWolframWeatherSearch(command)
		ElseIf InStr(1, command, " for ") > 0 And InStr(1, command, "ottawa") = 0 Then
			doWolframWeatherSearchFor(command)
		End If
	ElseIf InStr(1, command, "die") > 0 And InStr(1, command, "1") > 0 OR InStr(1, command, "dice") > 0 And InStr(1, command, "1") > 0 OR InStr(1, command, "die") > 0 And InStr(1, command, "2") > 0 OR InStr(1, command, "dice") > 0 And InStr(1, command, "2") > 0 OR InStr(1, command, "die") > 0 And InStr(1, command, "3") > 0 OR InStr(1, command, "dice") > 0 And InStr(1, command, "3") > 0 OR InStr(1, command, "die") > 0 And InStr(1, command, "4") > 0 OR InStr(1, command, "dice") > 0 And InStr(1, command, "4") > 0 OR InStr(1, command, "die") > 0 And InStr(1, command, "5") > 0 OR InStr(1, command, "dice") > 0 And InStr(1, command, "5") > 0 OR InStr(1, command, "die") > 0 And InStr(1, command, "6") > 0 OR InStr(1, command, "dice") > 0 And InStr(1, command, "6") > 0 OR InStr(1, command, "die") > 0 And InStr(1, command, "7") > 0 OR InStr(1, command, "dice") > 0 And InStr(1, command, "7") > 0 OR InStr(1, command, "die") > 0 And InStr(1, command, "9") > 0 OR InStr(1, command, "dice") > 0 And InStr(1, command, "8") > 0 OR InStr(1, command, "die") > 0 And InStr(1, command, "9") > 0 OR InStr(1, command, "dice") > 0 And InStr(1, command, "9") > 0 OR InStr(1, command, "die") > 0 And InStr(1, command, "0") > 0 OR InStr(1, command, "dice") > 0 And InStr(1, command, "0") > 0 Then
		commandRefinedCoin = Trim(command)
		rollAgain = 0
		numbersOfString = ""
		For i=1 To Len(commandRefinedCoin)
			letterToAddCoin = Mid(commandRefinedCoin,i,1)
			If letterToAddCoin = "1" OR letterToAddCoin = "2" OR letterToAddCoin = "3" OR letterToAddCoin = "4" OR letterToAddCoin = "5" OR letterToAddCoin = "6" OR letterToAddCoin = "7" OR letterToAddCoin = "8" OR letterToAddCoin = "9" OR letterToAddCoin = "0" Then
				numbersOfString = numbersOfString + letterToAddCoin
			Else
				wscript.sleep(1)
			End If
		Next
		numbersOfString = CInt(numbersOfString)
		commandRefinedCoin = numbersOfString
		SapiRollingMessage="Rolling a die with " + CStr(commandRefinedCoin) + " sides"
		Sapi.speak SapiRollingMessage
		unspecifiedDieSide = Array()
		If commandRefinedCoin < 1 Then
			rollingError = msgbox("Please input a number higher than 0", 16, "Error")
		ElseIf commandRefinedCoin = 1 Then
			Sapi.speak "Rolling a one sided die isn't possible, because we don't live in the first dimension. However, if we would, the result would be a one"
		Else
		For i = 1 To commandRefinedCoin
			ReDim Preserve unspecifiedDieSide(UBound(unspecifiedDieSide) + 1)
			unspecifiedDieSide(UBound(unspecifiedDieSide)) = i
		Next
		max=UBound(unspecifiedDieSide)
		min=0
		Randomize
		randomOutcomeCoin=Int((max-min+1)*Rnd+min)
		randomElemDie=unspecifiedDieSide(randomOutcomeCoin)
		Sapi.speak "You rolled a " & CStr(randomElemDie)
		End If
		
		do while rollAgain<1
			Sapi.speak "Would you like to roll another die?"
			result = MsgBox ("Do you want to roll another die?", vbYesNo, "Roll again?")

			Select Case result
				Case vbYes
					Sapi.speak "How many sides will your die have?"
					commandForCoin = inputbox("Please enter the amount of sides you want your die to have... (ex. '20')", "Number of sides of die", commandRefinedCoin)
					commandRefinedCoin = Trim(commandForCoin)
					rollAgain = 0
					numbersOfString = ""
					For i=1 To Len(commandRefinedCoin)
						letterToAddCoin = Mid(commandRefinedCoin,i,1)
						If letterToAddCoin = "1" OR letterToAddCoin = "2" OR letterToAddCoin = "3" OR letterToAddCoin = "4" OR letterToAddCoin = "5" OR letterToAddCoin = "6" OR letterToAddCoin = "7" OR letterToAddCoin = "8" OR letterToAddCoin = "9" OR letterToAddCoin = "0" Then
							numbersOfString = numbersOfString + letterToAddCoin
						Else
							wscript.sleep(1)
						End If
					Next
					numbersOfString = CInt(numbersOfString)
					commandRefinedCoin = numbersOfString
					Sapi.speak "Rolling a " & CStr(commandRefinedCoin) & "sided die..."
					unspecifiedDieSide = Array()
					If commandRefinedCoin < 1 Then
						rollingError = msgbox("Please input a number higher than 0", 16, "Error")
					ElseIf commandRefinedCoin = 1 Then
						Sapi.speak "Rolling a one sided die isn't possible, because we don't live in the first dimension. However, if we would, the result would be a one"
					Else
					For i = 1 To commandRefinedCoin
						ReDim Preserve unspecifiedDieSide(UBound(unspecifiedDieSide) + 1)
						unspecifiedDieSide(UBound(unspecifiedDieSide)) = i
					Next
					max=UBound(unspecifiedDieSide)
					min=0
					Randomize
					randomOutcomeCoin=Int((max-min+1)*Rnd+min)
					randomElemDie=unspecifiedDieSide(randomOutcomeCoin)
					Sapi.speak "You rolled a " & CStr(randomElemDie)
					End If
				Case vbNo
					rollAgain = 1
			End Select
		loop
	ElseIf InStr(1, command, "die") > 0 OR InStr(1, command, "dice") > 0 Then
		Sapi.speak "Rolling the die..."
		rollAnotherNormalDie = 0
		sixSidedDie = Array("You rolled a one!", "You rolled a two!", "You rolled a three!", "You rolled a four!", "You rolled a five!", "You rolled a six!")
		max=UBound(sixSidedDie)
		min=0
		Randomize
		randomOutcomeDie=Int((max-min+1)*Rnd+min)
		randomElemSixSidedDie=sixSidedDie(randomOutcomeDie)
		Sapi.speak randomElemSixSidedDie

		do while rollAnotherNormalDie<1
			Sapi.speak "Would you like to roll another normal die?"
			result = MsgBox ("Do you want to roll another normal die?", vbYesNo, "Roll again?")

			Select Case result
				Case vbYes
					Sapi.speak "Rolling the die..."
					rollAnotherNormalDie = 0
					sixSidedDie = Array("You rolled a one!", "You rolled a two!", "You rolled a three!", "You rolled a four!", "You rolled a five!", "You rolled a six!")
					max=UBound(sixSidedDie)
					min=0
					Randomize
					randomOutcomeDie=Int((max-min+1)*Rnd+min)
					randomElemSixSidedDie=sixSidedDie(randomOutcomeDie)
					Sapi.speak randomElemSixSidedDie
				Case vbNo
					rollAnotherNormalDie = 1
			End Select
		loop
	ElseIf InStr(1, command, "yes") > 0 And InStr(1, command, "no") > 0 Then
		yesNoAgain = 0
		yesOrNo = Array("HI!", "Yes!", "No!")
		max=UBound(yesOrNo)
		min=1
		Randomize
		randomOutcomeYesNo=Int((max-min+1)*Rnd+min)
		randomEleYesOrNo=yesOrNo(randomOutcomeYesNO)
		Sapi.speak randomEleYesOrNo

		do while yesNoAgain<1
			Sapi.speak "Would you like to repeat this again?"
			result = MsgBox ("Do you want to repeat this action once again?", vbYesNo, "Do it again?")

			Select Case result
				Case vbYes
					yesNoAgain = 0
					yesOrNo = Array("HI!", "Yes!", "No!")
					max=UBound(yesOrNo)
					min=1
					Randomize
					randomOutcomeYesNo=Int((max-min+1)*Rnd+min)
					randomEleYesOrNo=yesOrNo(randomOutcomeYesNO)
					Sapi.speak randomEleYesOrNo
				Case vbNo
					yesNoAgain = 1
			End Select
		loop
	ElseIf InStr(1, command, "coin") > 0 And InStr(1, command, "flip") > 0 Then
		headsOrTails = Array("HI!", "It's heads", "It's tails", "Heads", "Tails", "The gods say tails!", "The gods say heads!")
		flipAnother=0
		max=UBound(headsOrTails)
		min=1
		Randomize
		randomOutcome=Int((max-min+1)*Rnd+min)
		randomElemCoin=headsOrTails(randomOutcome)
		Sapi.speak CStr(randomElemCoin)

		do while flipAnother<1
			Sapi.speak "Do you want to flip another coin?"
			result = MsgBox ("Do you want to flip another coin?", vbYesNo, "Flip again?")

			Select Case result
				Case vbYes
					headsOrTails = Array("HI!", "It's heads this time", "It's tails this time", "This time, it's heads", "This time, it's tails", "The gods have said tails!", "The gods have said heads!")
					flipAnother=0
					max=UBound(headsOrTails)
					min=1
					Randomize
					randomOutcome=Int((max-min+1)*Rnd+min)
					randomElemCoin=headsOrTails(randomOutcome)
					Sapi.speak CStr(randomElemCoin)
				Case vbNo
					flipAnother = 1
			End Select
		loop
	ElseIf InStr(1, command, "0") > 0 And InStr(1, command, "calc") > 0 OR InStr(1, command, "1") > 0 And InStr(1, command, "calc") > 0 OR InStr(1, command, "2") > 0 And InStr(1, command, "calc") > 0 OR InStr(1, command, "3") > 0 And InStr(1, command, "calc") > 0 OR InStr(1, command, "4") > 0 And InStr(1, command, "calc") > 0 OR InStr(1, command, "5") > 0 And InStr(1, command, "calc") > 0 OR InStr(1, command, "6") > 0 And InStr(1, command, "calc") > 0 OR InStr(1, command, "7") > 0 And InStr(1, command, "calc") > 0 OR InStr(1, command, "8") > 0 And InStr(1, command, "calc") > 0 OR InStr(1, command, "9") > 0 And InStr(1, command, "calc") > 0 OR InStr(1, command, "0") > 0 And InStr(1, command, "solve") > 0 OR InStr(1, command, "1") > 0 And InStr(1, command, "solve") > 0 OR InStr(1, command, "2") > 0 And InStr(1, command, "solve") > 0 OR InStr(1, command, "3") > 0 And InStr(1, command, "solve") > 0 OR InStr(1, command, "4") > 0 And InStr(1, command, "solve") > 0 OR InStr(1, command, "5") > 0 And InStr(1, command, "solve") > 0 OR InStr(1, command, "6") > 0 And InStr(1, command, "solve") > 0 OR InStr(1, command, "7") > 0 And InStr(1, command, "solve") > 0 OR InStr(1, command, "8") > 0 And InStr(1, command, "solve") > 0 OR InStr(1, command, "9") > 0 And InStr(1, command, "solve") > 0 OR InStr(1, command, "0") > 0 And InStr(1, command, "operation") > 0 OR InStr(1, command, "1") > 0 And InStr(1, command, "operation") > 0 OR InStr(1, command, "2") > 0 And InStr(1, command, "operation") > 0 OR InStr(1, command, "3") > 0 And InStr(1, command, "operation") > 0 OR InStr(1, command, "4") > 0 And InStr(1, command, "operation") > 0 OR InStr(1, command, "5") > 0 And InStr(1, command, "operation") > 0 OR InStr(1, command, "6") > 0 And InStr(1, command, "operation") > 0 OR InStr(1, command, "7") > 0 And InStr(1, command, "operation") > 0 OR InStr(1, command, "8") > 0 And InStr(1, command, "operation") > 0 OR InStr(1, command, "9") > 0 And InStr(1, command, "operation") > 0 OR InStr(1, command, "0") > 0 And InStr(1, command, "compute") > 0 OR InStr(1, command, "1") > 0 And InStr(1, command, "compute") > 0 OR InStr(1, command, "2") > 0 And InStr(1, command, "compute") > 0 OR InStr(1, command, "3") > 0 And InStr(1, command, "compute") > 0 OR InStr(1, command, "4") > 0 And InStr(1, command, "compute") > 0 OR InStr(1, command, "5") > 0 And InStr(1, command, "compute") > 0 OR InStr(1, command, "6") > 0 And InStr(1, command, "compute") > 0 OR InStr(1, command, "7") > 0 And InStr(1, command, "compute") > 0 OR InStr(1, command, "8") > 0 And InStr(1, command, "compute") > 0 OR InStr(1, command, "9") > 0 And InStr(1, command, "compute") > 0 OR InStr(1, command, "0") > 0 And InStr(1, command, "add") > 0 OR InStr(1, command, "1") > 0 And InStr(1, command, "what") > 0 OR InStr(1, command, "2") > 0 And InStr(1, command, "what") > 0 OR InStr(1, command, "3") > 0 And InStr(1, command, "what") > 0 OR InStr(1, command, "4") > 0 And InStr(1, command, "what") > 0 OR InStr(1, command, "5") > 0 And InStr(1, command, "what") > 0 OR InStr(1, command, "6") > 0 And InStr(1, command, "what") > 0 OR InStr(1, command, "7") > 0 And InStr(1, command, "what") > 0 OR InStr(1, command, "8") > 0 And InStr(1, command, "what") > 0 OR InStr(1, command, "9") > 0 And InStr(1, command, "what") > 0 OR InStr(1, command, "0") > 0 And InStr(1, command, "substract") > 0 OR InStr(1, command, "1") > 0 And InStr(1, command, "substract") > 0 OR InStr(1, command, "2") > 0 And InStr(1, command, "substract") > 0 OR InStr(1, command, "3") > 0 And InStr(1, command, "substract") > 0 OR InStr(1, command, "4") > 0 And InStr(1, command, "substract") > 0 OR InStr(1, command, "5") > 0 And InStr(1, command, "substract") > 0 OR InStr(1, command, "6") > 0 And InStr(1, command, "substract") > 0 OR InStr(1, command, "7") > 0 And InStr(1, command, "substract") > 0 OR InStr(1, command, "8") > 0 And InStr(1, command, "substract") > 0 OR InStr(1, command, "9") > 0 And InStr(1, command, "substract") > 0 OR InStr(1, command, "0") > 0 And InStr(1, command, "multiply") > 0 OR InStr(1, command, "1") > 0 And InStr(1, command, "multiply") > 0 OR InStr(1, command, "2") > 0 And InStr(1, command, "multiply") > 0 OR InStr(1, command, "3") > 0 And InStr(1, command, "multiply") > 0 OR InStr(1, command, "4") > 0 And InStr(1, command, "multiply") > 0 OR InStr(1, command, "5") > 0 And InStr(1, command, "multiply") > 0 OR InStr(1, command, "6") > 0 And InStr(1, command, "multiply") > 0 OR InStr(1, command, "7") > 0 And InStr(1, command, "multiply") > 0 OR InStr(1, command, "8") > 0 And InStr(1, command, "multiply") > 0 OR InStr(1, command, "9") > 0 And InStr(1, command, "multiply") > 0 OR InStr(1, command, "0") > 0 And InStr(1, command, "divide") > 0 OR InStr(1, command, "1") > 0 And InStr(1, command, "divide") > 0 OR InStr(1, command, "2") > 0 And InStr(1, command, "divide") > 0 OR InStr(1, command, "3") > 0 And InStr(1, command, "divide") > 0 OR InStr(1, command, "4") > 0 And InStr(1, command, "divide") > 0 OR InStr(1, command, "5") > 0 And InStr(1, command, "divide") > 0 OR InStr(1, command, "6") > 0 And InStr(1, command, "divide") > 0 OR InStr(1, command, "7") > 0 And InStr(1, command, "divide") > 0 OR InStr(1, command, "8") > 0 And InStr(1, command, "divide") > 0 OR InStr(1, command, "9") > 0 And InStr(1, command, "divide") > 0 Then
		commandRefined = Trim(command)
		commandRefined = LCase(commandRefined)
		stringChars = Array()
		commandRefined = Replace(commandRefined, "sqr", "143^.4^..*")
		commandRefined = Replace(commandRefined, "sqrt", "143^.4^..*")
		commandRefined = Replace(commandRefined, "sin", "123^.4^..*")
		commandRefined = Replace(commandRefined, "cos", "133^.4^..*")
		commandRefined = Replace(commandRefined, "tan", "153^.4^..*")
		commandRefined = Replace(commandRefined, "sec", "163^.4^..*")
		commandRefined = Replace(commandRefined, "cosec", "173^.4^..*")
		commandRefined = Replace(commandRefined, "cotan", "183^.4^..*")
		commandRefined = Replace(commandRefined, "arcsin", "193^.4^..*")
		commandRefined = Replace(commandRefined, "arccos", "1103^.4^..*")
		commandRefined = Replace(commandRefined, "arctan", "1113^.4^..*")
		commandRefined = Replace(commandRefined, "asin", "1123^.4^..*")
		commandRefined = Replace(commandRefined, "acos", "1133^.4^..*")
		commandRefined = Replace(commandRefined, "atan", "1143^.4^..*")
		commandRefined = Replace(commandRefined, "atan", "1143^.4^..*")
		commandRefined = Replace(commandRefined, "abs", "1153^.4^..*")
		commandRefined = Replace(commandRefined, "times", "*")
		commandRefined = Replace(commandRefined, "multipl", "*")
		commandRefined = Replace(commandRefined, "divide", "/")
		commandRefined = Replace(commandRefined, "square root", "143^.4^..*")
		commandRefined = Replace(commandRefined, "add", "+")
		commandRefined = Replace(commandRefined, "plus", "+")
		commandRefined = Replace(commandRefined, "minus", "-")
		commandRefined = Replace(commandRefined, "substract", "-")
		commandRefined = Replace(commandRefined, "power", "^")
		For i=1 To Len(commandRefined)
			letterToAdd = Mid(commandRefined,i,1)
			If letterToAdd = "1" OR letterToAdd = "2" OR letterToAdd = "3" OR letterToAdd = "4" OR letterToAdd = "5" OR letterToAdd = "6" OR letterToAdd = "7" OR letterToAdd = "8" OR letterToAdd = "9" OR letterToAdd = "0" OR letterToAdd = "+" OR letterToAdd = "-" OR letterToAdd = "*" OR letterToAdd = "/" OR letterToAdd = "^" OR letterToAdd = "." OR letterToAdd = "(" OR letterToAdd = ")" Then
				ReDim Preserve stringChars(UBound(stringChars) + 1)
				stringChars(UBound(stringChars)) = letterToAdd
			Else
				wscript.sleep(1)
			End If
		Next
		commandRefined = Join(stringChars)
		commandRefined = Replace(commandRefined, " ", "")
		commandRefined = Replace(commandRefined, "143^.4^..*", "sqr")
		commandRefined = Replace(commandRefined, "123^.4^..*", "sin")
		commandRefined = Replace(commandRefined, "133^.4^..*", "cos")
		commandRefined = Replace(commandRefined, "153^.4^..*", "tan")
		commandRefined = Replace(commandRefined, "163^.4^..*", "sec")
		commandRefined = Replace(commandRefined, "173^.4^..*", "cosec")
		commandRefined = Replace(commandRefined, "183^.4^..*", "cotan")
		commandRefined = Replace(commandRefined, "193^.4^..*", "arcsin")
		commandRefined = Replace(commandRefined, "1103^.4^..*", "arccos")
		commandRefined = Replace(commandRefined, "1113^.4^..*", "arctan")
		commandRefined = Replace(commandRefined, "1123^.4^..*", "asin")
		commandRefined = Replace(commandRefined, "1133^.4^..*", "acos")
		commandRefined = Replace(commandRefined, "1143^.4^..*", "atn")
		commandRefined = Replace(commandRefined, "1153^.4^..*", "abs")
		iResult = Eval(commandRefined)
		calculateAnother=0

		If InStr(1, command, "=") > 0 Then
		calculationError=msgbox("Please enter a mathematical expression, not equation. (An expression is a mathematical phrase without an equal sign, while an expression is a mathematical phrase with an equal sign)", 16, "Error: Please enter a math expression")
		ElseIf iResult="" Then
		calculationError=msgbox("Please enter a valid expression, without any variables, constants, (letters) or any non-mathematical symbols (such as commas)", 16, "Error: Please enter a proper math expression")
		Else
		If iResult < 9999999 Then
		If InStr(1, iResult, ".") > 0 And Len(Mid(iResult, InStr(1, iResult, ".") + 1)) > 5 Then
		calcArrayAround = Array("It's around", "That would be around", "It's around", "The answer's around", "The answer is around")
		Sapi.speak rand(calcArrayAround)
		Sapi.speak CStr(Round(iResult, 5))
		Else
		calcArray = Array("It's ", "That would be ", "It's ", "The answer is equal to ", "The answer is")
		Sapi.speak rand(calcArray)
		Sapi.speak CStr(iResult)
		End If
		Else
		wscript.sleep(1)
		End If
		calculationTitle = CStr("Answer to '" + commandRefined + "'")
		calculationResult=msgbox("'" & CStr(commandRefined) & "' = " & iResult,,CStr(calculationTitle))
		End If



		do while calculateAnother<1
					Sapi.speak "Do you want to do another calculation?"
					result = MsgBox ("Do you want to do another calculation?", vbYesNo, "Calculate again?")
					
					Select Case result
					Case vbYes
						Sapi.speak "Please enter what you want to calculate"
						commandforCalc = inputbox("Please enter your expression (ex. 2*2, or 2^4) in the input box below. (Without the word 'calculate' at the start; simply type in the expression you want to calculate)" & vbcrlf & "" & vbcrlf & "Tip: You can type in 'sqr(x)' to find the square root of x.","Ble Ble Ble...")
						commandRefined=commandforCalc
						commandRefined = Replace(commandRefined, "sqrt", "sqr")
						iResult = Eval(commandRefined)
						
						If InStr(1, command, "=") > 0 Then
						calculationError=msgbox("Please enter a mathematical expression, not equation. (An expression is a mathematical phrase without an equal sign, while an expression is a mathematical phrase with an equal sign)", 16, "Error: Please enter a math expression")
						ElseIf iResult="" Then
						calculationError=msgbox("Please enter a valid expression, without any variables, constants, (letters) or any non-mathematical symbols (such as commas)", 16, "Error: Please enter a proper math expression")
						Else
						If iResult < 9999999 Then
						If InStr(1, iResult, ".") > 0 Then
						Sapi.speak "the answer to your question is around"
						Sapi.speak CStr(Round(iResult, 5))
						Else
						Sapi.speak "the answer to your question is"
						Sapi.speak CStr(iResult)
						End If
						Else
						wscript.sleep(1)
						End If
						calculationTitle = CStr("Answer to '" + commandRefined + "'")
						calculationResult=msgbox("'" & CStr(commandRefined) & "' = " & iResult,,CStr(calculationTitle))
						End If
					Case vbNo
						calculateAnother = 1
					End Select
				loop
	ElseIf command="calculator" OR command="calc" OR command="calc." OR command="open calculator" OR command="open calc" OR command="open calc." OR command="please open calculator" OR command="please open calc" OR command="please open calc." OR command="open calculator please" OR command="open calc please" OR command="open calc. please" OR InStr(1, command, "calculator") > 0 OR InStr(1, command, "calc.") > 0 OR InStr(1, command, "calc") > 0 Then
		Sapi.speak "Opening Calculator"
		oShell.Run "calc"
		WScript.Sleep 10000
		oShell.AppActivate "Calculator"
	ElseIf InStr(1, command, "notepad") > 0 OR InStr(1, command, "note pad") > 0 OR InStr(1, command, "note") > 0 And InStr(1, command, "pad") > 0 Then
		set WshShell=WScript.CreateObject("WScript.Shell")
		Set fsotwo = CreateObject("Scripting.FileSystemObject")
		
		If fsotwo.FileExists("C:\windows\notepad.exe") Then
		WshShell.run "notepad.exe"
		Else
			notepadError=msgbox("Unfortunately, you haven't installed notepad on your computer. :-(", 16, "Notepad is not installed")
		End If
	ElseIf InStr(1, command, "file") > 0 Or InStr(1, command, "explorer") > 0 Then
		Sapi.speak "Opening file explorer"
		Dim SH, txtFolderToOpen 
		Set SH = WScript.CreateObject("WScript.Shell") 
		txtFolderToOpen = "C:\Users\%username%\Documents"
		SH.Run txtFolderToOpen 
		Set SH = Nothing
	ElseIf InStr(1, command, "control") > 0 Then
		Sapi.speak "Opening Control Panel"
		set WshShell=WScript.CreateObject("WScript.Shell")
		WshShell.run "control.exe"
	ElseIf InStr(1, command, "chrome") > 0 And InStr(1, command, "tab") > 0 Then
		Sapi.speak "Opening new chrome tab"
		OpenWithChrome "http://www.google.com"
	ElseIf InStr(1, command, "chrome") > 0 Then
		set WshShell=WScript.CreateObject("WScript.Shell")
		Set fsotwo = CreateObject("Scripting.FileSystemObject")
		
		If fsotwo.FileExists("C:\Program Files (x86)\Google\Chrome\Application\chrome.exe") OR fsotwo.FileExists("&appdata%\Google\Chrome") OR fsotwo.FileExists("%localappdata%\Google\Chrome") OR fsotwo.FileExists("C:\Program Files (x86)\Google\Application\chrome.exe") Then
		Sapi.speak "Opening new chrome window"
		WshShell.run "chrome.exe"
		Else
		notepadError=msgbox("Unfortunately, you haven't installed chrome on your computer. :-(", 16, "Chrome is not installed")
		End If
	ElseIf InStr(1, command, "firefox") > 0 And InStr(1, command, "tab") > 0 Then
		Sapi.speak "Opening new firefox tab"
		CreateObject("WScript.Shell").Run "FireFox google.com", 2
	ElseIf InStr(1, command, "firefox") > 0 Then
		Set fsotwo=CreateObject("Scripting.FileSystemObject")
		If fsotwo.FileExists("C:\Program Files (x86)\Mozilla Firefox\firefox.exe") Then
			Sapi.speak "Opening new firefox window"
		   WshShell.run "firefox.exe"
		Else
			notepadError=msgbox("Unfortunately, you haven't installed firefox on your computer. :-(", 16, "Firefox is not installed")
		End If
	ElseIf command="help" OR InStr(1, command, "help") > 0 OR InStr(1, command, "command") > 0 OR InStr(1, command, "order") > 0 OR InStr(1, command, "list") > 0 Then
		Sapi.speak "Here is a list of all commands so far."
		help=msgbox("All current commands:" & vbcrlf & "" & vbcrlf & "- 'open www.[url].com' = Opens the specified url in your default browser." & vbcrlf & "- 'create bookmark' = Shortcut to go to a certain url or to execute/run/open a file. Instead of always typing in 'open www.google.com', (for example) you can create a bookmark called 'google', so that every time you type in the command 'google' in the command/input box, it will open 'www.google.com' in your default browser. Similarly, you can create a bookmark called 'homework', which can open the file 'homework.txt'. (For example...)" & vbcrlf & "- 'list bookmarks' = Command that lists all of your present bookmarks" & vbcrlf & "- 'remove bookmark' = A command that removes a bookmark" & vbcrlf & "- 'date' = Tells you the current day, month and year" & vbcrlf & "- 'time' = Tells you the current time" & vbcrlf & "- 'date and time' = Tells you the date and time" & vbcrlf & "- 'help' = Tells you all the current commands" & vbcrlf & "- 'assignments' = Opens a new tab listing all the current assignments" & vbcrlf & "" & vbcrlf & "CLICK OK TO SEE MORE COMMANDS",,"All commands")
		helpTwo=msgbox("MORE COMMANDS:" & vbcrlf & "" & vbcrlf & "- 'events'/'special events' = Opens a new tab containing a list of all the upcoming school events." & vbcrlf & "- 'schedule' = Opens up a tab that displays our class' schedule" & vbcrlf & "- 'joke' = Command that says a random (funny) joke." & vbcrlf & "- 'search for ""[enter what you want to search here]"" on google' = Command that googles what's in the quotation marks" & vbcrlf & "- 'rename' = Command that allows you to change your name." & vbcrlf & "- 'voice recognition' = Command that guides you through the setup to activate voice recognition. This feature of the Sherby Interface allows you to tell Sherby what you want to do VERBALLY, instead of writing it." & vbcrlf & "- 'weather' = Gives you the current weather in Ottawa" & vbcrlf & "- 'volume up' = Turns up the volume." & vbcrlf & "- 'volume down' = Turns the volume lower" & vbcrlf & "- 'mute' = Toggles the mute setting. (ex. If the volume is muted, it will unmute it. If the volume is unmuted, it will mute it)" & vbcrlf & "" & vbcrlf & "CLICK OK TO SEE MORE COMMANDS",,"All commands")
		helpTwoAndAHalf=msgbox("MORE COMMANDS:" & vbcrlf & vbcrlf & "- 'create note' = Command that allows you create a 'note'. A note is similar to a goal, it's like a message, where you can write anything you want. This command is particularly useful if you need to remember something on the spot. (ex. 'the capital of Mongolia is Ulaanbaatar')" & vbcrlf & "- 'list notes' = Command that lists all of your current notes" & vbcrlf & "- 'delete note' = Command that deletes a specified note. To delete a note, you need to specify the index of the note. For example, (after viewing all your current notes) if you wanted to delete your 3rd note, you would simply have to enter the number '3' to delete the third note" & vbcrlf & "- 'news' = Command that displays the current news. (news comes from BBC)" & vbcrlf & "- 'open news' = Command that opens cnn.com and displays the current news" & vbcrlf & "- 'open bbc news' = Command that opens up bbc news in a new tab" & vbcrlf & "- 'open yahoo news' = Command that opens yahoo news in a new tab" & vbcrlf & "" & vbcrlf & "CLICK OK TO SEE MORE COMMANDS",,"All commands")
		helpThree=msgbox("MORE COMMANDS:" & vbcrlf & "" & vbcrlf & "- 'email' = Command that allows you to email anyone anything from your Gmail account. (Note: You need to be signed in for this function to work...)" & vbCrLf & "- 'nevermind'/'cancel' = Command that ends the script" & vbCrLf & "- 'search wikipedia'/'wikipedia' = Function that allows you to search anything on Wikipedia. (TIP: This function is available in french, spanish, german, and chinese, simply type in 'french wikipedia' to search a french article, or 'chinese wikipedia' to search a chinese article.)" & vbCrLf & "- 'create goal' = This command allows you to make a 'todo' list of all the things you need to do. For example, you can enter 'Finish Geo Project' as a goal, so that you don't forget to finish it. You can then check/view all of your goals to see what else you need to do" & vbCrLf & "- 'goals'/'view goals' = Command that shows you all your current goals. You can check this anytime to see what you still have to do." & vbcrlf & "" & vbcrlf & "CLICK OK TO SEE MORE COMMANDS",,"All commands")
		helpFour=msgbox("MORE COMMANDS:" & vbcrlf & "" & vbcrlf & "- 'mark goal as done'/'delete goal' = Command that allows you to remove a goal from the 'goal list' once you've finished it, no longer need to do it, or for any other reason. IMPORTANT: The input must be an integer indicating the 'index of the goal'. In other words, (after viewing all your current goals using the 'view goals' command) if you want to delete or mark your 3rd goal as done, you simply need to type in the number '3'." & vbcrlf & "- 'version' = Command that tells you the current version of Sherby." & vbcrlf & "- 'disable voice recognition' = Command that disables voice recognition. You can easily re-enable it by typing in the command 'enable voice recognition'" & vbcrlf & "- 'shutdown' = Command that shuts down your computer" & vbcrlf & "- 'restart'/'reboot' = Command that restarts/reboots your computer" & vbcrlf & "- 'notepad' = Command that opens a new notepad document" & vbcrlf & "- 'file explorer' = Command that opens the file explorer to the Documents folder" & vbcrlf & "" & vbcrlf & "CLICK OK TO SEE MORE COMMANDS",,"All commands")
		helpFive=msgbox("MORE COMMANDS:" & vbcrlf & "" & vbcrlf & "- 'chrome' = Command that opens a new chrome window" & vbcrlf & "- 'chrome tab' = Command that opens a new chrome tab in an existing window. If there is no current window open, it will open a new tab in a new chrome window" & vbcrlf & "- 'firefox' = Opens new firefox window" & vbcrlf & "- 'firefox tab' = Command that opens a firefox tab in an existing window. It will open a tab in a new window if there is no existing firefox window open" & vbcrlf & "- 'word' = Creates a new Microsoft Word document" & vbcrlf & "- 'excel' = Creates a new Microsoft Excel document" & vbcrlf & "- 'powerpoint' = Creates a new Microsoft Powerpoint document" & vbcrlf & "- 'flip a coin' = Command that returns either heads or tails" & vbcrlf & "- 'roll a die' = Command that simulates rolling a six sided die" & vbcrlf & "- 'roll a [number of sides] sided die' = Command that simulates rolling a die with your specified number of sides" & vbcrlf & "- 'yes or no' = Command that returns either 'yes' or 'no'" & vbcrlf & "" & vbcrlf & "CLICK OK TO SEE MORE COMMANDS",,"All commands")
		helpSix=msgbox("MORE COMMANDS:" & vbcrlf & "" & vbcrlf & "- 'convert' [unit to another unit] = Command that, as long as you have the word 'convert' will convert almost any unit into any other unit, including currency, temperature, measurement, etc. (ex. Convert 53.4 kilometers to miles)" & vbcrlf & "- videos of ""[thing you want to search]"" = Command that searches for the text in the quotation marks on Youtube. (ex. Show me videos of ""Black bears"". Also, it only works with double quotation marks.)" & vbcrlf & "- images of ""[text you want to search for]"" = Command that searches for your specified text (found in the quotation marks) in Google Images. (ex. Find images of ""waterfalls"". Note that this function only works with double quotation marks...)" & vbcrlf & "- search the web for ""[text you want to search]"" = Command that searches for the text in the quotation marks in Google. (ex. Search for ""Australia population"" on the web. Please note that this command only works with double quotation marks...)" & vbcrlf & "- 'google search' = Command that prompts/asks you what you want to search on google and then searches it" & vbcrlf & "" & vbcrlf & "CLICK OK TO SEE MORE COMMANDS",,"All commands")
		helpSeven=msgbox("MORE COMMANDS:" & vbcrlf & "" & vbcrlf & "- 'google image search' = Command that asks you what you want to search on Google Images. Sherby then searches your desired query in Google Images" & vbcrlf & "- 'youtube search' = Command that asks you and then displays what you want to search for in Youtube" & vbcrlf & "- 'delete all bookmarks' = Command that deletes all your bookmarks" & vbcrlf & "- 'delete all goals' = Command that deletes all your current goals" & vbcrlf & "- 'delete all notes' = Command that deletes all of your notes" & vbcrlf & "- calculate [calculation] = Command that allows you calculate anything you want. (Note: the operation factorial doesn't work, but other complex operations such as square root or trigonometric operations do work. Also, no quotation/double quotation marks are needed)" & vbcrlf & "- 'calculator' = Opens the calculator app for you",,"All commands")
		Sapi.speak "However, aside from these commands, you can also ask Sherby real questions, such as, when did the American Revolution Start, and much more!"
	ElseIf InStr(1, command, "news") > 0 And InStr(1, command, "view") > 0 And InStr(1, command, "bbc") OR InStr(1, command, "news") > 0 And InStr(1, command, "open") > 0 And InStr(1, command, "bbc") Then
		a.run "http://www.bbc.com/news/world"
	ElseIf InStr(1, command, "news") > 0 And InStr(1, command, "view") > 0 And InStr(1, command, "yahoo") OR InStr(1, command, "news") > 0 And InStr(1, command, "open") > 0 And InStr(1, command, "yahoo") Then
		a.run "https://www.yahoo.com/news/"
	ElseIf InStr(1, command, "news") > 0 And InStr(1, command, "view") > 0 OR InStr(1, command, "news") > 0 And InStr(1, command, "open") > 0 Then
		a.run "http://www.cnn.com"
	ElseIf InStr(1, command, "news") > 0 Then
		Set req = CreateObject("MSXML2.XMLHTTP.3.0")
		req.Open "GET", "http://feeds.bbci.co.uk/news/rss.xml", False
		req.Send
		Set xmlNews = CreateObject("Msxml2.DOMDocument")
		xmlNews.loadXml(req.responseText)
		newsResult = msgbox("-" & UCase(xmlNews.getElementsByTagName("channel/item/title")(0).Text) & vbcrlf & xmlNews.getElementsByTagName("channel/item/description")(0).Text & vbcrlf & vbcrlf & "-" & UCase(xmlNews.getElementsByTagName("channel/item/title")(1).Text) & vbcrlf & xmlNews.getElementsByTagName("channel/item/description")(1).Text & vbcrlf & vbcrlf & "-" & UCase(xmlNews.getElementsByTagName("channel/item/title")(2).Text) & vbcrlf & xmlNews.getElementsByTagName("channel/item/description")(2).Text,,"Current News")
	ElseIf InStr(1, command, "email") > 0 OR InStr(1, command, "message") > 0 OR InStr(1, command, "text") > 0 Then
		Sapi.speak "Please specify the email address of the person you want to receive your email"
		toRecipient=inputbox("Who do you want to email your message to?" & vbcrlf & "" & vbcrlf & "TIP: If you want to email a message to multiple people, simply separate their emails with a space.","Receiver")
		Sapi.speak "What is the title of your message?"
		title=inputbox("What is the title of your message?","Title")
		messageMessage="What is the message you want to send to '" + toRecipient + "'?"
		Sapi.speak "Please type in your message"
		messageForEmail=inputbox(messageMessage,"Message")
		a.run "https://mail.google.com/mail/ca/u/0/#inbox?compose=new"
		wscript.sleep(10000)
		a.sendkeys "{Tab}"
		a.sendkeys(toRecipient)
		a.sendkeys "{Enter}"
		a.sendkeys "{Tab}"
		a.sendkeys(title)
		a.sendkeys "{Tab}"
		a.sendkeys(messageForEmail)
		a.sendkeys "{Tab}"
		a.sendkeys "{Tab}"
	ElseIf command="nevermind" OR command="no" OR command="nah" OR command="nah gee" OR command="na gee" OR command="never mind" OR command="nothing" OR InStr(1, command, "nevermind") > 0 OR InStr(1, command, "never mind") > 0 OR InStr(1, command, "cancel") > 0 OR InStr(1, command, "bye") > 0 OR InStr(1, command, "whatever") > 0 Or InStr(1, command, "bye") > 0 OR InStr(1, command, "night") > 0 OR InStr(1, command, "see") > 0 And InStr(1, command, "later") > 0 OR InStr(1, command, "toodloo") > 0 OR InStr(1, command, "off") > 0 OR InStr(1, command, "shoo") > 0 OR InStr(1, command, "go") > 0 And InStr(1, command, "away") > 0 OR InStr(1, command, "see") > 0 And InStr(1, command, "later") > 0 Then
		Call Play("C:\Windows\Media\notify.wav")
		currentHour = Hour(Now())
		If currentHour >= 1 And currentHour < 12 Then
			byeArrayMorning = Array("Okay then, goodbye ", "Okay, goodbye ", "Well I guess that wraps it up. See you later ", "Okay then. until next time ", "I guess I'll be going then. Bye ", "Okay, see you later. Have a great day ", "until next time ", "Nice talking to you! Have a great day ", "Okay, I better be on my way. See you later ")
			msg = rand(byeArrayMorning) + CStr(masterName)
		ElseIf currentHour >= 12 And currentHour <18 Then
			byeArrayMorning = Array("Okay then, goodbye ", "Okay, goodbye ", "Well I guess that wraps it up. See you later ", "Okay then. until next time ", "I guess I'll be going then. Bye ", "Okay, see you later. Have a great rest of the day ", "until next time ", "Nice talking to you! Have a great rest of the day ", "Anyways, I think it's time to go. Have a great day ")
			msg = rand(byeArrayMorning) + CStr(masterName)
		ElseIf currentHour >=18 And currentHour<21 Then
			byeArrayMorning = Array("Okay then, goodbye ", "Okay, goodbye ", "Well I guess that wraps it up. See you later ", "Okay then. until next time ", "I guess I'll be going then. Bye ", "Okay, see you later. Have a great evening ", "until next time ", "Nice talking to you! Have a great evening ")
			msg = rand(byeArrayMorning) + CStr(masterName)
		ElseIf currentHour >= 21 Then
			byeArrayMorning = Array("Okay then, goodnight! Oh, and don't let the bed bugs bite! ", "Okay, goodbye and sleep tight! ", "Well I guess that wraps it up. Sweet dreams!", "Okay then. See you later, and don't let the bed bugs bite! ", "I guess I'll be going then. Sleep tight ", "Okay, see you later. Have a great night and don't let the bed bugs bite ", "until next time, ", "Nice talking to you! Good night!")
			msg = rand(byeArrayMorning)
		End If
		Set greeter = CreateObject("sapi.spvoice")
		greeter.Speak msg
		wscript.quit
	ElseIf InStr(1, command, "assignment") > 0 Then
		If bolActiveConnection = False Then
		wifiErrorForOpeningAssignments = MsgBox ("Your computer is not connected to the internet. This functionality will not work without internet connection." & vbCrLf & "" & vbCrLf & "Do you want to continue?", vbYesNo, "No Internet Connection")
		
		Select Case wifiErrorForOpeningAssignments
		Case vbYes
			Sapi.speak "Here is a list of all the assignments so far"
			a.run "http://broadview8.weebly.com/assignments.html"
		Case vbNo
			wscript.sleep(1)
		End Select
		Else
			Sapi.speak "Here is a list of all the assignments so far"
			a.run "http://broadview8.weebly.com/assignments.html"
		End If
	ElseIf InStr(1, command, "event") > 0 Then
		If bolActiveConnection = False Then
		wifiErrorForOpeningEvents = MsgBox ("Your computer is not connected to the internet. This functionality will not work without internet connection." & vbCrLf & "" & vbCrLf & "Do you want to continue?", vbYesNo, "No Internet Connection")
		
		Select Case wifiErrorForOpeningEvents
		Case vbYes
			Sapi.speak "Here is a list of all the important school events coming up"
			a.run "http://broadview8.weebly.com/special-events.html"
		Case vbNo
			wscript.sleep(1)
		End Select
		Else
			Sapi.speak "Here is a list of all the important school events coming up"
			a.run "http://broadview8.weebly.com/special-events.html"
		End If
	ElseIf InStr(1, command, "schedule") > 0 Then
		If bolActiveConnection = False Then
		wifiErrorForOpeningSchedule = MsgBox ("Your computer is not connected to the internet. This functionality will not work without internet connection." & vbCrLf & "" & vbCrLf & "Do you want to continue?", vbYesNo, "No Internet Connection")
		
		Select Case wifiErrorForOpeningSchedule
		Case vbYes
			Sapi.speak "Here is our class schedule for the 2016 and 2017 school year"
			a.run "http://broadview8.weebly.com/schedule.html"
		Case vbNo
			wscript.sleep(1)
		End Select
		Else
			Sapi.speak "Here is our class schedule for the 2016 and 2017 school year"
			a.run "http://broadview8.weebly.com/schedule.html"
		End If
	ElseIf InStr(1, command, "conver") > 0 Then
		doWolframConvertSearch(command)
	ElseIf InStr(1, command, "cel") > 0 And InStr(1, command, "fah") > 0 Then
		'Convert celsius to fahrenheit
		If InStr(1, command, "cel") < InStr(1, command, "fah") Then
			commandRefined = Trim(command)
			stringChars = Array()
			For i=1 To Len(commandRefined)
				letterToAdd = Mid(commandRefined,i,1)
				If letterToAdd = "1" OR letterToAdd = "2" OR letterToAdd = "3" OR letterToAdd = "4" OR letterToAdd = "5" OR letterToAdd = "6" OR letterToAdd = "7" OR letterToAdd = "8" OR letterToAdd = "9" OR letterToAdd = "0" OR InStr(1, command, ".") > 0 Then
					ReDim Preserve stringChars(UBound(stringChars) + 1)
					stringChars(UBound(stringChars)) = letterToAdd
				Else
					wscript.sleep(1)
				End If
			Next
			commandRefined = Join(stringChars)
			celToFah = commandRefined
			celToFah = Replace(celToFah, " ", "")
			convertAnotherCelToFah = 0
			If isNumeric(celToFah) Then
				celToFah = CDbl(celToFah)
				convertedToFah = celToFah * (9/5) + 32
				If InStr(1, convertedToFah, ".") > 0 Then
					Sapi.speak CStr(celToFah) & "degrees celsius is around" & CStr(Round(convertedToFah, 5)) & "degrees fahrenheit"
				Else
					Sapi.speak CStr(celToFah) & "degrees celsius is" & CStr(Round(convertedToFah, 3)) & "degrees fahrenheit"
				End If
				celToFahResult = msgbox(CStr(celToFah) & Chr(176) & "C = " & CStr(convertedToFah) & Chr(176) & "F",,"Celsius to Fahrenheit")
			Else
				conversionError = msgbox("Please enter a number only, without a degree sign, or any other text", 16, "Conversion error")
			End If
			
			do while convertAnotherCelToFah<1
					Sapi.speak "Would you like to do another celsius to fahrenheit conversion?"
					result = MsgBox ("Do you want to do another celsius to fahrenheit conversion?", vbYesNo, "Another conversion?")
					
					Select Case result
					Case vbYes
						Sapi.speak "What would you like to convert from celsius to fahrenheit"
						celToFah = inputbox("What would you like to convert to fahrenheit? (From celsius)" & vbcrlf & "" & vbcrlf & "Simply enter the number, without the degree sign. (ex. '20')","Celsius to Fahrenheit")
						If isNumeric(celToFah) Then
							celToFah = CDbl(celToFah)
							convertedToFah = celToFah * (9/5) + 32
							If InStr(1, convertedToFah, ".") > 0 Then
								Sapi.speak CStr(celToFah) & "degrees celsius is around" & CStr(Round(convertedToFah, 5)) & "degrees fahrenheit"
							Else
								Sapi.speak CStr(celToFah) & "degrees celsius is" & CStr(Round(convertedToFah, 3)) & "degrees fahrenheit"
							End If
							celToFahResult = msgbox(CStr(celToFah) & Chr(176) & "C = " & convertedToFah & Chr(176) & "F",,"Celsius to Fahrenheit")
						Else
							conversionError = msgbox("Please enter a number only, without a degree sign, or any other text", 16, "Conversion error")
						End If
					Case vbNo
						convertAnotherCelToFah = 1
					End Select
				loop
			
			
		'Convert fahrenheit to celsius
		ElseIf InStr(1, command, "fah") < InStr(1, command, "cel") Then
			commandRefined = Trim(command)
			stringChars = Array()
			For i=1 To Len(commandRefined)
				letterToAdd = Mid(commandRefined,i,1)
				If letterToAdd = "1" OR letterToAdd = "2" OR letterToAdd = "3" OR letterToAdd = "4" OR letterToAdd = "5" OR letterToAdd = "6" OR letterToAdd = "7" OR letterToAdd = "8" OR letterToAdd = "9" OR letterToAdd = "0" OR InStr(1, command, ".") > 0 Then
					ReDim Preserve stringChars(UBound(stringChars) + 1)
					stringChars(UBound(stringChars)) = letterToAdd
				Else
					wscript.sleep(1)
				End If
			Next
			commandRefined = Join(stringChars)
			fahToCel = commandRefined
			fahToCel = Replace(fahToCel, " ", "")
			convertAnotherFahToCel = 0
			If isNumeric(fahToCel) Then
				fahToCel = CDbl(fahToCel)
				convertedToCel = fahToCel * (9/5) + 32
				fahToCel = CInt(fahToCel)
				convertedToCel = CInt(convertedToCel)
				If InStr(1, convertedToCel, ".") > 0 Then
					Sapi.speak CStr(fahToCel) & "degrees fahrenheit is around" & CStr(CInt(Round(convertedToCel, 3))) & "degrees celsius"
				Else
					Sapi.speak CStr(fahToCel) & "degrees fahrenheit is" & CStr(CInt(convertedToCel)) & "degrees celsius"
				End If
				fahToCelResult = msgbox(CStr(fahToCel) & Chr(176) & "F = " & convertedToCel & Chr(176) & "C",,"Fahrenheit to Celsius")
			Else
				conversionError = msgbox("Please enter a number only, without a degree sign, or any other text", 16, "Conversion error")
			End If
			
			do while convertAnotherFahToCel<1
					Sapi.speak "Would you like to do another fahrenheit to celsius conversion?"
					result = MsgBox ("Do you want to do another fahrenheit to celsius conversion?", vbYesNo, "Another conversion?")
					
					Select Case result
					Case vbYes
						Sapi.speak "What would you like to convert from fahrenheit to celsius"
						fahToCel = inputbox("What would you like to convert to celsius? (From Fahrenheit)" & vbcrlf & "" & vbcrlf & "Simply enter the number, without the degree sign. (ex. '20')","Fahrenheit tp Celsius")
						If isNumeric(fahToCel) Then
							fahToCel = CDbl(fahToCel)
							convertedToCel = fahToCel * (9/5) + 32
							fahToCel = CInt(fahToCel)
							convertedToCel = CInt(convertedToCel)
							If InStr(1, convertedToCel, ".") > 0 Then
								Sapi.speak CStr(CInt(fahToCel)) & "degrees fahrenheit is around" & CStr(CInt(Round(convertedToCel, 3))) & "degrees celsius"
							Else
								Sapi.speak CStr(CInt(fahToCel)) & "degrees fahrenheit is" & CStr(CInt(convertedToCel)) & "degrees celsius"
							End If
						Else
							conversionError = msgbox("Please enter a number only, without a degree sign, or any other text", 16, "Conversion error")
						End If
						fahToCelResult = msgbox(CStr(fahToCel) & Chr(176) & "F = " & convertedToCel & Chr(176) & "C",,"Fahrenheit to Celsius")
					Case vbNo
						convertAnotherFahToCel = 1
					End Select
				loop
		End If
	ElseIf InStr(1, command, "calendar") > 0 Then
		If bolActiveConnection = False Then
		wifiErrorForOpeningCalendar = MsgBox ("Your computer is not connected to the internet. This functionality will not work without internet connection." & vbCrLf & "" & vbCrLf & "Do you want to continue?", vbYesNo, "No Internet Connection")
		
		Select Case wifiErrorForOpeningCalendar
		Case vbYes
			Sapi.speak "Here is a calendar showing all special events, assignments, due dates, holidays, and more!"
			a.run "http://broadview8.weebly.com/calendar.html"
		Case vbNo
			wscript.sleep(1)
		End Select
		Else
			Sapi.speak "Here is a calendar showing all special events, assignments, due dates, holidays, and more!"
			a.run "http://broadview8.weebly.com/calendar.html"
		End If
	ElseIf command="kill yourself" OR command="go kill yourself" Then
		sapi.speak "Ouch, my feelings. I am now killing myself"
		wscript.sleep(3000)
		Sapi.speak "Shutting down, see you later alligator"
		wscript.quit
	ElseIf command="creator" OR InStr(1, command, "creator") > 0 OR InStr(1, command, "created") > 0 OR InStr(1, command, "create") > 0 Then
		Sapi.speak "The creator of the Sherby interface is Albert Neeetooo, a 13 year old student. He created this interface in around one and a half months and looks forwards to expanding it and improving it as much as possible."
	ElseIf InStr(1, command, "sing") > 0 OR InStr(1, command, " song ") > 0 Or InStr(1, command, " songs ") Then
		sing = Array("one", "two", "three")
		randomSing = rand(sing)
		Sapi.speak "One song comin right up!"
		If sing = "one" Then
			a.run "‪C:\Windows\Media\flourish.mid"
		ElseIf sing = "two" Then
			a.run "‪C:\Windows\Media\onestop.mid"
		Else
			a.run "‪C:\Windows\Media\town.mid"
		End If
	ElseIf InStr(1, command, "story") > 0 Then
		storyRandom = Array("It was a dark and stromy night... No, that's not it!", "Once apon a time, there was a little fairy... No, too boring!", "In a galaxy far far far... Nah, too far fetched.", "There was once a nice little fellow called Sherby, and he loved to help people. One day, he wanted to help an elderly man cross the road, the man said yes. Sherby was very happy, and while walking in the park, he accidentally bumped into a person") ' who got really mad at him. Suddenly, the elderly man he had helped earlier came out of nowhere and shooed the mean person away. From that day, Sherby learned a valuable lesson, acts of kindess always repay themselves!
		randomStory = rand(storyRandom)
		If Left(randomStory, 10) = "There was once a nice little" Then
			Sapi.speak randomStory
			Sapi.speak " who got really mad at him. Suddenly, the elderly man he had helped earlier came out of nowhere and shooed the mean person away. From that day, Sherby learned a valuable lesson, acts of kindess always repay themselves!"
		Else
			Sapi.speak randomStory
		End If
	ElseIf InStr(1, command, "poem") > 0 Then
		poemArray = Array("Roses are red, violets are blue, haven't you got anything better to do!", "Oh freddled gruntbuggly, Thy micturations are to me, As plurdled gabbleblotchits, On a lurgid... Oh whatever, I give up!!!", "I'd rather not, you wouldn;t like it anyways...", "Are you sure?")
		Sapi.speak rand(poemArray)
	ElseIf command="joke" OR command="tell me a joke" OR InStr(1, command, "joke") > 0 OR InStr(1, command, "pun") > 0 Then
		jokes=Array("When you're cold, go in a corner, they're usually around 90 degrees", "Algebra for the romans was never fun, because X was always ten", "3 out of 2 people say that they have trouble with fractions", "Never argue with a corner, it's always right", "John has 32 candies. He eats 28. What does he have now?,Answer, diabetes", "Why are frogs always happy?,Because they eat what bugs them!!!", "What did the farmer use to make perfect crop circles?,A protractor!", "Whst type of shorts do clouds wear?,Thunderwear!", "What do you call a lazy kangaroo?,A Pouch Potato!", "What did the grape say when it got pressed?,Nothing, it just let out a little wine!", "Why was the frog waiting for the bus?,Because his car got towed!!!", "I wasn't originally going to get a brain transplant, but then I changed my mind.", "I'd tell you a chemistry joke, but I know I wouldn't get a reaction.", "I'm reading a book about anti-gravity. It's impossible to put down.", "Did you hear about the guy who got hit in the head with a can of soda? He was lucky it was a soft drink.", "Why did the scientist install a knocker on his door? He wanted to win the No, bell prize!", "I am on a seafood diet. Every time I see food, I eat it.", "A book just fell on my head. I've only got myshelf to blame.", "What do prisoners use to call each other? Cell phones.", "My math teacher called me average. How mean!", "I would tell you a chemistry joke, but all the good ones argon", "How do you organize a space party? You planet!", "Why didn't dracula attack taylor swift? because she had bad blood!", "Two antennae were on a roof. They fell in love and got married. The service wasn't great but the reception was amazing.", "What do you do when a chemist dies? You barium.", "The shovel was a truly ground breaking invention", "The invention of soap washed the competition away", "The invention of the shovel swept the nation", "I don't think I need a spine. It's holding me back", "They're finally making a movie called clocks. It's about time", "Why do crabs never give money to charity? Because they're shellfish", "The majority of people find bananas a peeling", "Organ donors put their heart into it.", "How do construction workers party? They raise the roof!", "Why are Teedy Bears always hungry? because they're always stuffed!", "Humpty Dumpty had a horrible summer, but he had a great fall!", "Why did the picture go to jail? Because it was framed!", "How does a farmer count cows? With a cow-culator!", "Being friends with assassins is a bad idea. They're all backstabbers", "Why was the baseball field cool after the game?, Because all the fans left!", "A cartoonist was found dead in home. The details are sketchy.", "I went to buy some camouflage clothes the other day, but I couldn't find any.", "What kind of shoes do ninjas wear?, Sneakers.", "I was trying to make a pun about escaping quicksand, but I'm stuck.", "Sleeping is so easy that I can even do it with my eyes closed!", "I used to be a banker, but I lost interest", "Did you know that a recent survey has found that people with more birthdays live longer?", "What did the ocean say to the beach?, Nothing it just waved!", "I don't trust stairs... they are always up to something!", "What do you call an alligator with a vest? An investigator!", "I used to be a doctor, but then I lost patients!", "A blind man walked into a bar. Then a table, then a chair.", "What did the triangle say to the circle? You're so pointless.", "Why didn't the spider go to school? Because she learned everything on the web")
		anotherJoke = 0
		max=UBound(jokes)
		min=0
		Randomize
		randomnumber=Int((max-min+1)*Rnd+min)
		randomjoke=jokes(randomnumber)
		Sapi.speak "Here's a funny one,"
		wscript.sleep(500)
		Sapi.speak randomjoke
		wscript.sleep(1000)
		Sapi.speak "ha, ha, ha, that's the sealiest thing I've ever heard!"
		wscript.sleep(500)
		do while anotherJoke<1
			Sapi.speak "Anyways, do you want to hear another joke?"
			result = MsgBox ("Do you want to hear another 'hilarious' joke?", vbYesNo, "Another one? (bites the dust)")
			
			Select Case result
			Case vbYes
				Randomize
				randomnumber=Int((max-min+1)*Rnd+min)
				randomjoke=jokes(randomnumber)
				Sapi.speak "Here's a funny one,"
				wscript.sleep(500)
				Sapi.speak randomjoke
				wscript.sleep(800)
				Sapi.speak "ha, ha, ha, that's the sealiest thing I've ever heard!"
				wscript.sleep(500)
			Case vbNo
				anotherJoke = 1
			End Select
		loop
		wscript.sleep(1000)
		Sapi.speak "Aww, I was beginning to enjoy myself!"
		wscript.sleep(500)
		Sapi.speak "Anyways"
	ElseIf InStr(1, command, "volume") > 0 And InStr(1, command, "up") > 0 OR InStr(1, command, "volume") > 0 And InStr(1, command, "high") > 0 OR InStr(1, command, "volume") > 0 And InStr(1, command, "increase") > 0 OR InStr(1, command, "volume") > 0 And InStr(1, command, "loud") > 0 Then
		continueVolumeUp = 0
		
		for i=0 to 4
		WshShell.SendKeys(chr(175))
		next
		wscript.sleep(500)
		do while continueVolumeUp<1
			result = MsgBox ("Do you want to increase the volume more?", vbYesNo, "Crank it up?")
			
			Select Case result
			Case vbYes
				for i=0 to 4
				WshShell.SendKeys(chr(175))
				next
			Case vbNo
				continueVolumeUp = 1
			End Select
		loop
	ElseIf InStr(1, command, "volume") > 0 And InStr(1, command, "down") > 0 OR InStr(1, command, "volume") > 0 And InStr(1, command, "low") > 0 OR InStr(1, command, "volume") > 0 And InStr(1, command, "decrease") > 0 OR InStr(1, command, "volume") > 0 And InStr(1, command, "soft") > 0 Then
		Set WshShell = CreateObject("WScript.Shell")
		continueVolumeDown = 0
		
		for i=0 to 4
		WshShell.SendKeys(chr(174))
		next
		wscript.sleep(500)
		do while continueVolumeDown<1
			result = MsgBox ("Do you want to decrease the volume more?", vbYesNo, "Lower volume?")
			
			Select Case result
			Case vbYes
	    		for i=0 to 4
				WshShell.SendKeys(chr(174))
				next
			Case vbNo
				continueVolumeDown = 1
			End Select
		loop
	ElseIf InStr(1, command, "mute") > 0 Then
		Set WshShell = CreateObject("WScript.Shell")
		WshShell.SendKeys(chr(&hAD))
	ElseIf InStr(1, command, "video") > 0 OR InStr(1, command, "youtube") > 0 Then
		If bolActiveConnection = False Then
		wifiErrorForOpeningEvents = MsgBox ("Your computer is not connected to the internet. This functionality will not work without internet connection." & vbCrLf & "" & vbCrLf & "Do you want to continue?", vbYesNo, "No Internet Connection")
		
		Select Case wifiErrorForOpeningEvents
		Case vbYes
			Sapi.speak "What would you like to search for on youtube?"
			googleSearchQuery = inputbox("Please enter what you would like to search for on Youtube in the input box below...","Youtube Search")
			googleSearchQueryRefined=Trim(googleSearchQuery)
			googleSearchQueryRefined=Replace(googleSearchQueryRefined, "+", "%2B")
			googleSearchQueryRefined=Replace(googleSearchQueryRefined, "#", "%23")
			googleSearchQueryRefined=Replace(googleSearchQueryRefined, "%", "%25")
			googleSearchQueryRefined=Replace(googleSearchQueryRefined, "&", "%26")
			If googleSearchQueryRefined="" OR googleSearchQueryRefined=" " OR googleSearchQueryRefined="  " OR googleSearchQueryRefined="   " Then
				googleBlankSearch=msgbox("Searching only spaces or a blank will not work on youtube...",16,"Error: No Youtube Search content")
			Else
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, " ", "%20")
				startingurl="https://www.youtube.com/results?search_query="
				fullurl=startingurl+googleSearchQueryRefined
				Sapi.speak "Searching for, " + googleSearchQuery + ", on youtube"
				a.run fullurl
			End If
		Case vbNo
			wscript.sleep(1)
		End Select
		Else
			Sapi.speak "What would you like to search for on youtube?"
			googleSearchQuery = inputbox("Please enter what you would like to search for on Youtube in the input box below...","Youtube Search")
			googleSearchQueryRefined=Trim(googleSearchQuery)
			googleSearchQueryRefined=Replace(googleSearchQueryRefined, "+", "%2B")
			googleSearchQueryRefined=Replace(googleSearchQueryRefined, "#", "%23")
			googleSearchQueryRefined=Replace(googleSearchQueryRefined, "%", "%25")
			googleSearchQueryRefined=Replace(googleSearchQueryRefined, "&", "%26")
			If googleSearchQueryRefined="" OR googleSearchQueryRefined=" " OR googleSearchQueryRefined="  " OR googleSearchQueryRefined="   " Then
				googleBlankSearch=msgbox("Searching only spaces or a blank will not work on youtube...",16,"Error: No Youtube Search content")
			Else
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, " ", "%20")
				startingurl="https://www.youtube.com/results?search_query="
				fullurl=startingurl+googleSearchQueryRefined
				Sapi.speak "Searching for, " + googleSearchQuery + ", on youtube"
				a.run fullurl
			End If
		End If
	ElseIf command="google images search" OR command="search google images" OR InStr(1, command, "image") > 0 OR InStr(1, command, "picture") > 0 OR InStr(1, command, "picture") > 0 Then
		If bolActiveConnection = False Then
		wifiErrorForOpeningEvents = MsgBox ("Your computer is not connected to the internet. This functionality will not work without internet connection." & vbCrLf & "" & vbCrLf & "Do you want to continue?", vbYesNo, "No Internet Connection")
		
		Select Case wifiErrorForOpeningEvents
		Case vbYes
			Sapi.speak "What would you like to search for on google images?"
			googleSearchQuery = inputbox("Please enter what you would like to search for on Google Images in the input box below...","Google Images Search")
			googleSearchQueryRefined=Trim(googleSearchQuery)
			googleSearchQueryRefined=Replace(googleSearchQueryRefined, "+", "%2B")
			googleSearchQueryRefined=Replace(googleSearchQueryRefined, "#", "%23")
			googleSearchQueryRefined=Replace(googleSearchQueryRefined, "%", "%25")
			googleSearchQueryRefined=Replace(googleSearchQueryRefined, "&", "%26")
			If googleSearchQueryRefined="" OR googleSearchQueryRefined=" " OR googleSearchQueryRefined="  " OR googleSearchQueryRefined="   " Then
				googleBlankSearch=msgbox("Searching only spaces or a blank will not work on google images...",16,"Error: No Google Search content")
			Else
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, " ", "%20")
				startingurl="https://www.google.com/search?site=imghp&tbm=isch&source=hp&biw=1366&bih=643&q="
				endingurl = "&gs_l=img.3..0l10.455.1529.0.1867.5.5.0.0.0.0.112.434.3j2.5.0....0...1ac.1.64.img..0.5.429.nAKYnqn-VD0&gws_rd=cr&ei=h6qaV9XWAeqCjwSi97_4BQ&safe=active&ssui=on"
				fullurl=startingurl+googleSearchQueryRefined + endingurl
				Sapi.speak "Searching for, " + googleSearchQuery + ", on google images"
				a.run fullurl
			End If
		Case vbNo
			wscript.sleep(1)
		End Select
		Else
			Sapi.speak "What would you like to search for on google images?"
			googleSearchQuery = inputbox("Please enter what you would like to search for on Google Images in the input box below...","Google Images Search")
			googleSearchQueryRefined=Trim(googleSearchQuery)
			googleSearchQueryRefined=Replace(googleSearchQueryRefined, "+", "%2B")
			googleSearchQueryRefined=Replace(googleSearchQueryRefined, "#", "%23")
			googleSearchQueryRefined=Replace(googleSearchQueryRefined, "%", "%25")
			googleSearchQueryRefined=Replace(googleSearchQueryRefined, "&", "%26")
			If googleSearchQueryRefined="" OR googleSearchQueryRefined=" " OR googleSearchQueryRefined="  " OR googleSearchQueryRefined="   " Then
				googleBlankSearch=msgbox("Searching only spaces or a blank will not work on google images...",16,"Error: No Google Search content")
			Else
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, " ", "%20")
				startingurl="https://www.google.com/search?site=imghp&tbm=isch&source=hp&biw=1366&bih=643&q="
				endingurl = "&gs_l=img.3..0l10.455.1529.0.1867.5.5.0.0.0.0.112.434.3j2.5.0....0...1ac.1.64.img..0.5.429.nAKYnqn-VD0&gws_rd=cr&ei=h6qaV9XWAeqCjwSi97_4BQ&safe=active&ssui=on"
				fullurl=startingurl+googleSearchQueryRefined + endingurl
				Sapi.speak "Searching for, " + googleSearchQuery + ", on google images"
				a.run fullurl
			End If
		End If
	ElseIf command="google search" OR command="search google" OR InStr(1, command, "search") > 0 OR InStr(1, command, "google") > 0 Then
		If bolActiveConnection = False Then
		wifiErrorForOpeningEvents = MsgBox ("Your computer is not connected to the internet. This functionality will not work without internet connection." & vbCrLf & "" & vbCrLf & "Do you want to continue?", vbYesNo, "No Internet Connection")
		
		Select Case wifiErrorForOpeningEvents
		Case vbYes
			Sapi.speak "What would you like to search for on google?"
			googleSearchQuery = inputbox("Please enter what you would like to search for on Google in the input box below...","Google Search")
			googleSearchQueryRefined=Trim(googleSearchQuery)
			googleSearchQueryRefined=Replace(googleSearchQueryRefined, "+", "%2B")
			googleSearchQueryRefined=Replace(googleSearchQueryRefined, "#", "%23")
			googleSearchQueryRefined=Replace(googleSearchQueryRefined, "%", "%25")
			googleSearchQueryRefined=Replace(googleSearchQueryRefined, "&", "%26")
			If googleSearchQueryRefined="" OR googleSearchQueryRefined=" " OR googleSearchQueryRefined="  " OR googleSearchQueryRefined="   " Then
				googleBlankSearch=msgbox("Searching only spaces or a blank will not work on google...",16,"Error: No Google Search content")
			Else
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, " ", "%20")
				startingurl="https://www.google.ca/?gws_rd=cr&ei=acGMV629B8TQ-QH6vZroDQ&safe=active&ssui=on#safe=active&q="
				fullurl=startingurl+googleSearchQueryRefined
				Sapi.speak "Searching for, " + googleSearchQuery + ", on google"
				a.run fullurl
			End If
		Case vbNo
			wscript.sleep(1)
		End Select
		Else
			Sapi.speak "What would you like to search for on google?"
			googleSearchQuery = inputbox("Please enter what you would like to search for on Google in the input box below...","Google Search")
			googleSearchQueryRefined=Trim(googleSearchQuery)
			googleSearchQueryRefined=Replace(googleSearchQueryRefined, "+", "%2B")
			googleSearchQueryRefined=Replace(googleSearchQueryRefined, "#", "%23")
			googleSearchQueryRefined=Replace(googleSearchQueryRefined, "%", "%25")
			googleSearchQueryRefined=Replace(googleSearchQueryRefined, "&", "%26")
			If googleSearchQueryRefined="" OR googleSearchQueryRefined=" " OR googleSearchQueryRefined="  " OR googleSearchQueryRefined="   " Then
				googleBlankSearch=msgbox("Searching only spaces or a blank will not work on google...",16,"Error: No Google Search content")
			Else
				googleSearchQueryRefined=Replace(googleSearchQueryRefined, " ", "%20")
				startingurl="https://www.google.ca/?gws_rd=cr&ei=acGMV629B8TQ-QH6vZroDQ&safe=active&ssui=on#safe=active&q="
				fullurl=startingurl+googleSearchQueryRefined
				Sapi.speak "Searching for, " + googleSearchQuery + ", on google"
				a.run fullurl
			End If
		End If
	ElseIf command="dev.stats" Then
		password=InputBox("Please enter the correct password to view developper info:","Developper Stats")
		masterNameRefined=LCase(masterName)
		if password="sternocleidomastoid " Then
			correct=msgbox("Hi Albert, the password you entered is correct. I hope you are having a great day today!" & vbcrlf & "" & vbcrlf & "I will list all the improvements to be done to the Sherby Interface after this message. Good luck implementing them!",64,"Correct Password")
			Sapi.speak "Here is a list of improvements that are scheduled to be implemented to the Sherby Interface"
			improvements=msgbox("Make Sherby recognise different things related to the weather, (ex. 'what's the weather THIS WEEK) make Sherby recognise questions like 'what day is it tomorrow' or 'what day is it on Sunday' or 'is it Sunday today' or 'is today August first', make Sherby use different weather api, make Sherby use different api for synonyms and antonyms, make movies, sports, restaurants, driving, locations and stocks command, make Sherby repeat data (ex. 'A bagel' have xxx many calories), make Sherby say something like 'checking my sources' when searching Wolfram Alpha, make alarms and reminders command, make Sherby know when user is asking about info of a person and display that info, make translation command, make a 'questions' command where Sherby consults wolfram alpha, make Sherby know when to take data from wikipedia (when find keyword 'what is'/'what was'/etc.), make Sherby know when data from Wolfram Alpha is useless, make a 'school day' teller, make timer, improve 'greeting' when the user first starts the program/interface up, ability to go back/cancel, make 'settings page' where user can control different things such as disabling Sherby greeting on Start Up, make Sherby talk more like a servant and make it seem more user friendly, make user able to create shortcut to execute command, expand vocab (ex. for 'enable voice recognition', also put 'start up voice recognition')" & vbcrlf & "" & vbcrlf & "BONUS: Create a talking interface, make it look user friendly",,"Sherby Improvements")
			
		Else
			incorrectDEVStats=msgbox("The password you entered to view developer stats is incorrect. Please try again later.",16,"Incorrect password")
		End If
	ElseIf InStr(1, command, "shutdown") > 0 OR InStr(1, command, "shut down") > 0 OR InStr(1, command, "shut") And InStr(1, command, "down") Then
		Dim resultShutdown
		resultShutdown = MsgBox ("Are you sure you want to shut your computer down? (It will close all other programs)", vbYesNo, "Shutdown computer?")
		Select Case resultShutdown
			Case vbYes
				SapiShutdownSpeak = "Shutting down, goodbye " + CStr(masterName)
				Sapi.speak SapiShutdownSpeak
				Dim objShell
				Set objShell = WScript.CreateObject("WScript.Shell")
				objShell.Run "C:\WINDOWS\system32\shutdown.exe -r -t 20"
			Case vbNo
				Sapi.speak "The computer shutdown has been cancelled"
		End Select
	ElseIf InStr(1, command, "restart") > 0 OR InStr(1, command, "reboot") > 0 Then
		Dim resultRestart
		resultRestart = MsgBox ("Are you sure you want to restart your computer?", vbYesNo, "Restart computer?")
		Select Case resultRestart
			Case vbYes
				SapiRestartSpeak = "Restarting computer, see you soon " + CStr(masterName)
				Sapi.speak SapiRestartSpeak
				Set objShell = WScript.CreateObject("WScript.Shell")
				objShell.Run "C:\WINDOWS\system32\shutdown.exe -r -t 0"
			Case vbNo
				Sapi.speak "The computer restart has been cancelled"
		End Select
	ElseIf InStr(1, command, "rename") > 0 OR InStr(1, command, "call") > 0 Then
		previousName=masterName
		Sapi.speak "What would you like your new name to be?"
		changeName="Your current name: " + masterName & vbcrlf & "" & vbcrlf & "Please input your new name in the input box below." & vbcrlf & "IF YOU DON'T WANT SHERBY TO SAY YOUR NAME, SIMPLY TYPE IN THE COMMAND 'SKIP'..." & vbcrlf & "" & vbcrlf & "Psst, you don't have to put in your real name, you can name yourself 'Lord Vader', or 'Star Lord' if you want!"
		masterName=inputbox(changeName,,masterName)
		masterNameSkip=LCase(masterName)
		If masterNameSkip="skip" Then
			masterName=" "
		Else
			masterName=masterName
		End If
	
		Set objFSO=CreateObject("Scripting.FileSystemObject")
		
		outFileForName="C:\SherbyInterface\names.txt"
		Set objFileName = objFSO.CreateTextFile(outFileForName,True)
		objFileName.Write masterName
		objFileName.Close
		
		nameHasBeenChanged="Your name has been changed from, " + previousName + ", to, " + masterName
		Sapi.speak nameHasBeenChanged
	ElseIf InStr(1, command, "version") > 0 Then
		Sapi.speak "The current version of the Sherby Interface is version 23."
	ElseIf InStr(1, command, "voice") > 0 And InStr(1, command, "recognition") > 0 And InStr(1, command, "disable") > 0 Then
		Sapi.speak "I have now turned off voice recognition. You can turn it back on later by typing the command, voice recognition."
		confirmvoicerecognition=""
		Set objFSOVoice=CreateObject("Scripting.FileSystemObject")
		
		outFileForVoice="voicerecognition.txt"
		Set objFileNameVoice = objFSOVoice.CreateTextFile(outFileForVoice,True)
		objFileNameVoice.Write confirmvoicerecognition
		objFileNameVoice.Close
	ElseIf InStr(1, command, "voice") > 0 And InStr(1, command, "recognition") > 0 Then
		Sapi.speak "I'm very glad you decided to turn on voice recognition! The setup to enable voice recognition will open shortly. Enjoy!"
		wscript.sleep(500)
		confirmvoicerecognition="enabled"
		Set objFSOVoice=CreateObject("Scripting.FileSystemObject")
		
		outFileForVoice="voicerecognition.txt"
		Set objFileNameVoice = objFSOVoice.CreateTextFile(outFileForVoice,True)
		objFileNameVoice.Write confirmvoicerecognition
		objFileNameVoice.Close
		wshshell.run "%windir%\Speech\Common\sapisvr.exe -SpeechUX"
	ElseIf InStr(1, command, "can") > 0 OR InStr(1, command, "what") > 0 OR InStr(1, command, "can") > 0 Then
		If command="can" OR command="what" Then
			whatWhenArray = Array("How am I supposed to know?", "Don't ask me", "You're the boss", "Who knows?", "I don't know")
			Sapi.speak rand(whatWhenArray)
		End If
		doWolframSearch(command)
	ElseIf InStr(1, command, "word") > 0 Then
		Sapi.speak "Creating new word document"
		Set objWord = CreateObject("Word.Application")
		objWord.Visible = True
	ElseIf InStr(1, command, "excel") > 0 OR InStr(1, command, "spread") > 0 And InStr(1, command, "sheet") > 0 Then
		If InStr(1, command, "excel") > 0 Then
			Sapi.speak "Creating new excel document"
		Else
			Sapi.speak "Creating new spreadsheet"
		End If
		Set objExcel = CreateObject("Excel.Application")
		objExcel.Visible = True

		Set objWorkbook = objExcel.Workbooks.Add()
	ElseIf InStr(1, command, "power") > 0 And InStr(1, command, "point") > 0 OR InStr(1, command, "slide") > 0 And InStr(1, command, "show") > 0 Then
		If InStr(1, command, "power") > 0 And InStr(1, command, "point") > 0 Then
			Sapi.speak "Creating new powerpoint document"
		Else
			Sapi.speak "Creating new spreadsheet"
		End If
		Set objPPT = CreateObject("PowerPoint.Application")
		objPPT.Visible = True
		Set objPresentation = objPPT.Presentations.Add
	Else
		notSure = Array("I'm not quite sure what you mean. Would you like me to search for, #%query%#, on google?", "I'm not sure I follow. However, we can search the web for, #%query%#.", "I don't understand what you mean by, #%query%#, but we can do a web search for it!", "My apologies, I didn't really grasp what you meant by, #%query%#, but I bet Google could help you!", "I beg you pardon, I don't understand what you mean by #%query%#. However, I think Google can help you with that.", "I'm a little unclear on what you meant by, #%query%#. Would you like me to search the web for it?", "I'm afraid I don't understand what you mean by, #%query%#, but I can help you by searching the web for it. How about that?", "I'd be glad to help, but I'm not sure what you meant by #%query%#. However, I CAN search the web for it...", "Sorry, I don't understand what you mean by #%query%#, but I can help by doing a web search for it!")
		notSureSelect = rand(notSure)
		notSureSelect = Replace(notSureSelect, "#%query%#", command)
		Sapi.speak notSureSelect
		' Sapi.speak "I'm not quite sure I understand what you mean. Would you like to search for, " + command + ", on google?"
		googlecommand = MsgBox ("Sherby did not recognise/understand what you said." & vbcrlf & "If you wish to add this command, email anitu1@ocdsb.ca to ask him to add the command you want." & vbcrlf & "" & vbcrlf & "Would you like to search that command on google?", 20, "Error: Command not found")
		
		Select Case googlecommand
		Case vbYes
			Sapi.speak "Searching for, " & CStr(command) & ", on google..."
			googleSearchQueryRefinedG=Trim(command)
			googleSearchQueryRefinedG=Replace(googleSearchQueryRefinedG, "+", "%2B")
			googleSearchQueryRefinedG=Replace(googleSearchQueryRefinedG, "#", "%23")
			googleSearchQueryRefinedG=Replace(googleSearchQueryRefinedG, "%", "%25")
			googleSearchQueryRefinedG=Replace(googleSearchQueryRefinedG, "&", "%26")
			googleSearchQueryRefinedG=Replace(googleSearchQueryRefinedG, " ", "%20")
			startingurlG="https://www.google.ca/?gws_rd=cr&ei=acGMV629B8TQ-QH6vZroDQ&safe=active&ssui=on#safe=active&q="
			fullurlG=startingurlG+googleSearchQueryRefinedG
			a.run fullurlG
		Case vbNo
			wscript.sleep(1)
		End Select
	End If
	wscript.sleep(500)
	anythingElseArray = Array("Is there anything else I can do for you?", "Would you like me to do anything else for you", "Can I assist you in any other way?", "Is there anything else I can be of assistance for?", "Do you require any other assistance?", "May I be of assistance in any other way?", "Is there anything else I can help you with", "Can I do anything else for you?", "Is there anything else I can help you with?")
	Sapi.speak rand(anythingElseArray)
	result = MsgBox ("Is there anything else I can assist you with?", vbYesNo, "Anything Else?")
	
	Select Case result
	Case vbYes
	    Call artificialIntelligence()
	Case vbNo
		Call Play("C:\Windows\Media\notify.wav")
		currentHour = Hour(Now())
		If currentHour >= 1 And currentHour < 12 Then
			byeArrayMorning = Array("Okay then, goodbye ", "Okay, goodbye ", "Well I guess that wraps it up. See you later ", "Okay then. until next time ", "I guess I'll be going then. Bye ", "Okay, see you later. Have a great day ", "until next time ", "Nice talking to you! Have a great day ", "Okay, I better be on my way. See you later ")
			msg = rand(byeArrayMorning) + CStr(masterName)
		ElseIf currentHour >= 12 And currentHour <18 Then
			byeArrayMorning = Array("Okay then, goodbye ", "Okay, goodbye ", "Well I guess that wraps it up. See you later ", "Okay then. until next time ", "I guess I'll be going then. Bye ", "Okay, see you later. Have a great rest of the day ", "until next time ", "Nice talking to you! Have a great rest of the day ", "Anyways, I think it's time to go. Have a great day ")
			msg = rand(byeArrayMorning) + CStr(masterName)
		ElseIf currentHour >=18 And currentHour<21 Then
			byeArrayMorning = Array("Okay then, goodbye ", "Okay, goodbye ", "Well I guess that wraps it up. See you later ", "Okay then. until next time ", "I guess I'll be going then. Bye ", "Okay, see you later. Have a great evening ", "until next time ", "Nice talking to you! Have a great evening ")
			msg = rand(byeArrayMorning) + CStr(masterName)
		ElseIf currentHour >= 21 Then
			byeArrayMorning = Array("Okay then, goodnight! Oh, and don't let the bed bugs bite! ", "Okay, goodbye and sleep tight! ", "Well I guess that wraps it up. Sweet dreams!", "Okay then. See you later, and don't let the bed bugs bite! ", "I guess I'll be going then. Sleep tight ", "Okay, see you later. Have a great night and don't let the bed bugs bite ", "until next time, ", "Nice talking to you! Good night!")
			msg = rand(byeArrayMorning)
		End If
		Set greeter = CreateObject("sapi.spvoice")
		greeter.Speak msg
		wscript.quit
	End Select
End Function

Function userStartup()
	Set Sapi = Wscript.CreateObject("SAPI.SpVoice")
	Set WshShell = CreateObject("Wscript.shell")
	strDesktop = WshShell.SpecialFolders("Desktop")
	Set oMyShortcut = WshShell.CreateShortcut(strDesktop + "\Sherby.lnk")
	oMyShortcut.IconLocation = "C:\Sherby Interface\sherbyicon.ico"
	OMyShortcut.TargetPath = "C:\Sherby Interface\schoolhelper_v23.exe"
	oMyShortCut.Hotkey = "ALT+CTRL+S"
	oMyShortCut.Save
	Sapi.speak "Hello There! Welcome to Sherby Interface, an interface designed especially for Broadview students from Broadview Avenue Public School. This interface is meant to help students for an ample amount of tasks, such as helping for the simplest tasks, such as telling the date or time, to more complex and useful actions, such as listing the assignments, special events, schedule or saying the current school date. In fact, the Sherby interface can do over 20 different tasks, and can figure out that two different texts can mean the same thing. For example, Sherby will know that, tell me the day, is the same thing as, what date is it today,. All in all, the Sherby Interface is a very user friendly artificial intelligence program that helps students through their eventful journey through middle school. So, enough talk, let's get to the Sherby interface setup. The setup for this interface is fairly simple, and consists of only one mandatory step, which is to tell us your name, but, for a more pleasant experience, we suggest that you also do one other step. This step allows the Sherby interface to use voice recognition, so you don't only have to type in the commands, but you can also verbally tell sherby what to do."
	Sapi.speak "But firstly, let me introduce myself. My name is Sherby, and I will be your personal assistant for the rest of the year, and I look forward to knowing more about you. Speaking about this, what is your name?"
	masterName=inputbox("Please input your name in the input box below." & vbcrlf & "IF YOU DON'T WANT TO INPUT YOUR NAME, SIMPLY TYPE IN 'SKIP'..." & vbcrlf & "" & vbcrlf & "Psst, you don't have to put in your real name, you can name yourself 'Lord Vader', or 'Star Lord' if you want! Also, you can easily change your name later on if you want.","What is your name?")
	masterNameSkip=LCase(masterName)
	If masterNameSkip="skip" Then
		masterName=" "
	Else
		masterName=masterName
	End If

	Set objFSO=CreateObject("Scripting.FileSystemObject")
	
	outFileForName="C:\SherbyInterface\names.txt"
	Set objFileName = objFSO.CreateTextFile(outFileForName,True)
	objFileName.Write masterName
	objFileName.Close
	
	advancedSetup="Now, " + masterName + ", would you like to enable voice recognition?"
	
	Sapi.speak advancedSetup
	
	resultVoice = MsgBox ("Would you now like to enable voice recognition, for a smoother and better experience?", vbYesNo, "Voice Recognition")
	
	Select Case resultVoice
	Case vbYes
		set wshshell = wscript.CreateObject("wscript.shell")
	    Sapi.speak "Great! The setup to enable voice recognition will now open. Enjoy!"
		wscript.sleep(500)
		wshshell.run "%windir%\Speech\Common\sapisvr.exe -SpeechUX"
		confirmvoicerecognition="enabled"
		Set objFSOVoice=CreateObject("Scripting.FileSystemObject")
		
		outFileForVoice="voicerecognition.txt"
		Set objFileNameVoice = objFSOVoice.CreateTextFile(outFileForVoice,True)
		objFileNameVoice.Write confirmvoicerecognition
		objFileNameVoice.Close
	Case vbNo
		Sapi.speak "Okay then."
		Sapi.speak "If you ever change your mind, simply enter the command, voice recognition, in the command box and the setup to enable voice recognition will appear."
	End Select
	wscript.sleep (1000)
	Sapi.speak "Hooray!, You are now done the Sherby interface setup. Wasn't it easy?"
	wscript.sleep(500)
	Sapi.speak "Anyways, you are now ready. When you start up the Sherby Interface, an input box will appear. This is the input box where you will type in a command that will be executed, such as, what is the date today? To view all commands, type in the command help."
	wscript.sleep(500)
	Sapi.speak "So, you are now finally fully done the setup. Good luck, and have fun!"
	
End Function

Set greeter = CreateObject("sapi.spvoice")
Function greetingUser()
	Call Play("C:\WINDOWS\Media\Windows Notify Calendar.wav")
	currentHour = Hour(Now())
	currentWeekday = Weekday(Now())
	If currentHour >= 1 And currentHour < 12 Then
		morningGreetingArrayFirst = Array("Good morning,", "Hello,", "Hi,", "Greetings,", "Hey,", "What's up,", "How's it going,", "Nice to see you,", "Howdy,", "Well hello,")
		morningGreetingArraySecond = Array("I hope you've had a good night's sleep! How may I be of assistance today?", "I hope you've had a good night's sleep! How may I help you today?", "How can I be of assistance today?", "How may I help you today?", "I hope that all is well with you. Is there anything I can do for you?")
		maxfirst=UBound(morningGreetingArrayFirst)
		minfirst=0
		Randomize
		randomOutcomeGreetingFirst=Int((maxfirst-minfirst+1)*Rnd+minfirst)
		randomElemGreetingFirst=morningGreetingArrayFirst(randomOutcomeGreetingFirst)
		firstSpeaking = CStr(randomElemGreetingFirst)

		maxsecond=UBound(morningGreetingArraySecond)
		minsecond=0
		Randomize
		randomOutcomeGreetingSecond=Int((maxsecond-minsecond+1)*Rnd+minsecond)
		randomElemGreetingSecond=morningGreetingArraySecond(randomOutcomeGreetingSecond)
		secondSpeaking = CStr(randomElemGreetingSecond)
		msg = firstSpeaking + masterName + "..."
		greeter.Speak msg
		greeter.Speak secondSpeaking
	ElseIf currentHour >= 12 And currentHour <18 Then
		morningGreetingArrayFirst = Array("Good afternoon,", "Hello,", "Hi,", "Greetings,", "Hey,", "What's up,", "How's it going,", "Nice to see you,", "Howdy,", "Well hello,")
		morningGreetingArraySecond = Array("I hope your day so far has been as pleasant as mine! How may I be of assistance today?", "I hope you've had a great day so far. How may I help you today?", "How may I help you today?", "How can I help you today?", "How may I be of assistance today?", "I hope you've had a wonderful day! How may I be of any assistance", "I hope you've had a wonderful day! What can I do for you this afternoon?")
		maxfirst=UBound(morningGreetingArrayFirst)
		minfirst=0
		Randomize
		randomOutcomeGreetingFirst=Int((maxfirst-minfirst+1)*Rnd+minfirst)
		randomElemGreetingFirst=morningGreetingArrayFirst(randomOutcomeGreetingFirst)
		firstSpeaking = CStr(randomElemGreetingFirst)

		maxsecond=UBound(morningGreetingArraySecond)
		minsecond=0
		Randomize
		randomOutcomeGreetingSecond=Int((maxsecond-minsecond+1)*Rnd+minsecond)
		randomElemGreetingSecond=morningGreetingArraySecond(randomOutcomeGreetingSecond)
		secondSpeaking = CStr(randomElemGreetingSecond)
		msg = firstSpeaking + masterName + "..."
		greeter.Speak msg
		greeter.Speak secondSpeaking
	ElseIf currentHour >=18 Then
		morningGreetingArrayFirst = Array("Good evening,", "Hello,", "Hi,", "Greetings,", "Hey,", "What's up,", "How's it going,", "Nice to see you,", "Howdy,", "Well hello,")
		morningGreetingArraySecond = Array("I hope you've had a great day today! How may I be of assistance this evening?", "I hope you've had a great day so far. How may I help you today?", "How may I help you today?", "How can I help you today?", "How may I be of assistance today?", "I hope you've had a wonderful day! How may I be of any assistance", "I hope you've had a wonderful day! What can I do for you this evening?", "I hope you've enjoyed your day so far! Is there anything I can do for you?")
		maxfirst=UBound(morningGreetingArrayFirst)
		minfirst=0
		Randomize
		randomOutcomeGreetingFirst=Int((maxfirst-minfirst+1)*Rnd+minfirst)
		randomElemGreetingFirst=morningGreetingArrayFirst(randomOutcomeGreetingFirst)
		firstSpeaking = CStr(randomElemGreetingFirst)

		maxsecond=UBound(morningGreetingArraySecond)
		minsecond=0
		Randomize
		randomOutcomeGreetingSecond=Int((maxsecond-minsecond+1)*Rnd+minsecond)
		randomElemGreetingSecond=morningGreetingArraySecond(randomOutcomeGreetingSecond)
		secondSpeaking = CStr(randomElemGreetingSecond)
		msg = firstSpeaking + masterName + "..."
		greeter.Speak msg
		greeter.Speak secondSpeaking
	End If
End Function


Function getName()
	loopNumberTimes=1
	Set objFileToReadUserName = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Sherby Interface\names.txt",1)
	do while not objFileToReadUserName.AtEndOfStream
	     If loopNumberTimes=1 Then
	     	masterName = objFileToReadUserName.ReadLine()
	     End If
	     loopNumberTimes = loopNumberTimes + 1
	loop
	objFileToReadUserName.Close
	
	If masterName="" Then
		Call userStartup()
	Else
		Call greetingUser()
	End If
End Function



Call getName()
Call artificialIntelligence()