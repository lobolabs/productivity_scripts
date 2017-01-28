#!/usr/bin/osascript
tell application "Google Chrome"
	set theURL to URL of active tab of window 1
end tell

tell application "Alfred 3" to activate
tell application "System Events"
	keystroke "," using command down
	delay 1
	-- hit the down button twice
	tell application "System Events" to key code 125
	tell application "System Events" to key code 125
	try
		tell process "Alfred Preferences"
			delay 1
			click button "Add Custom Search" of tab group 1 of window 1
			delay 1
			-- UI elements of sheet of window 1
			set value of text field 2 of sheet 1 of window 1 to theURL
		end tell
		-- hit tab
		tell application "System Events" to key code 48
		delay 5
		tell process "Alfred Preferences"
			click button "Save" of sheet 1 of window 1
		end tell
	on error errMsg
		display dialog "Error: " & errMsg
	end try
end tell


