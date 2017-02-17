#!/usr/bin/osascript

tell application "System Events"
    set ProcessList to name of every process
    if "Alfred Preferences" is in ProcessList then
        set ThePID to unix id of process "Alfred Preferences"
        do shell script "kill -KILL " & ThePID
    end if
end tell

tell application "System Events" to quit
set delay_time to 0.3
tell application "Google Chrome"
    set theURL to URL of active tab of window 1
end tell

tell application "Alfred 3" to activate
tell application "System Events"
    keystroke "," using command down
    delay delay_time
    tell process "Alfred Preferences"
        delay delay_time
        -- hack: to ensure that the the "Default results" button isn't selected
        -- by default, we select another tab first, then select the "features" tab
        click button "Appearance" of toolbar 1 of window 1
        click button "Features" of toolbar 1 of window 1
        delay delay_time
        UI elements of scroll area 1 of window 1
        -- bring the left menu to focus first (see hack above, without hack this would work sometimes only)
        keystroke tab
        delay delay_time
        -- hit the down button twice
        key code 125
        key code 125
    end tell
    try
        tell process "Alfred Preferences"
            delay delay_time
            click button "Add Custom Search" of tab group 1 of window 1
            delay delay_time
            -- UI elements of sheet of window 1
            set value of text field 2 of sheet 1 of window 1 to theURL
        end tell
        -- hit tab
        tell application "System Events" to key code 48
    on error errMsg
        display dialog "Error: " & errMsg
    end try
    
end tell
