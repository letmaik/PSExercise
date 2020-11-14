# Tip: Use VS Code to edit this file with syntax highlighting!

# Options used for task registration.
# Any change requires re-running register.ps1.
$startTime = "09:00"
$stopTime  = "17:00"
$interval  = "1H" # or "30M" or "1H30M" etc.
$weekDays = @( "Monday" ; "Tuesday" ; "Wednesday" ; "Thursday" ; "Friday" )

# Options used at task execution.
# Changes can be made at any time.
$ask = $true # if $false, video starts immediately
$askTitle = "Time to move!"
$askText = "Are you ready?"
$videoMonitor = "largest" # or screen number, $videoMonitor = 1
$otherMonitorsOverlay = "glass" # "none" or "glass" or "#rrggbb" (black: #000000) or image path/URL

$videos = @(

# No end time given, need to press Alt+F4 to exit. Add t2=332 to fix it!
@{ url="https://youtu.be/RqcOCBb4arc?t=132"; }

# End time given, automatically exits based on computed duration.
@{ url="https://youtu.be/K4dmZ5_n6uU?t=35"; t2=251 }

# Add more videos here...

)
