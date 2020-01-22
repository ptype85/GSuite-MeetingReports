###Send GSuite Calendars as Spreadsheet
###pete.endacott@gmail.com
###January 2020
###Requires PSGSuite module installed. Must be configured to Google API user with necessary permissions (https://www.googleapis.com/auth/calendar). See https://psgsuite.io/Initial%20Setup/ for more details.

Import-Module PSGSuite

$currentDate = (Get-Date).ToString("dd-MM-yyyy HH:mm:ss")
$futureDate = (Get-Date).AddDays(+7).ToString("dd-MM-yyyy HH:mm:ss")

############EDITABLE VARIABLES############

###User List
$users = "bob.test@test.com", "peter.endacott@gmail.com"

############END OF EDITABLE VARIABLES############

###FUNCTIONS
function getCalendarEvents($uN){

    ##Get Variables
    $userName = $uN

    ##Get Events
    $events = Get-GSCalendarEvent -User $userName -TimeMin (Get-Date $currentDate) -TimeMax (Get-Date $futureDate) -SingleEvents
    forEach ($event in $events){
    $calendar = $event.User
    $startDateTime = $event.Start.DateTime
    $isEndTime = $event.EndTimeUnspecified
    if ($isEndTime -eq $null){
        $endDateTime = $event.End.DateTime
    } else
    {
        $endDateTime = (Get-Date $startDateTime[0]).AddHours(+1).ToString("dd-MM-yyyy HH:mm:ss")
    }
    $title = $event.Summary
    $attList = $event.Organizer.Email
    $attLength = $attList.length
    "Calndar = $calendar, Start Time = $startDateTime, End Time = $endDateTime, Title = $title, Attendees = $attList" | Out-File C:\Test\mettings.txt -Append
    }
}


###Loop
foreach($user in $users){
getCalendarEvents $user
}
