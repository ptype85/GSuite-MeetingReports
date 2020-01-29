###Send GSuite Calendars as Spreadsheet
###pete.endacott@gmail.com
###January 2020
###Requires PSGSuite module installed. Must be configured to Google API user with necessary permissions (https://www.googleapis.com/auth/calendar). See https://psgsuite.io/Initial%20Setup/ for more details.

Import-Module PSGSuite

$currentDate = (Get-Date).ToString("dd-MM-yyyy HH:mm:ss")
$futureDate = (Get-Date).AddDays(+7).ToString("dd-MM-yyyy HH:mm:ss")
$freindlyDate = (Get-Date).ToString("dd MMMM yyyy")
$wCDate = (Get-Date).ToString("dd-MM-yy")

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
    $eventCount = 0
    #$userName | Out-File C:\Test\Details.txt -Append
    forEach ($event in $events){
    $eventcount ++
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
    $UserName, $startDateTime, $endDateTime, $title, $attList -join ',' | Out-File C:\Test\Details.csv -Append
    }
    $output = "$userName has $eventCount meetings this week. <br>"
    $output | Out-File C:\Test\summary.txt -Append
}


###Loop
foreach($user in $users){
getCalendarEvents $user
}

$summary = get-content "C:\Test\summary.txt"
#$details = (get-content "C:\Test\details.txt") -join [Environment]::NewLine
$details = Import-csv "C:\test\Details.csv" -Header user,start,end,title,atts
    foreach($detail in $details){
        $user = $detail.user
        $start = $detail.start
        $end = $detail.end
        $title = $detail.title
        $atts = $detail.atts

        if($start.length -gt 1){

                ####Date formatting
                $sdatetime = $start.Split(' ')
                $sdate = $sdatetime[0]
                $sdateDF = Get-Date($sdate)
                $sWeekDay = $sdateDF.DayOfWeek
                $stime = $sdatetime[1]
                $sdatesplit = $sdate.split('/')
                $sday = $sdatesplit[0]
                $smonth = $sdatesplit[1]
                $smonth = $smonth - 1
                $syear = $sdatesplit[2]
                $stimesplit = $stime.split(":")
                $shour = $stimesplit[0]
                $sminute = $stimesplit[1]

                $edatetime = $end.Split(' ')
                $edate = $edatetime[0]
                $etime = $edatetime[1]
                $edatesplit = $edate.split('/')
                $eday = $edatesplit[0]
                $emonth = $edatesplit[1]
                $emonth = $emonth - 1
                $eyear = $edatesplit[2]
                $etimesplit = $etime.split(":")
                $ehour = $etimesplit[0]
                $eminute = $etimesplit[1]

                if($sWeekDay -eq "Monday"){
                        $mhtml = "['$user', null, createCustomHTMLContent('https://upload.wikimedia.org/wikipedia/commons/thumb/a/a9/Google_Calendar_icon.svg/1024px-Google_Calendar_icon.svg.png', '$title', '$user', '$shour`:$sminute'), new Date($syear, $sday, $smonth, $shour, $sminute), new Date($eyear, $eday, $emonth, $ehour, $eminute)],"

                        "$mhtml" | Out-File C:\Test\mhtml.txt -Append
                                                                   }
                if($sWeekDay -eq "Tuesday"){
                         $thtml = "['$user', null, createCustomHTMLContent('https://upload.wikimedia.org/wikipedia/commons/thumb/a/a9/Google_Calendar_icon.svg/1024px-Google_Calendar_icon.svg.png', '$title', '$user', '$shour`:$sminute'), new Date($syear, $sday, $smonth, $shour, $sminute), new Date($eyear, $eday, $emonth, $ehour, $eminute)],"
                         "$thtml" | Out-File C:\Test\thtml.txt -Append                         
                        
                                           }
                if($sWeekDay -eq "Wednesday"){
                         $whtml = "['$user', null, createCustomHTMLContent('https://upload.wikimedia.org/wikipedia/commons/thumb/a/a9/Google_Calendar_icon.svg/1024px-Google_Calendar_icon.svg.png', '$title', '$user', '$shour`:$sminute'), new Date($syear, $sday, $smonth, $shour, $sminute), new Date($eyear, $eday, $emonth, $ehour, $eminute)],"
                         "$whtml" | Out-File C:\Test\whtml.txt -Append                         
                                                                   }
                if($sWeekDay -eq "Thursday"){
                         $thhtml = "['$user', null, createCustomHTMLContent('https://upload.wikimedia.org/wikipedia/commons/thumb/a/a9/Google_Calendar_icon.svg/1024px-Google_Calendar_icon.svg.png', '$title', '$user', '$shour`:$sminute'), new Date($syear, $sday, $smonth, $shour, $sminute), new Date($eyear, $eday, $emonth, $ehour, $eminute)],"
                         "$thhtml" | Out-File C:\Test\thhtml.txt -Append 
                                                                    }
                if($sWeekDay -eq "Friday"){
                         $fhtml = "['$user', null, createCustomHTMLContent('https://upload.wikimedia.org/wikipedia/commons/thumb/a/a9/Google_Calendar_icon.svg/1024px-Google_Calendar_icon.svg.png', '$title', '$user', '$shour`:$sminute'), new Date($syear, $sday, $smonth, $shour, $sminute), new Date($eyear, $eday, $emonth, $ehour, $eminute)],"
                         "$fhtml" | Out-File C:\Test\fhtml.txt -Append                         
                                           }  

                                           }
            
                
        }

$emailFn = $wCDate
#(Get-content "G:\My Drive\Projects\2020-01 - Meeting Report\CalEmailTemplate.html").Replace('%SUMMARY%', $summary).Replace('%EMAILDATE%', $wCDate).Replace('%DETAILS%', $details) | Set-Content "G:\My Drive\Projects\2020-01 - Meeting Report\$emailFn.html"

$mondayRows = Get-Content "C:\Test\mhtml.txt"
                $tuesdayRows = Get-Content "C:\Test\thtml.txt"
                $wednesdayRows = Get-Content "C:\Test\whtml.txt"
                $thursdayRows = Get-Content "C:\Test\thhtml.txt"
                $fridayRows = Get-Content "C:\Test\fhtml.txt"
                
                $mdate = Get-Date
                while ($mdate.DayOfWeek -ne "Monday") {$mdate = $mdate.AddDays(1)}
                $mdate = $mdate.ToString("dd MMMM yyyy")
                $tdate = Get-Date
                while ($tdate.DayOfWeek -ne "Tuesday") {$tdate = $tdate.AddDays(1)}
                $tdate = $tdate.ToString("dd MMMM yyyy")
                $wdate = Get-Date
                while ($wdate.DayOfWeek -ne "Wednesday") {$wdate = $wdate.AddDays(1)}
                $wdate = $wdate.ToString("dd MMMM yyyy")
                $thdate = Get-Date
                while ($thdate.DayOfWeek -ne "Thursday") {$thdate = $thdate.AddDays(1)}
                $thdate = $thdate.ToString("dd MMMM yyyy")
                $fdate = Get-Date
                while ($fdate.DayOfWeek -ne "Friday") {$fdate = $fdate.AddDays(1)}
                $fdate = $fdate.ToString("dd MMMM yyyy")
                                                    
                (Get-content "G:\My Drive\Projects\2020-01 - Meeting Report\edittest.html").Replace('%%SUMMARY%%', $summary).Replace('%%MONDAYROWS%%', $mondayRows).Replace('%%TUESDAYROWS%%', $tuesdayRows).Replace('%%WEDNESDAYROWS%%', $wednesdayRows).Replace('%%THURSDAYROWS%%', $thursdayRows).Replace('%%FRIDAYROWS%%', $fridayRows).Replace('%%TODAYSDATE%%', $freindlyDate).Replace('%%MONDAYDATE%%', "Monday $mdate").Replace('%%TUESDAYDATE%%', "Tuesday $tdate").Replace('%%WEDNESDAYDATE%%', "Wednesday $wdate").Replace('%%THURSDAYDATE%%', "Thursday $thdate").Replace('%%FRIDAYDATE%%', "Friday $fdate") | Set-Content "G:\My Drive\Projects\2020-01 - Meeting Report\$emailFn.html"

