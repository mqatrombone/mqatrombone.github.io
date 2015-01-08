---
layout: post
title: Powershell, EWS, and DDay.iCal
---


Sometimes you have to work with constraints. In this case, I need a way to
allow individuals access to a calendar. People who aren't part of the 
Exchange environment. My first inclination was to use Exchange's published
ICS feature, but I also need at least some authentication. It would also
be nice if I didn't have to use Outlook's upload to WebDAV feature.

Thankfully Microsoft has a .NET API to talk to Exchange via Exchange Web
Services. I can use Powershell to put together a script to consume the
calendar, and then I'm using
[DDay.iCal](http://www.ddaysoftware.com/Pages/Projects/DDay.iCal/)
to spit out a nice, friendly ICS file.

I'm going to break the script apart and explain the pieces. At the end
will be the entire script.

{% codeblock Script arguments and boilerplate - EWStoICS.ps1 lang:posh %}
param([Parameter(Mandatory=$true)]$calendarName, [Parameter(Mandatory=$true)]$destinationFile)

Set-StrictMode -Version 2

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Push-Location $ScriptDir
[Environment]::CurrentDirectory = $PWD

Import-Module -name "C:\Program Files\Microsoft\Exchange\Web Services\1.2\Microsoft.Exchange.WebServices.dll"
[Reflection.Assembly]::LoadFile("$ScriptDir\DDay.iCal.dll")

[Reflection.Assembly]::LoadWithPartialName("System.DirectoryServices.AccountManagement")
{% endcodeblock %}
First up, the script arguments. We'll need the name of the calendar and the file we're going to write to.
I'm going ahead and setting strict mode because Powershell version 3 is coming soon, and I don't want to get
surprised by syntax and behavior changes. I'm also going to change the working directory to the script's dir.
After that, we'll load the EWS Managed API from where it gets installed, DDay.iCal from the local dir, and
the System.DrectoryServices.AccountManagement class.
{% codeblock Connecting to Exchange Web Services - EWStoICS.ps1 %}
# Get Current User's E-mail address the powershell script assumes that it is being run as the user with access
$MailboxAddress = [System.DirectoryServices.AccountManagement.UserPrincipal]::Current.EmailAddress

# Create service to talk to Exchange
$exchService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService
$exchService.UseDefaultCredentials = $true

# Connect
$exchService.AutodiscoverUrl($MailboxAddress)
{% endcodeblock %}
In order to talk to EWS we need credentials and a URL. In the onterest of simplicity, we're going to use the
account that the script is running under. We'll also use autodiscover and the account's e-mail address to find
the URL for EWS. I prefer to do it this way, but you can use any PSCredential to connect.
{% codeblock EWStoICS.ps1 %}
# Get the Root Mailbox folder
$MailboxRootId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot)
$MailboxRoot = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchService, $MailboxRootId)
{% endcodeblock %}
Now we get root folder of the Mailbox. We'll use this to search for the calendar
folders.
{% codeblock Find the Calendar - EWStoICS.ps1 %}
#Get the ID of the Calendar Folder
$fvFolderView =  New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)  
$fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep;  
$SfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,$calendarName)  
$findFolderResults = $MailboxRoot.FindFolders($SfSearchFilter,$fvFolderView)  

$exchCalendarFolder = $findFolderResults.Folders[0]
{% endcodeblock %}
Now we find the calendar that was requested. Note: I'm being really lazy here,
and I'm not checking to make sure that I found the calendar or that I didn't
find more than one. Using the script as is requires that there only be one
folder with that name.
{% codeblock EWStoICS.ps1 %}

$startDate = New-Object System.DateTime(2012, 05, 01)
$endDate = New-Object System.DateTime(2013, 07, 31)

$CalendarView = New-Object Microsoft.Exchange.WebServices.Data.CalendarView($startDate, $endDate)

$fiItems = $exchService.FindAppointments($exchCalendarFolder.Id, $CalendarView)
{% endcodeblock %}
Exchange will do the heavy lifting of enumerating all appointment occurences, but we're limited to a 2 year range.
{% codeblock EWStoICS.ps1 %}

$iCalendar = New-Object DDay.iCal.iCalendar

$iCalendar.AddProperty("X-WR-CALNAME",$calendarName)

{% endcodeblock %}
Creating the DDay.iCal Calendar object and naming it based on the calendar name
{% codeblock EWStoICS.ps1 %}

foreach ($exchangeEvent in $fiItems) {
    // stuff ...
}

{% endcodeblock %}
The main loop. Goes through every item that we pulled and creates the pieces for the iCal.
{% codeblock EWStoICS.ps1 %}
    $iCalEvent = New-Object DDay.iCal.Event
        
    $iCalEvent.UID = $exchangeEvent.iCalUID
    $iCalEvent.Created = New-Object DDay.iCal.iCalDateTime($exchangeEvent.ICalDateTimeStamp.ToUniversalTime())
    $iCalEvent.Start = New-Object DDay.iCal.iCalDateTime($exchangeEvent.Start.ToUniversalTime())
    $iCalEvent.End = New-Object DDay.iCal.iCalDateTime($exchangeEvent.End.ToUniversalTime())
    $iCalEvent.LastModified = New-Object DDay.iCal.iCalDateTime($exchangeEvent.LastModifiedTime.ToUniversalTime())
    $iCalEvent.Location = $exchangeEvent.Location
    $iCalEvent.Summary = $exchangeEvent.Subject
    $iCalEvent.Description = $exchangeEvent.Body
{% endcodeblock %}
The basic things that need to be in every event. UID, start and end, last modified, location, summary, and description. The only interesting thing is that I've converted the time to UTC instead of dealing with timezones. This also makes everything work with Google Calendar.
{% codeblock EWStoICS.ps1 %}
    switch($exchangeEvent.Importance) {
        "Normal" { $iCalEvent.Priority = 5 }
        "High" { $iCalEvent.Priority = 1 }
        "Low" { $iCalEvent.Priority = 9 }
        default { $iCalEvent.Priority = 0 }
    }
{% endcodeblock %}
I'm following how Outlook does iCal in converting priorities.
{% codeblock EWStoICS.ps1 %}
    switch($exchangeEvent.LegacyFreeBusyStatus) {
        "Busy" { $iCalEvent.AddProperty("TRANSP", "OPAQUE") }
        "Free" { $iCalEvent.AddProperty("TRANSP", "TRANSPARENT") } 
        default {}
    }
{% endcodeblock %}
Again, I'm following how Outlook's iCal conversion does things. I'm not dealing with tentative, because the calendars I'm converting shouldn't have it.
{% codeblock EWStoICS.ps1 %}    
    $iCalEvent.Class = "PUBLIC"
{% endcodeblock %}
Our calendars will all be public, so there's no need for a private event.
{% codeblock EWStoICS.ps1 %}    
    if ($exchangeEvent.AppointmentType -eq "Occurrence") {
        # Determine if this is the first time we've seen this event
        if ($iCalendar.Events.Count -ne 0) {
            $matchingUIDs = $iCalendar.Events | Where-Object { $_.UID -eq $exchangeEvent.ICalUid }
            
            if ($matchingUIDs -ne $null) {
                break; 
            }
        }
        
        # Get Recurrence
        $exchangeEventRecurrenceMaster = [Microsoft.Exchange.WebServices.Data.Appointment]::BindToRecurringMaster($exchService, $exchangeEvent.Id)
        $exchangeRecurrence = $exchangeEventRecurrenceMaster.Recurrence
        
        # Create iCal Recurrence
        $iCalRecurrence = New-Object DDay.ICal.RecurrencePattern
        
        if ($exchangeRecurrence.NumberOfOccurrences -ne $null) {
            $iCalRecurrence.Count = $exchangeRecurrence.NumberOfOccurrences
        }
        
        $recurrenceType = $exchangeRecurrence.GetType()
        if ($recurrenceType.Name -eq "WeeklyPattern") {
        
            $iCalRecurrence.Frequency = "Weekly"
            if ($exchangeRecurrence.FirstDayOfWeek) {
                $iCalRecurrence.FirstDayOfWeek = $exchangeRecurrence.FirstDayOfWeek
            }
            if ($exchangeRecurrence.DaysOfTheWeek.Count -ne 0) {
                foreach ($day in $exchangeRecurrence.DaysOfTheWeek) {
                    $icalDay = New-Object DDay.iCal.WeekDay($day.ToString())
                    $iCalRecurrence.ByDay.Add($iCalDay)
                }
            }
        } elseif ($recurrenceType.Name -eq "DailyPattern") {
            $iCalRecurrence.Frequency = "Daily"
        }
        
        # Add Recurrence to iCal Event
        $icalEvent.RecurrenceRules.Add($icalRecurrence)
        
    }
{% endcodeblock %}
Recurring events are tricky. This code only covers Weekly or Daily recurrence. Even though most of the calendars I'm converting don't have recurrence, it does make setting up calendars easier. I didn't deal with monthly or yearly recurrence because of the time scale of these calendars.
{% codeblock EWStoICS.ps1 %}
 elseif ($exchangeEvent.AppointmentType -eq "Exception") {
        
        $iCalEvent.RecurrenceID = New-Object DDay.iCal.iCalDateTime($exchangeEvent.ICalRecurrenceId.ToUniversalTime())
        
        $iCalendar.AddChild($iCalEvent)
        break;
    }
{% endcodeblock %}
If you do recurrence, you have to be prepared for recurring events that get moved. For example, my boss moves almost every bi-weekly team meeting because of conflicts in his schedule.
{% codeblock EWStoICS.ps1 %}
    $iCalEvent.IsAllDay = $exchangeEvent.IsAllDayEvent
{% endcodeblock %}
Yes, we can do All Day Events, too!
{% codeblock EWStoICS.ps1 %}    
    if ($exchangeEvent.Categories.Count -ne 0) {
        foreach ($category in $exchangeEvent.Categories) {
            $iCalEvent.Categories.Add($category)
        }
    }
{% endcodeblock %}
Add category information. This is pretty important for Outlook users, but not as much for iCal or Google Calendar users.
{% codeblock EWStoICS.ps1 %}    
    $iCalendar.AddChild($iCalEvent)
{% endcodeblock %}
Actually add the event to the calendar. This is the last step of the loop that goes through events.
{% codeblock EWStoICS.ps1 %}
$iCalSerializer = New-Object DDay.iCal.Serialization.iCalendar.iCalendarSerializer

$iCalSerializer.Serialize($icalendar, $destinationFile)

{% endcodeblock %}
This is the part the actually contructs the iCal file. Figuing our the Serializer object to use was harder than it should be because the DDay.iCal documentation that I could find isn't up-to-date.

And that's the script. Powershell really begins to shine if you can learn the .NET Framework and pull in it in. You do some pretty awesome things with `System.Net.WebClient` that makes things almost as good as having `curl` or `wget` available.

Below is the entire script:

{% codeblock EWStoICS.ps1 %}
param([Parameter(Mandatory=$true)]$calendarName, [Parameter(Mandatory=$true)]$destinationFile)

Set-StrictMode -Version 2

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Push-Location $ScriptDir
[Environment]::CurrentDirectory = $PWD

Import-Module -name "C:\Program Files\Microsoft\Exchange\Web Services\1.2\Microsoft.Exchange.WebServices.dll"
[Reflection.Assembly]::LoadFile("$ScriptDir\DDay.iCal.dll")

[Reflection.Assembly]::LoadWithPartialName("System.DirectoryServices.AccountManagement")

# Get Current User's E-mail address the powershell script assumes that it is being run as the user with access
$MailboxAddress = [System.DirectoryServices.AccountManagement.UserPrincipal]::Current.EmailAddress

# Create service to talk to Exchange
$exchService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService
$exchService.UseDefaultCredentials = $true

# Connect
$exchService.AutodiscoverUrl($MailboxAddress)

# Get the Root Mailbox folder
$MailboxRootId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot)
$MailboxRoot = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchService, $MailboxRootId)

#Get the ID of the Calendar Folder
$fvFolderView =  New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)  
$fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep;  
$SfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,$calendarName)  
$findFolderResults = $MailboxRoot.FindFolders($SfSearchFilter,$fvFolderView)  

$exchCalendarFolder = $findFolderResults.Folders[0]

$startDate = New-Object System.DateTime(2012, 05, 01)
$endDate = New-Object System.DateTime(2013, 07, 31)

$CalendarView = New-Object Microsoft.Exchange.WebServices.Data.CalendarView($startDate, $endDate)

$fiItems = $exchService.FindAppointments($exchCalendarFolder.Id, $CalendarView)


$iCalendar = New-Object DDay.iCal.iCalendar

$iCalendar.AddProperty("X-WR-CALNAME",$calendarName)

foreach ($exchangeEvent in $fiItems) {
    $iCalEvent = New-Object DDay.iCal.Event
        
    $iCalEvent.UID = $exchangeEvent.iCalUID
    $iCalEvent.Created = New-Object DDay.iCal.iCalDateTime($exchangeEvent.ICalDateTimeStamp.ToUniversalTime())
    $iCalEvent.Start = New-Object DDay.iCal.iCalDateTime($exchangeEvent.Start.ToUniversalTime())
    $iCalEvent.End = New-Object DDay.iCal.iCalDateTime($exchangeEvent.End.ToUniversalTime())
    $iCalEvent.LastModified = New-Object DDay.iCal.iCalDateTime($exchangeEvent.LastModifiedTime.ToUniversalTime())
    $iCalEvent.Location = $exchangeEvent.Location
    $iCalEvent.Summary = $exchangeEvent.Subject
    $iCalEvent.Description = $exchangeEvent.Body

    switch($exchangeEvent.Importance) {
        "Normal" { $iCalEvent.Priority = 5 }
        "High" { $iCalEvent.Priority = 1 }
        "Low" { $iCalEvent.Priority = 9 }
        default { $iCalEvent.Priority = 0 }
    }
    
    switch($exchangeEvent.LegacyFreeBusyStatus) {
        "Busy" { $iCalEvent.AddProperty("TRANSP", "OPAQUE") }
        "Free" { $iCalEvent.AddProperty("TRANSP", "TRANSPARENT") } 
        default {}
    }
    
    $iCalEvent.Class = "PUBLIC"
    
    if ($exchangeEvent.AppointmentType -eq "Occurrence") {
        # Determine if this is the first time we've seen this event
        if ($iCalendar.Events.Count -ne 0) {
            $matchingUIDs = $iCalendar.Events | Where-Object { $_.UID -eq $exchangeEvent.ICalUid }
            
            if ($matchingUIDs -ne $null) {
                break; 
            }
        }
        
        # Get Recurrence
        $exchangeEventRecurrenceMaster = [Microsoft.Exchange.WebServices.Data.Appointment]::BindToRecurringMaster($exchService, $exchangeEvent.Id)
        $exchangeRecurrence = $exchangeEventRecurrenceMaster.Recurrence
        
        # Create iCal Recurrence
        $iCalRecurrence = New-Object DDay.ICal.RecurrencePattern
        
        if ($exchangeRecurrence.NumberOfOccurrences -ne $null) {
            $iCalRecurrence.Count = $exchangeRecurrence.NumberOfOccurrences
        }
        
        $recurrenceType = $exchangeRecurrence.GetType()
        if ($recurrenceType.Name -eq "WeeklyPattern") {
        
            $iCalRecurrence.Frequency = "Weekly"
            if ($exchangeRecurrence.FirstDayOfWeek) {
                $iCalRecurrence.FirstDayOfWeek = $exchangeRecurrence.FirstDayOfWeek
            }
            if ($exchangeRecurrence.DaysOfTheWeek.Count -ne 0) {
                foreach ($day in $exchangeRecurrence.DaysOfTheWeek) {
                    $icalDay = New-Object DDay.iCal.WeekDay($day.ToString())
                    $iCalRecurrence.ByDay.Add($iCalDay)
                }
            }
        } elseif ($recurrenceType.Name -eq "DailyPattern") {
            $iCalRecurrence.Frequency = "Daily"
        }
        
        # Add Recurrence to iCal Event
        $icalEvent.RecurrenceRules.Add($icalRecurrence)
        
    } elseif ($exchangeEvent.AppointmentType -eq "Exception") {
        
        $iCalEvent.RecurrenceID = New-Object DDay.iCal.iCalDateTime($exchangeEvent.ICalRecurrenceId.ToUniversalTime())
        
        $iCalendar.AddChild($iCalEvent)
        break;
    }

    $iCalEvent.IsAllDay = $exchangeEvent.IsAllDayEvent
    
    if ($exchangeEvent.Categories.Count -ne 0) {
        foreach ($category in $exchangeEvent.Categories) {
            $iCalEvent.Categories.Add($category)
        }
    }
    
    $iCalendar.AddChild($iCalEvent)
}

$iCalSerializer = New-Object DDay.iCal.Serialization.iCalendar.iCalendarSerializer

$iCalSerializer.Serialize($icalendar, $destinationFile)

{% endcodeblock %}
