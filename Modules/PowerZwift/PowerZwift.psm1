# important paths you'll need to know
# $home\Documents\Zwift
# $home\Pictures\Zwift
#  ${env:ProgramFiles(x86)}\Zwift

<#
.Synopsis
   Sets your Zwift preferences to the nominated course
.DESCRIPTION
   Choose from

   Watopia
   Richmond
   London
   Default
.EXAMPLE
   Set-ZwiftCourse -course London
.EXAMPLE
   Set-ZwiftCourse -course Default
.FUNCTIONALITY
   The functionality that best describes this cmdlet
#>
Function Set-ZwiftCourse
{
    [Cmdletbinding()]
    param
    (
        [string]
        [ValidateSet('Watopia', 'Richmond', 'London', 'Default')]
        [Parameter(Mandatory=$true)]
        $course
    )

    # $home\Documents\Zwift\prefs.xml 
    # http://zwiftblog.com/world-tag/

    if(!(Test-Path $home\Documents\Zwift\prefs.xml))
    {
        Write-Error "ZWIFT PREFS FILE NOT FOUND"
        throw;
    }
    
    $prefs = [xml](gc $home\Documents\Zwift\prefs.xml)

    if($course -eq 'Default')
    {
        # delete the WORLD node
        if($prefs.ZWIFT.WORLD -eq $null)
        {
            # nothing to do here. We're already default
        }
        else
        {
            $worldnode = $prefs.ZWIFT.selectSingleNode("WORLD")
            $prefs.ZWIFT.RemoveChild($worldnode) | out-null
            Write-Output "Prefs set to Default"
        }
    }
    else
    {
        # enum this later, because switches are a bit ugly
        switch($course)
        {
            "Watopia" { $coursenum = "1"}         
            "Richmond" { $coursenum = "2"}         
            "London" { $coursenum = "3"} 
        }
        if($prefs.ZWIFT.WORLD -eq $null)
        {
            # we don't currently have a world element
            $node = $prefs.CreateElement("WORLD")
            $node.InnerXml = $coursenum  
            $prefs.ZWIFT.AppendChild($node) | out-null
        }
        else
        {   # we have a world element, set it and save it
            $prefs.ZWIFT.WORLD = $coursenum

        }   
        Write-Output "Prefs set to $course"
    }    
    $prefs.Save("$home\Documents\Zwift\prefs.xml") | out-null
}

Function Get-ZwiftCourse
{
    [cmdletbinding()]
    param()
    # if we have a <world> tag, return that value.
    # if not, return "default"
    if(!(Test-Path $home\Documents\Zwift\prefs.xml))
    {
        Write-Error "ZWIFT PREFS FILE NOT FOUND"
        throw;
    }

    $prefs = [xml](gc $home\Documents\Zwift\prefs.xml)

    if($prefs.ZWIFT.WORLD -eq $null)
    {
        # nothing to do here. We're already default
        return "default"
    }
    else
    {
        return $prefs.ZWIFT.WORLD
    }
}

Function Get-ZwiftPreferences
{
    [cmdletbinding()]
    param()
    $prefs = [xml](gc $home\Documents\Zwift\prefs.xml)
    return $prefs 
    # not yet fully implemented
}

Function Set-ZwiftPreferences
{
    [cmdletbinding()]
    param()
    # not yet fully implemented
}

<#
.SYNOPSIS
    Gives a list of activities stored in your \Zwift folder and their basic parameters
#>
Function Get-ZwiftActivities
{
    [cmdletbinding()]
    param()
    # $home\Documents\Zwift\Activities 
    # requires https://github.com/jflam/FastFitParser 

    # not yet implemented
}

<#
.SYNOPSIS
    Gets Zwift events and returns them as powershell objects, so you can filter down and build alerting based on them
#>
Function Get-ZwiftEvent
{
    [CmdletBinding()]
    param
    (
        $eventID
    )

    # needs a little enhancement
    
    if($eventID)
    {
        return Invoke-RestMethod https://zwift.com/json/events/$eventID
    }
    else
    {
        return Invoke-RestMethod https://zwift.com/json/events 
    }
}

Function Install-Zwift
{
    [cmdletbinding()]
    param()
    Invoke-WebRequest https://zwift.com/download/pc -outFile $env:tmp\ZwiftInstaller.exe
    & $env:tmp\ZwiftInstaller.exe
}

<#
.Synopsis
    Pulls the latest ZwiftMap executable and installs it
.EXAMPLE
    Install-ZwiftMap -verbose
.COMPONENT
    PowerZwift - http://github.com/stopthatastronaut/PowerZwift
#>
Function Install-ZwiftMap
{
    [cmdletbinding()]
    param()
    # http://zwifthacks.com/download/zwiftmap-portable/ 
    # needs powershell 5.0 and/or pscx to do the unzip
    Invoke-WebRequest http://zwifthacks.com/download/zwiftmap-portable/ -outfile $home\Documents\Zwift\zwiftmap-portable.zip
    Expand-Archive -path $home\Documents\Zwift\zwiftmap-portable.zip -DestinationPath $home\Documents\Zwift\Scripts -Force
}

<#
.Synopsis
    Reads ZwiftMap config out into an object which can be manipulated
.EXAMPLE
    $config = Get-ZwiftMapConfig
    $config.SettingsMinimizeOnStart = 1
    $config | Write-ZwiftMapConfig
.EXAMPLE
    $config = Get-ZwiftMapConfig
    $config.SettingsMinimizeOnStart = 1
    $config.SettingsSaveShowLog = 1
    $config.SettingsSavePosition = 0
    Write-ZwiftMapConfig -InputObject $config
.INPUTS
    An object as derived from Get ZwiftMapSettings
.COMPONENT
    PowerZwift - http://github.com/stopthatastronaut/PowerZwift
#>
Function Get-ZwiftMapConfig
{
    [CmdletBinding()]
    param()
    # find the file
    if(Test-Path $home\Documents\Zwift\Scripts\ZwiftMap.exe)
    {
        $inipath = gci $home\Documents\Zwift\Scripts | ? {$_.Name -like "ZwiftMap*" -and $_.Name -like "*.ini"} | select -expand FullName
        $inifile = gc $inipath
        $obj = [pscustomobject]@{}

        $inifile | % {
            if($_[0] -ne "[")
            {
                $val = $_ -split "="
                $obj | Add-Member -Name $val[0] -Value $val[1] -Type NoteProperty
            }
        }
        return $obj
    }
    else
    {
        Write-Error "ZwiftMap doesn't appear to be installed where expected"
        throw
    }
}


<#
.Synopsis
    Writes ZwiftMap config out to the default file
.DESCRIPTION
    Long description
.EXAMPLE
    $config = Get-ZwiftMapConfig
    $config.SettingsMinimizeOnStart = 1
    $config | Write-ZwiftMapConfig
.EXAMPLE
    $config = Get-ZwiftMapConfig
    $config.SettingsMinimizeOnStart = 1
    $config.SettingsSaveShowLog = 1
    $config.SettingsSavePosition = 0
    Write-ZwiftMapConfig -InputObject $config
.INPUTS
    An object as derived from Get ZwiftMapSettings
.COMPONENT
    PowerZwift - http://github.com/stopthatastronaut/PowerZwift
#>
Function Write-ZwiftMapConfig
{
    [cmdletbinding()]
    param
    (
        [Parameter(ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true)]
        $inputObject
    )

    $inipath = gci $home\Documents\Zwift\Scripts | ? {$_.Name -like "ZwiftMap*" -and $_.Name -like "*.ini"} | select -expand FullName
    "[Settings]" | out-file $inipath -Force

    Write-Verbose "Writing $_ to output stream" 

    $InputObject | gm | ? {$_.MemberType -eq "NoteProperty"} | % {
        $propertyname = $_.Name
        $propertyvalue = $inputObject."$propertyname"

        Write-Verbose "Writing $propertyname with value $propertyvalue"
        "$propertyname=$propertyvalue" | out-file $inipath -Append -Force
    }   
}

<#
.Synopsis
   Invokes the Zwift executable, with optional third-party extras
.DESCRIPTION
   Provided to automate starting Zwift with optional extras running. Notably, running Zwiftmap at the same time as running Zwift
.EXAMPLE
   Invoke-Zwift
.EXAMPLE
   Invoke-Zwift -WithMap
.COMPONENT
   PowerZwift - http://github.com/stopthatastronaut/PowerZwift
#>
Function Invoke-Zwift
{
    # starts zwift up
    # http://zwifthacks.com/zwiftmap/
    [CmdletBinding()]
    param
    (
        [switch]
        $WithMap
    )

    & "C:\Program Files (x86)\Zwift\ZwiftLauncher.exe" 
    Write-Verbose "Invoked Zwift Executable"

    if($withmap)
    {
        # find ZwiftMap and run it
        if(test-Path $home\Documents\Zwift\Scripts\ZwiftMap.exe)
        {
            & $home\Documents\Zwift\Scripts\ZwiftMap.exe 
            Write-Verbose "Invoked ZwiftMap Executable"
        }
        else
        {
            Write-Warning "ZwiftMap not found. Maybe try running the Install-ZwiftMap command?"
        }
    }
}

<#
.Synopsis
   Turns off the startup music
.DESCRIPTION
   Annoyed by the default startup music? Turn it off here
.EXAMPLE
   Disable-ZwiftStartupMusic
.EXAMPLE
   Disable-ZwiftStartupMusic -Verbose
.COMPONENT
   PowerZwift - http://github.com/stopthatastronaut/PowerZwift
#>
Function Disable-ZwiftStartupMusic
{
    [cmdletbinding()]
    param()
    # http://zwiftblog.com/turn-off-zwifts-startup-music-easy-hack/
    if(Test-Path 'C:\Program Files (x86)\Zwift\data\Audio\PC\777200017.wem')
    {
        Write-Verbose "Found .wem file at C:\Program Files (x86)\Zwift\data\Audio\PC\777200017.wem"
        Rename-Item 'C:\Program Files (x86)\Zwift\data\Audio\PC\777200017.wem' '777200017.tmp'
        Write-Warning "Opening Music disabled"
    }
    else
    {
        if(Test-Path 'C:\Program Files (x86)\Zwift\data\Audio\PC\777200017.tmp')
        {
            Write-Warning "Opening music has already been disabled"
        }
        else
        {
            Write-Error "Could not find Zwift music file"
        }
    }
}

<#
.Synopsis
   Turns on the startup music
.DESCRIPTION
   Miss the opening music? Turn it back on here
.EXAMPLE
   Disable-ZwiftStartupMusic
.EXAMPLE
   Disable-ZwiftStartupMusic -Verbose
.COMPONENT
   PowerZwift - http://github.com/stopthatastronaut/PowerZwift
#>
Function Enable-ZwiftStartupMusic
{
    [cmdletbinding()]
    param()
    # http://zwiftblog.com/turn-off-zwifts-startup-music-easy-hack/ 
        # http://zwiftblog.com/turn-off-zwifts-startup-music-easy-hack/
    if(Test-Path 'C:\Program Files (x86)\Zwift\data\Audio\PC\777200017.tmp')
    {
        Write-Verbose "Found .tmp file at C:\Program Files (x86)\Zwift\data\Audio\PC\777200017.tmp"
        Rename-Item 'C:\Program Files (x86)\Zwift\data\Audio\PC\777200017.tmp' '777200017.wem'
        Write-Warning "Opening Music enabled"
    }
    else
    {
        if(Test-Path 'C:\Program Files (x86)\Zwift\data\Audio\PC\777200017.wem')
        {
            Write-Warning "Opening music has already been enabled"
        }
        else
        {
            Write-Error "Could not find Zwift music file"
        }
    }
}

Function Update-PowerZwift
{
    # updates this module to the latest release found on github
    # not yet implemented
}

<#
.Synopsis
   Finds a user's Zwift ID
.DESCRIPTION
   Finds a user's Zwift ID. This is useful for third party signups like Zwift.Community, and if you're a racer, you'll want to know this
.EXAMPLE
   PS:> Get-ZwiftId
   123456
.COMPONENT
   PowerZwift - http://github.com/stopthatastronaut/PowerZwift
#>
Function Get-ZwiftId
{
    $foldername = gci $home\Documents\Zwift\cp\ | ? { $_.Name -like "user*" } | select -first 1
    return $foldername -replace "user", ""
}

<#
.Synopsis
   Puts a new clickable shortcut on your desktop, with your preferences
.DESCRIPTION
    Want a shortcut that takes you direct to your chosen course? Use this. It'll place a new shortcut on your desktop. Feel free to drag that wherever you need it.
.EXAMPLE
   New-ZwiftShortcut -course Richmond -Withmap
.COMPONENT
   PowerZwift - http://github.com/stopthatastronaut/PowerZwift
#>
Function New-ZwiftShortcut
{
    [CmdletBinding()]
    param
    (
        [string]
        [ValidateSet('Watopia', 'Richmond', 'London', 'Default')]
        [Parameter(Mandatory=$true)]
        $course,
        [switch]
        $withMap
    )

    $m = ""
    if($withMap)
    {
        $m = "-WithMap"
    }

    $TargetFile = "powershell.exe"
    $Arguments = "-NoProfile -Windowstyle Minimized -NoExit -command `"ipmo PowerZwift; Set-ZwiftCourse -course $course; Invoke-Zwift $m -verbose`""
    $ShortcutFile = "$home\Desktop\Zwift - $course.lnk"
    $WScriptShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
    $Shortcut.TargetPath = $TargetFile
    $Shortcut.Arguments = $Arguments
    $Shortcut.iconlocation = "C:\Program Files (x86)\Zwift\ZwiftLauncher.exe"
    $Shortcut.Save()

}

Function Get-ZwiftWorkout   # at the moment, only lists what you have. In future: will list what you have and allow you to drill down into a specific workout
{
    [CmdletBinding()]
    param
    (
        
    )
    # lists workouts
    if(-not (Test-Path $home\Documents\Zwift\Workouts))
    {
        throw "Workouts folder not found"
    }
    else
    {
        $w = (gci $home\Documents\Zwift\Workouts -filter "*.zwo")

        if($w -eq $null)
        {
            throw "No custom workouts found"
        }
        else
        {
            $w | %  {
                $wxml = [xml](gc $_.FullName)

                # get the intervals out of it

                # return it straight to the pipeline for processing
                return [pscustomobject]@{Name = $wxml.workout_file.name; Author = $wxml.workout_file.author; Description = $wxml.workout_file.description; Workout = $wxml.workout_file.workout.SteadyState}
            }
        }
    }
}

Function New-ZwiftWorkout
{
    [CmdletBinding()]
    param
    (
        $Name,
        $WorkoutType,
        $Author,
        $Description,
        [switch]
        $Random
    )

    # not yet implemented


}

FUnction Set-EventBackGround
{
    # let's grab all the images off the event feed, overlay the event details on it, then desktop background it


    # not implemented

    (irm https://zwift.com/json/events).GetEnumerator() | select -expand imageUrl
}