# PowerZwift
PowerShell module for Zwift hacks. Allows you to automate your Zwifting on Windows for maximum wattage.

## Highlights

`Set-ZwiftCourse -course [Watopia|Richmond|London|Default]` allows you to override the default Zwift calendar and select the course YOU want. Of course, you may be a little lonely, but sometimes a rider needs to be alone, right?

`Install-ZwiftMap` automates instalation (and updating) of the ZwiftMap unofficial add-on

`Disable-ZwiftStartupMusic` and `Enable-ZwiftStartupMusic` allow you to toggle the startup music on and off

`Invoke-Zwift -withmap` is a gateway to automatically starting Zwift with all kinds of addons and option toggles

There are more cmdlets coming.

### Find out more

`PS:>Get-Help about_PowerZwift`

## Requirements

- Zwift for Windows, naturally
- Windows 7 or later
- Windows PowerShell 5.0 recommended - though most features will work with 4.0
- A willingness to occasionally type something

## Installation

Take the `PowerZwift` subfolder from this repo's `Modules` folder and copy it (or junction it) to `C:\Program Files\WindowsPowerShell\Modules\PowerZwift`

I recommend the junction point method, as you can just point it to a git clone of this repo to update PowerZwift quickly

`junction.exe` can be found in the Sysinternals tools suite. To use that method, run

`junction.exe C:\Program Files\WindowsPowerShell\Modules\PowerZwift <Your Cloned Repo>\Modules\PowerZwift`

## Getting Started

Fire up the powershell command window

Type `Import-Module PowerZwift`

Start hacking away!

##Who Why Where

*Jason Brown KoS*, Zwifter, Knight of Sufferlandria, MTB racer, coach-in-training and Windows automation engineer.

I spend most of my working day in a PowerShell environment. When I get home and want to mess around with Zwift, it makes sense to do it the same way.

Sydney, Australia

## Disclaimers

This module is not directly affiliated with Zwift Inc. You use it at your own risk. Etc etc.

PRs gratefully accepted
