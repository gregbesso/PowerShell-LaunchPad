## Synopsis

The purpose of this PowerShell script is to provide an easy way for team members to bring up one or multiple 
PowerShell scripting tools without having to know the location or parameters for any of them. Also it makes it easier to 
setup one shortcut that can be configured and refreshed without any further changes by anyone whenever scripts are updated 
or added.

## Code Example

One way to bring up this tool to avoid having to use the command line is to create a Windows shortcut to PowerShell.exe 
and then call the LaunchPad.ps1 file...

C:\WINDOWS\System32\WindowsPowerShell\v1.0\powershell.exe -File "\\server\share\scripts\PowerShell-LaunchPad.ps1"

*If you need to bypass a local workstation's execution policy, you could include the following in the shortcut target...
-ExecutionPolicy Bypass 

The LaunchPad.ps1 script looks in the $global:lpLaunchFiles directory and pulls up any .ps1 files to populate the drop-down 
with choices.

## Motivation

I found that the manual steps involved with sharing and launching PowerShell scripts on other computers tended to be a PITA 
and I wanted an easier way to share tools as I was making them without having to keep updating people with more manual steps 
that they needed. I thought a "set it once and forget about it" tactic would work well here, so came up with the idea of making 
this "Launch Pad" tool.

Since I have used it, it has proven to work as expected, and any changes I make to the underlying launch pad itself, or to the 
scripts it is linking to, has gone seamlesslessly and unnoticed. Less time maintaining the user settings equals more time spent 
learning and coding tools :-)

## Installation

This is just a stand-alone PowerShell script.

## API Reference

No API here. <sounds of crickets>

## Tests

No testing info here. <sounds of crickets>

## Contributors

Just a solo script project by moi, Greg Besso. Hi there :-)

## License

Copyright (c) 2015 Greg Besso

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.