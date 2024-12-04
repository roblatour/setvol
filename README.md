# SetVol (v4.3 December 4, 2024)
SetVol is a free open source command line program for Windows to change and report on your computer's audio and recording levels.  It also allows you to set your default audio and recording devices.

# License
SetVol is licensed under a [MIT license](https://github.com/roblatour/setvol/blob/main/LICENSE)

## Download 

You are welcome to download and use it for free on as many computers as you would like.

A downloadable zipped signed executable of the program is available from [here](https://6ec1f0a2f74d4d0c2019-591364a760543a57f40bab2c37672676.ssl.cf5.rackcdn.com/SetVol.zip).


## Setup

SetVol does not need to be installed, rather just unzip it from the download file (above) and run it.

## Using SetVol

From the command prompt just type "setvol ?" (without the quotes) to see the help (as shown below).

**SetVol Help screen** 
![SetVol screenshot](/images/screenshot.jpg) 

## Additional information

SetVol ignores the (upper/lower) case used in words

For example the following both set the volume level to 10 percent:   
    setvol ten  
    setvol TEN

SetVol ignores the "%" symbol and the word "percent"

For example the following all set the volume level to 25 percent:  
    setvol 25  
    setvol 25 %  
    setvol twenty-five  
    setvol twenty-five percent

SetVol doesn't care if there is a hyphen between numbered words

For example the following both set the volume level to 35 percent:   
    setvol thirty-five  
    setvol thirty five

You can use the balance, beep and report options together to verify individual speakers

For example on an eight channel device you can use the following commands:

setvol 50 balance 100:0:0:0:0:0:0:0 beep report

setvol 50 balance 0:100:0:0:0:0:0:0 beep report

...

setvol 50 balance 0:0:0:0:0:0:0:100 beep report

Here is an example of how to call SetVol via a .bat file and have it return the current volume:

call setvol report

echo %ERRORLEVEL%

Using SetVol with Windows PowerShell:

Unlike the Windows command line, Windows PowerShell requires that special characters be escaped.  The is done with the '\`' character which is often found on the keyboard on the same key as the tilda character '~'.

So for example with the Windows command line where you might enter:

    setvol report device BenQ BL3200 (NVIDIA High Definition Audio)

with Windows PowerShell you would enter:

   .\\setvol report device BenQ BL3200 \`(NVIDIA High Definition Audio\`)

where the open and closed brackets have been escaped. 

Other ways to get help:

Here are two additional ways to view the help   
    setvol  
    setvol help


## Components

The project is built upon the .net framework 4.7.2 and includes the following:

![components](/images/components.jpg)

Thanks also to IronRazerz for a significant portion of the 'makedefault' functionality as excerpted from:
https://social.msdn.microsoft.com/Forums/windowsdesktop/en-US/3c437708-0e90-483d-b906-63282ddd2d7b/change-audio-input

* * *
 ## Support SetVol

 To help support SetVol, or to just say thanks, you're welcome to 'buy me a coffee'<br><br>
[<img alt="buy me  a coffee" width="200px" src="https://cdn.buymeacoffee.com/buttons/v2/default-blue.png" />](https://www.buymeacoffee.com/roblatour)
* * *
Copyright © 2024, Rob Latour
* * *
