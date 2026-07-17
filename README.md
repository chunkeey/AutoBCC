# AutoBCC
Automatically set a custom defined email address as bcc recipient when answering or creating new emails.

## Credits / Copyright
This was forked from [SentOnBehalfBCC by urbans0ft](https://github.com/urbans0ft/SentOnBehalfBCC). This was
already 99% done. Unfortunately the author didn't add any LICENSE. All there is, is a Copyright with a Date
in the [AssemblyInfo.cs](Properties/AssemblyInfo.cs). So, this is as is at it gets. Don't be surprised if this
all disappears. I myself have no problem if this was just Public Domain.

## Building
Just download and install Visual Studio. I used Visual Studio Community 2022 and it basically did everything
once I opened the [AutoBCC.csproj](AutoBCC.csproj) file. Just follow the AIs guidance and either Build & Debug
it (starts Outlook automatically) or Publish it to get an installer.

## Usage
Works here with Outlook Classic Application from regular Microsoft Office 2024 LTSC. Once installed, a new
menu item called "Auto BCC" should show up in the menu/ribbon bar. Press it and fill in the E-Mail-Address
you want. Make sure that the "Ein/Aus" toggle is appropriately set. And from then on, every mail you 
sent/forward/reply has the BCC field automatically set. But be aware, this is local! If you need it 
on a different PC or account (even on the same PC!), you'll have to set it up there as well in the same way.
