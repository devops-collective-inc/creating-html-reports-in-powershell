# HTML Report Basics

First, understand that PowerShell isn't limited to creating reports in HTML. But I like HTML because it's flexible, can be easily e-mailed, and can be more easily made to look pretty than a plain-text report. But before you dive in, you do need to know a bit about how HTML works.

An HTML page is just a plain text file, looking something like this:

```html
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"  "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>HTML TABLE</title>
</head><body>
<table>
<colgroup><col/><col/><col/><col/><col/></colgroup>
<tr><th>ComputerName</th><th>Drive</th><th>Free(GB)</th><th>Free(%)</th><th>Size(GB)</th></tr>
<tr><td>CLIENT</td><td>C:</td><td>49</td><td>82</td><td>60</td></tr>
</table>
</body></html>
```

When read by a browser, this file is rendered into the display you see within the browser's window. The same applies to e-mail clients capable of displaying HTML content. While you, as a person, can obviously put anything you want into the file, if you want the output to look right you need to follow the rules that browsers expect.

One of those rules is that each file should contain one, and only one, HTML document. That's all of the content between the `<HTML>` tag and the `</HTML>` tag (tag names aren't case-sensitive, and it's common to see them in all-lowercase as in the example above). I mention this because one of the most common things I'll see folks do in PowerShell looks something like this:

```
Get-WmiObject -class Win32_OperatingSystem | ConvertTo-HTML | Out-File report.html
Get-WmiObject -class Win32_BIOS | ConvertTo-HTML | Out-File report.html -append
Get-WmiObject -class Win32_Service | ConvertTo-HTML | Out-File report.html -append 
```

"Aaarrrggh," says my colon every time I see that. You're basically telling PowerShell to create three complete HTML documents and jam them into a single file. While some browsers (Internet Explorer, notable) will figure that out and display something, it's just wrong. Once you start getting fancy with reports, you'll figure out pretty quickly that this approach is painful. It isn't PowerShell's fault; you're just not following the rules. Hence this guide!

You'll notice that the HTML consists of a lot of other tags, too: `<TABLE>, <TD>, <HEAD>`, and so on. Most of these are _paired_, meaning they come in an opening tag like `<TD>` and a closing tag like `</TD>`. The `<TD>` tag represents a table cell, and everything between those tags is considered the contents of that cell.

The `<HEAD>` section is important. What's inside there isn't normally visible in the browser; instead, the browser focuses on what's in the `<BODY>` section. The `<HEAD>` section provides additional meta-data, like what the title of the page will be (as displayed in the browser's window title bar or tab, not in the page itself), any style sheets or scripts that are attached to the page, and so on. We're going to do some pretty awesome stuff with the `<HEAD>` section, trust me.

You'll also notice that this HTML is pretty "clean," as opposed to, say, the HTML output by Microsoft Word. This HTML doesn't have a lot of visual information embedded in it, like colors or fonts. That's good, because it follows correct HTML practices of separating formatting information from the document structure. It's disappointing at first, because your HTML pages look really, really boring. But we're going to fix that, also.

In order to help the narrative in this book stay focused, I'm going to start with a single example. In that example, we're going to retrieve multiple bits of information about a remote computer, and format it all into a pretty, dynamic HTML report. Hopefully, you'll be able to focus on the techniques I'm showing you, and adapt those to your own specific needs.

In my example, I want the report to have five sections, each with the following information:

- Computer Information

- The computer's operating system version, build number, and service pack version.

- Hardware info: the amount of installed RAM and number of processes, along with the manufacturer and model. 

- An list of all processes running on the machine.

- A list of all services which are set to start automatically, but which aren't running.

- Information about all physical network adapters in the computer. Not IP addresses, necessarily - hardware information like MAC address.

I realize this isn't a universally-interesting set of information, but these sections will allow be to demonstrate some specific techniques. Again, I'm hoping that you can adapt these to your precise needs.
