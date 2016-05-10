# Combining HTML Reports and a GUI Application

I've had a number of folks ask questions in the forums at PowerShell.org, with the theme of "how can I use a RichTextBox in a Windows GUI application to display nicely formatted data?" My answer is don't. Use HTML instead. For example, let's say you followed the examples in the previous chapter and produced a beautiful HTML report. Keep in mind that the report stays "in memory," not in a text file, until the very end:

```
      $params = @{'CssStyleSheet'=$style;
                    'Title'="System Report for $computer";
                    'PreContent'="<h1>System Report for $computer</h1>";
                    'CssIdsToMakeDataTables'=@('tableProc','tableNIC','tableSvc');
                    'HTMLFragments'=@($html_os,$html_cs,$html_pr,$html_sv,$html_na)}
        ConvertTo-EnhancedHTML @params |
        Out-File -FilePath $filepath
```

For the sake of illustration, let's say that's now in a file named C:\Report.html. I'm going to use SAPIEN's PowerShell Studio 2012 to display that report in a GUI, rather than popping it up in a Web browser. Here, I've started a simple, single-form project. I've changed the text of the form to "Report," and I've added a WebBrowser control from the toolbox. That control automatically fills the entire form, which is perfect. I named the WebBrowser control "web," which makes it accessible from code via the variable $web.

I'll note that PowerShell Studio 2012 is very out-of-date at this point, but you should still get the general idea.

![image006.png](/manuscript/image006.png)

I expect you'd make a form like this part of a larger overall project, but I'm just focusing on how to do this one bit. So I'll have the report load into the WebBrowser control when this form loads:

```
$OnLoadFormEvent={
#TODO: Initialize Form Controls here
    $web.Navigate('file://C:\report.html')
} 
```

Now I can run the project:

![image007.png](/manuscript/image007.png)

I get a nice pop-up dialog that displays the HTML report. I can resize it, minimize it, maximize it, and close it using the standard buttons on the window's title bar. Easy, and it only took 5 minutes to create.
