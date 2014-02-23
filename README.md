WordProgressBar
===============

Progress Bar for Microsoft Word using vba to display progress against a target word count.

Installation
============

Install by copying code into your macro edit window.

> Optionally [create buttons on your ribbon bar or quick access toolbar](http://office.microsoft.com/en-us/word-help/create-or-run-a-macro-HA102919734.aspx "Office 2013 Instructions") for _ProgressCheck_ and _StopProgressCheck_ 

Run _ProgressCheck_ to kick off the process. Provide your target word count.  The status bar will update with 
your progress percentage and a _fancy_ progress bar made with the character "__I__" and spaces.

Keep typing and the status bar will return to normal.  Every 20 seconds the progress check will run again automatically,
but you might not see it if other things write to the status bar.  Pause for a few seconds and it should become visible again.

Run _StopProgressCheck_ to stop the automatic processing every 20 seconds.

Running _NumberOfWords_ by itself is not going to get you anywhere.

Current Known Issues
====================

1. _NumberOfWords_ should really be a private function
2. entering something other than a bare number (i.e. 10000 or 100 or 3758) will likely cause problems.



