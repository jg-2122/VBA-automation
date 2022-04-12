# VBA-automation

If your organization is using an Excel-based control tracker, list, or matrix and 
manually updates leadsheets by copying the information over into each workpaper,
here's a minimum-viable macro solution built using VBA code

I have not tested it in every setting and this has yet to go through a proper SDLC,
but I have provided the logic in the document containing the code, which is widely applicable
and can be expanded upon

I have highlighted the lines of code in the txt file that need to be reconfigured depending
on how you want to implement this. The ranges are hardcoded at the moment but those
can be replaced with FileDialog and InputBox options that provide a safe interface for a user to
interact with instead of fudging with the code

Could also be built in Python (easier syntax and expanded functionality), Alterx (low-code), etc

The code requires more testing and definitely could use improvements - I was just playing around
with VBA for fun and thought this would be a good use case for writing a VBA macro

