# VBARoomScheduler
Simple (not real) conference room reservation backend for Excel VBA frontend.

## What This Is
These are the bones of a Python/Flask application that manages a simple database of conference rooms. The
app responds to a few routes, e.g., to list the rooms, and to reserve a room for a certain block of hours.
The database is Google Cloud NDB (basically, Firebase in "datastore mode").

This is **not** a real conference room reservation system. It is a simple backend for an also-simple frontend
written in Excel VBA. That VBA code is contained in the VBA_RoomRes.xlsm Excel workbook.

I used to teach a programming course that used Visual Basic. Occasionally, I would introduce students to
Excel VBA just for some exposure to it. I didn't want to use VBA simply to fiddle with the contents of a
worksheet, so I got the (arguably insane) idea to use an Excel worksheet as the frontend to a "conference
room reservation" system.

# This file still in progress...
