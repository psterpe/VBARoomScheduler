# VBARoomScheduler
Simple (not real) conference room reservation backend for Excel VBA frontend.

## What This Is
These are the bones of a Python/Flask application that manages a simple database of conference
rooms. The app responds to a few routes, e.g., to list the rooms, and to reserve a room for a
certain block of hours. The database is Google Cloud NDB (which is Firebase in "datastore mode").

This is **not** a real conference room reservation system. It is a simple backend for an
also-simple frontend written in Excel VBA. That VBA code is contained in the
VBA_RoomRes.xlsm Excel workbook.

I used to teach a programming course that used Visual Basic. Occasionally, I would introduce
students to Excel VBA just for some exposure to it. I didn't want to use VBA simply to fiddle
with the contents of a worksheet, so I got the (arguably insane) idea to use an Excel worksheet
as the frontend to a mock "conference room reservation" system.

Here's what the spreadsheet frontend looks like after "PJS" has reserved the "Earth" conference
room from 10:00 AM to Noon:

![Image of spreadsheet frontend](https://github.com/psterpe/VBARoomScheduler/blob/master/RoomRes1.png)

Of course this is simplistic -- there are no dates! I just wanted students, working in small teams,
to add enough VBA code to the worksheet so that the _Refresh Rooms_ and _Reserve_ buttons would work.
Once they got things working, the idea was that teams would race to see who could make the most
reservations.

_Refresh Rooms_ contacts the backend for a list of the conference rooms, including their capacities
and information on when each room has been reserved (and by whom). _Reserve_ lets you reserve a
room. To use it, you select some adjacent cells in the spreadsheet (for example, to reserve Mars
for an hour starting at 1:00pm, you'd select cells O10 and P10 because Mars is on row 10, and columns O
and P represent 1:00 PM and 1:30 PM), put your initials in cell B15 (labeled "Team" because students
worked in teams), and click _Reserve_.

## Backend Data Model

The backend uses just one NDB entity named _Room_. It has these four properties:
* name (String)
* capacity (Integer)
* schedule (String)
* takers (String)

See _Format of the response to /list below_ for details on the format of the strings for _schedule_ and _takers_.

## What the Backend Can Do

Not much! Here are the routes it responds to:

| Route   | HTTP Method | Action |
|---------|:-----------:|--------|
| /list   | GET         |Responds with data about all rooms (details below) |
| /update | POST        |Reserves a room|
| /save   | POST        |Adds a conference room to the database (for initial setup) |
| /purge  | GET         |Deletes all conference rooms and associated data |

### Format of the response to /list

The backend responds to /list with a semicolon-delimited string; each segment contains data for
one conference room. The format of a segment is `key|name|capacity|schedule|takers` where:
* **key** is NDB's internal key for this entity
* **name** is the name of the conference room
* **capacity** is how many people the room holds (not used anywhere)
* **schedule** is a string of 16 0's or 1's, each representing a 30-minute time slot beginning
with 9:00 AM. This covers the 8 hours from 9:00 AM to 5:00 PM. A 0 means that the room is not
reserved in that time slot; a 1 means that it is reserved.
* **takers** is a colon-delimited string containing 16 segments, each of which represents a
30-minute time slot corresponding to the same time slots represented by **schedule**. A segment is
empty if nobody has the room reserved in that time slot, or it contains the initials of whoever
has reserved the room (the initials are whatever the user entered in cell B15 of the frontend
worksheet).
  
So why use all these delimited strings that contain delimited strings when we could just use
something like JSON? At the time, the programming course did not cover Internet APIs or JSON.
(It does now, and now it teaches Python, not Visual Basic.) My students could easily handle
substring operations, so I went with a sub-optimal data representation to keep the code simple.
They didn't have to learn JSON or any JSON parsing methods.

The code that sends data to the backend _does_ send it in JSON format because I wrote that part
of the code. It's only a little data, so I construct the JSON by hand, again because that uses
simple string operations my students could understand if they decided to read my code. (Most
did not.)

### Format of the data when POSTing to /update

To reserve a room, POST to the /update route with JSON like this:

```
{"roomkey":key,"schedule":"0011110000000000","taker":initials}
```

## Setting Up

If you deploy this code, you'll need to initialize your database of conference rooms. The bash
script _createdata.sh_ does this. It takes one argument -- the URL to your backend.

## Running Locally

Spreadsheet cell B17 contains the backend URL. During development, if you run the backend locally,
you can change this cell to http://127.0.0.1 or whatever address your local server is
listening on. You will still be modifying your production datastore unless you set up Google's
datastore emulator (see [datastore emulator docs](https://cloud.google.com/datastore/docs/tools/datastore-emulator)).