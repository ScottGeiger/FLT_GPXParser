# FLT_GPXParser
Parse FLT GPX Files generating an output file for an E2E spreadsheet.  This is a work in progress and as of June 2022 is currently in beta v1.

## Description:
The script first catalogs all waypoints into a collection.  Next it will present a list of available tracks (unless the -t command line parameter has been used) to process; only one may be selected.  The script will "walk" the selected track and catalog all of the track points into a collection.  It will then iterate through the collection of waypoints and track points to determine the closest track point to the waypoint.  Finally it will build "segments" by "walking" the track collection creating a running distance total until it reaches a nearby waypoint.  When a waypoint is near the segment is ended and a new segment begins.  The code also looks back three track points and ahead three track points to calculate a change in bearing and determine a left/right turn and the degree to which there is a turn.

## Required Libraries:
The following Python libraries are necessary for gpxparse to function.  Use `pip install -r requirements.txt` to install.
  * coloredlogs (v15.0.1) - see: https://github.com/xolox/python-coloredlogs
  * geopy (v2.2.0) - see: https://github.com/geopy/geopy
  * gpxpy (v1.5.0) - see: https://github.com/tkrajina/gpxpy
  * openpyxl (v3.0.9) - see: https://openpyxl.readthedocs.io/en/stable/

## Usage:
```

FLT Map GPX Segment Processor
-----------------------------
  You must specify a GPX file to process and optionally a track.  If you do not specify the track you will be prompted.
  The output will be Excel, however you can specify CSV via the -c (--csv) optional command parameter.

positional arguments:
  gpxfile               Specify the gpx file to process

optional arguments:
  -h, --help            show this help message and exit
  -c, --csv             Output as CSV instead of Excel
  -r, --reverse         Reverse the direction of the track
  -t TRACK, --track TRACK
                        Specify the track name to process
  -v, --verbose         Verbose mode. Multiple -v options increase the verbosity. The maximum is 3
  -i INCLUDE, --include INCLUDE
                        Symbol to include. Use multiple -i options for additional symbols. Cannot be used with
                        -x/--exclude
  -x EXCLUDE, --exclude EXCLUDE
                        Symbol to exclude. Use multiple -x options for additional symbols. Cannot be used with
                        -i/--include

```

## TODO:
  * Add ability to process folder/multiple GPX files
  * Add ability to "merge" tracks
  * Add detection of secondary track crossing/junction

