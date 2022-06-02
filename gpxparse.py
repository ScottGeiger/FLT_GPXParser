#!/usr/bin/env python3
from collections import deque
from openpyxl import Workbook
from openpyxl.styles import Font, DEFAULT_FONT
from os.path import splitext
import sys, csv, argparse, textwrap
import logging,coloredlogs
from re import sub
import math, statistics
import gpxpy
import gpxpy.gpx
from geopy.distance import geodesic
from geopy.geocoders import Nominatim

#TODO: make tracks and waypoints a class

parser = argparse.ArgumentParser(
    formatter_class=argparse.RawDescriptionHelpFormatter,
    description=textwrap.dedent('''\
        FLT Map GPX Segment Processor
        -----------------------------
          You must specify a GPX file to process and optionally a track.  If you do not specify the track you will be prompted.
          The output will be Excel, however you can specify CSV via the -c (--csv) optional command parameter.
    '''))
parser.add_argument('gpxfile',help='Specify the gpx file to process')
parser.add_argument('-c','--csv',action='store_true',help='Output as CSV instead of Excel')
parser.add_argument('-r','--reverse',action='store_true',help='Reverse the direction of the track')
parser.add_argument('-t','--track',help='Specify the track name to process')
parser.add_argument('-v','--verbose',action='count',default=0,help='Verbose mode.  Multiple -v options increase the verbosity.  The maximum is 3')
parser.add_argument('-x','--exclude',action='append',help='Symbol to exclude.  Use multiple -x options for additional symbols')
args = vars(parser.parse_args())

# 0: log level 3 (warning)
# 1: log level 2 (info)
# 2: log level 1 (debug)
LOG_LEVEL = max((3-args.get('verbose')),1)*10

## Helper Functions ##
def initLog():
    format_string = '%(asctime)s [%(levelname)8s]:  %(message)s'
    global logger
    logger = logging.getLogger(__name__)
    logger.setLevel(LOG_LEVEL)
    levelStyles = {
        'debug':{'color':'cyan'},
        'info':{'color':'white'},
        'warning':{'color':'yellow'},
        'error':{'color':'red'},
        'critical':{'bold':True,'color':'magenta'}
    }
    coloredlogs.install(logger=logger,level=LOG_LEVEL,fmt=format_string,level_styles=levelStyles)

def camel_case(s):
  s = sub(r"[^0-9a-zA-Z]+", " ", s).title().replace(" ", "")
  return ''.join([s[0].lower(), s[1:]])

def calc_bearing(pointA,pointB):
    deg2rad = math.pi/180
    latA = pointA['latitude'] * deg2rad
    latB = pointB['latitude'] * deg2rad
    lonA = pointA['longitude'] * deg2rad
    lonB = pointB['longitude'] * deg2rad
    delta_ratio = math.log(math.tan(latB/ 2 + math.pi / 4) / math.tan(latA/ 2 + math.pi / 4))
    delta_lon = abs(lonA - lonB)

    delta_lon %= math.pi
    bearing = math.atan2(delta_lon, delta_ratio)/deg2rad
    return bearing

## Main ##
initLog();
try:
    gpx_file = open(args.get('gpxfile'),'r')
except:
    logger.critical('Could not open GPX File: {}'.format(args.get('gpxfile')))
    sys.exit(1)
gpx = gpxpy.parse(gpx_file)

# Get track to process
trk = None
if args.get('track') is not None:
    for track in gpx.tracks:
        if track.name == args.get('track'): trk = track
    if trk is None:
        print('\nNo track with that name found, please make a selection.')

while trk is None:
    print('\nTrack List for {}:'.format(args.get('gpxfile')))
    for i,track in enumerate(gpx.tracks):
        print('  {0}.\t{1}'.format(int(i)+1,track.name))
    print('\n  X.\tExit')
    n = input('\nEnter the # of the track to process: ')
    if n == 'x' or n == 'X': sys.exit(0)
    try:
        trk = gpx.tracks[int(n)-1]
        print('')
    except:
        print('Invalid track selection.  Please select a valid track #')
        trk = None

logger.info('Beginning Processing of {0} for track {1}'.format(args.get('gpxfile'),trk.name))

# Collect Waypoints
logger.info('Processing waypoints...')
#waypoints = list()
waypoints = deque()
wpt_symbols = dict()
for waypoint in gpx.waypoints:
    if waypoint.symbol not in args.get('exclude'):
        waypoints.append({'name':waypoint.name,'description':waypoint.description,'symbol':waypoint.symbol,'latitude':waypoint.latitude,'longitude':waypoint.longitude,'nearest_trkpoint':None,'nearest_distance':math.inf})
        wpt_symbols[camel_case(waypoint.symbol)] = waypoint.symbol
#    geolocator = Nominatim(user_agent="Finger Lakes Trail Mapper")
#    location = geolocator.reverse('{0},{1}'.format(waypoint.latitude, waypoint.longitude))
#    print(location.raw)
logger.info('Including {0} Symbols: "{1}"'.format(len(wpt_symbols),'","'.join(sorted(wpt_symbols.values()))))
logger.info('Ignoring {0} Symbols: "{1}"'.format(len(args.get('exclude')),'","'.join(sorted(args.get('exclude')))))
logger.info('Processed {} waypoints'.format(len(waypoints)))

# Collect Track Point
logger.info('Processing trakpoints for {}...'.format(trk.name))
#trkpoints = list()
trkpoints = deque()
for track in gpx.tracks:
    if track.name == trk.name:
        for segment in track.segments:
            for point in segment.points:
                trkpoints.append({'latitude':point.latitude,'longitude':point.longitude,'near_waypoint':None,'distance_from_waypoint':None})
logger.info('Processed {} trackpoints'.format(len(trkpoints)))

# reverse the track direction if option indicated
if args.get('reverse'):
    logger.warning('Reversing track direction')
    trkpoints.reverse()

# Determine closest trackpoint to waypoint
for i,waypoint in enumerate(waypoints):
    for j,trkpoint in enumerate(trkpoints):
        dist = geodesic((trkpoint['latitude'],trkpoint['longitude']),(waypoint['latitude'],waypoint['longitude'])).feet
        if dist < waypoint['nearest_distance']:
            waypoints[i]['nearest_trkpoint'] = j
            waypoints[i]['nearest_distance'] = dist
for k,waypoint in enumerate(waypoints):
    if waypoint['nearest_trkpoint'] is not None:
        trkpoints[waypoint['nearest_trkpoint']]['near_waypoint'] = k
        trkpoints[waypoint['nearest_trkpoint']]['distance_from_waypoint'] = waypoint['nearest_distance']


#Build segments
logger.info('Building segments...')
logger.info('-'*60)
segments = deque()
track_distance = 0
segment_distance = 0
measure = deque()
for i,trkpoint in enumerate(trkpoints):
    measure.append(trkpoint)
    if len(measure) == 2:
        dist = geodesic((measure[0]['latitude'],measure[0]['longitude']),(measure[1]['latitude'],measure[1]['longitude'])).miles
        track_distance += dist
        segment_distance += dist
        #del measure[0]
        measure.popleft()
    wpt = trkpoint['near_waypoint']
    if wpt is not None:
        try:
            h0 = calc_bearing(trkpoints[i-3],trkpoints[i])
            h1 = calc_bearing(trkpoints[i],trkpoints[i+3])        
            dH = ((h1-h0)+360)%360
            #TODO: need to reverse directions with reverse
            direction = ''
            dirlist = ['continue straight','slight right','turn right','sharp right','u-turn to the right','turn around','u-turn to the left','sharp left','turn left','slight left','continue straight']
            if args.get('reverse'): dirlist.reverse()
            if dH >= 0 and dH <= 10: direction = dirlist[0]
            if dH >= 11 and dH <= 30: direction = dirlist[1] #slight (20)
            if dH >= 31 and dH <= 110: direction = dirlist[2] #turn (80)
            if dH >= 111 and dH <= 160: direction = dirlist[3] #sharp (50)
            if dH >= 161 and dH <= 170: direction = dirlist[4] #uturn (10)
            if dH >= 171 and dH <= 190: direction = dirlist[5]
            if dH >= 191 and dH <= 200: direction = dirlist[6] #uturn
            if dH >= 201 and dH <= 250: direction = dirlist[7] #sharp
            if dH >= 251 and dH <= 330: direction = dirlist[8] #turn
            if dH >= 331 and dH <= 350: direction = dirlist[9] #slight
            if dH >= 351 and dH <= 360: direction = dirlist[10]
            segments.append([waypoints[wpt]['name'],waypoints[wpt]['description'],direction,round(segment_distance,1),round(track_distance,1)])
            segment_distance = 0
        except:
            # if i < 3 then we don't need to measure the segment
            pass

for segment in segments:
    logger.info(', '.join(str(e) for e in segment))

logger.info('-'*60)
logger.info('Total Track Distance (in miles): {}'.format(round(track_distance,2)))

ext = 'csv' if args.get('csv') else 'xlsx'
if args.get('reverse'):
    filename = splitext(args.get('gpxfile'))[0] + '_reverse.'+ext
else: 
    filename = splitext(args.get('gpxfile'))[0] + '.'+ext

fldnames = ["Way Point Name","Description","Direction","Segment","Running Total"]
if args.get('csv'):
    logger.info('Writing data to CSV file: {}...'.format(filename))
    csvfile = open(filename,'w')
    writer = csv.writer(csvfile,delimiter=',',quoting=csv.QUOTE_ALL)
    writer.writerow(fldnames)
    for segment in segments:
        writer.writerow(segment)
    csvfile.close()
    logger.info('Finished writing to CSV file')
else: # Excel Output
    DEFAULT_FONT.name = 'Arial'
    logger.info('Writing data to Excel file: {}...'.format(filename))
    wb = Workbook()
    ws = wb.active
    ws.title = 'FLT Tracking'
    ws.append(fldnames)
    ws.row_dimensions[1].font = Font(name='Arial',bold=True)
    row = 2
    for i,segment in enumerate(segments):
        segment.pop()
        if i == 0:
            segment.append('=INDIRECT(ADDRESS(ROW(),COLUMN()-1))')
        else:
            segment.append('=INDIRECT(ADDRESS(ROW(),COLUMN()-1))+INDIRECT(ADDRESS(ROW()-1,COLUMN()))')
        ws.append(segment)
        row += 1
    ws.column_dimensions['A'].width = 20
    ws.column_dimensions['B'].width = 40
    ws.column_dimensions['C'].width = 20
    ws.column_dimensions['D'].number_format = '##0.0'
    ws.column_dimensions['E'].number_format = '##0.0'
    #map_range = openpyxl.workbook.defined_name.DefinedName('Map_M18', attr_text='A1:E{}'.format(row-1))
    #wb.defined_names.append(map_range)
    total_row = ['Total:']
    total_row += ['' for i in range(len(fldnames)-2)]
    total_row.append('=SUM(INDEX(A:E,2,COLUMN()-1):INDEX(A:F,ROW()-1,COLUMN()-1))')
    ws.append(total_row)
    ws.merge_cells(start_row=row,start_column=1,end_row=row,end_column=len(fldnames)-1)
    ws.row_dimensions[row].font = Font(name='Arial',bold=True)
    ws.freeze_panes = ws['A2']
    wb.save(filename)
    logger.info('Finished writing to Excel file')