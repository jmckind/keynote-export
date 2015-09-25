#!/usr/bin/env python

from AppKit import NSURL, NSMutableDictionary
from ScriptingBridge import SBApplication
import sys

DEBUG = False
BUNDLE = 'com.apple.iWork.Keynote'

SAVING_OPTIONS = {
    'yes': 0x79657320,  # 'yes '
    'no':  0x6E6F2020,  # 'no  '
    'ask': 0x61736B20,  # 'ask '
}

EXPORT_FORMAT = {
    'Khtm': 0x4B68746D, # HTML
    'Kmov': 0x4B6D6F76, # QuickTime movie
    'Kpdf': 0x4B706466, # PDF
    'Kimg': 0x4B696D67, # Image
    'Kppt': 0x4B707074, # Microsoft PowerPoint
    'Kkey': 0x4B6B6579, # Keynote 09
}

IMAGE_FORMATS = {
    'Kifj': 0x4B69666A, # JPEG
    'Kifp': 0x4B696670, # PNG
    'Kift': 0x4B696674, # TIFF
}

MOVIE_FORMATS = {
    'Kmf3': 0x4B6D6633, # 360p
    'Kmf5': 0x4B6D6635, # 540p
    'Kmf7': 0x4B6D6637, # 720p
}

PRINT_WHAT = {
    'Kpwi': 0x4B707769, # individual slides
    'Kpwn': 0x4B70776E, # slides with notes
    'Kpwh': 0x4B707768, # handouts
}

EXPORT_PROPERTIES = {
    'Kxic': 0x4B786963, # compressed image quality, ranging from 0.0 to 1.0
    'Kxif': 0x4B786966, # format for resulting images
    'Kxmf': 0x4B786D66, # format for exported movie
    'Kxpw': 0x4B787077, # choose whether to include notes, etc.
    'Kxpa': 0x4B787061, # print each stage of builds
    'Kxps': 0x4B787073, # include skipped slides
    'Kxpb': 0x4B787062, # add borders around slides
    'Kxpn': 0x4B78706E, # include slide numbers
    'Kxpd': 0x4B787064, # include date
    'Kxkf': 0x4B786B66, # export in raw KPF
    'KxPW': 0x4B785057, # password
    'KxPH': 0x4B785048, # password hint
}

if len(sys.argv) < 2:
    print "usage: %s <keynote-file>" % sys.argv[0]
    sys.exit(-1)

# Export options
keynote_file = sys.argv[1]
to_file = NSURL.fileURLWithPath_(keynote_file.split('.key')[0] + '.pdf')
as_format = EXPORT_FORMAT['Kpdf']
with_properties = NSMutableDictionary.dictionaryWithDictionary_({
})

if DEBUG:
    print("   KEYNOTE_FILE: %s" % keynote_file)
    print("        TO_FILE: %s" % to_file)
    print("      AS_FORMAT: %s" % as_format)
    print("WITH_PROPERTIES: %s" % with_properties)

# Open Keynote file
keynote = SBApplication.applicationWithBundleIdentifier_(BUNDLE)
doc = keynote.open_(keynote_file)

# Export to format
doc.exportTo_as_withProperties_(to_file, as_format, with_properties)

# Close keynote
doc.closeSaving_savingIn_(SAVING_OPTIONS['no'], None)
keynote.quitSaving_(SAVING_OPTIONS['no'])
