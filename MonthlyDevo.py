import sys, os, os.path, re, zipfile, shutil, stat, time, subprocess, glob
from datetime import date, timedelta

# The next two datastructures, days and daysTimes must be sorted by weekdays
# which days are included, and the numerical equivalent for the datetime module
days = (("Tuesday", 2), ("Wednesday", 3), ("Thursday", 4))

# required days and times, organized in a hierarchy of weekdays corresponding to days above
daysTimes = (
    ("10:00 AM", "10:30 AM", "4:00 PM"),  # Tues
    ("10:00 AM", "10:30 AM"),  # Weds
    ("10:00 AM", "10:30 AM", "1:00 PM", "3:00 PM"),
)  # Thurs

inputFiles = [f"Devotion-{n}.odt" for n in range(1, 6)]
subdir = "unpack"

# regular expression library
reDay = re.compile(
    r'<text:user-field-decl office:value-type="string" office:string-value="[^"]+" text:name="Day"/>'
)
reDate = re.compile(
    r'<text:user-field-decl office:value-type="string" office:string-value="[^"]+" text:name="DateRange"/>'
)
reTime = re.compile(
    r'<text:user-field-decl office:value-type="string" office:string-value="[^"]+" text:name="Time"/>'
)
reTitle = re.compile(r"<dc:title>[^<]+</dc:title>")

# build up calendar
firstDay = days[0][1]  # numerical first weekday of our schedule
nextMonth = date.today() + timedelta(
    days=27
)  # pick a day in the next month, to get the month/year
# correct
firstNextMonth = date(
    nextMonth.year, nextMonth.month, 1
)  # find the first of next month
workingMonth = firstNextMonth.month
monthName = firstNextMonth.strftime("%B")
# how many days from the first of the next month to our first active weekday
n = firstNextMonth.isoweekday()
activeOffset = (firstDay - n) if (n < firstDay) else (firstDay + 7 - n)
# this value is initialized, but will be incremented as we do work:
workingDate = firstNextMonth + timedelta(days=activeOffset)
# now come up with the delta times to cycle through dates - so days to go from first weekday
# to second, second to third, ... and last to the first weekday of the next week
cycle = [days[i + 1][1] - days[i][1] for i in range(len(days) - 1)]
cycle.append(7 + days[0][1] - days[-1][1])


# working directory
os.chdir("c:/Users/buchs/odp/Documents/FHP-Ministries/Materials")
# output directory same as month name
if not os.path.exists(monthName):
    os.mkdir(monthName)

# loop over the weeks, stop when we hit the first weekday in the next month
weekIndex = 0
while workingDate.month == workingMonth:

    print("starting with ", inputFiles[weekIndex])
    # unpack our input file.
    zf = zipfile.ZipFile(inputFiles[weekIndex], "r")
    zf.extractall(path=subdir)
    zf.close()

    # work in unpacked dir
    os.chdir(subdir)

    # grab content to be ready to edit content
    fp = open("content.xml")
    content = fp.read()
    fp.close()

    # grab meta data to be ready to edit it
    fp = open("meta.xml")
    meta = fp.read()
    fp.close()

    # loop over days of week
    for dayIndex in range(len(cycle)):

        print("date is ", workingDate.isoformat())
        # Update the day of the week and date in the content
        daySub = f'<text:user-field-decl office:value-type="string" office:string-value="{days[dayIndex][0]}" text:name="Day"/>'
        content = reDay.sub(daySub, content)
        dateString = workingDate.strftime("%B %d, %Y").replace(" 0", " ")
        dateSub = f'<text:user-field-decl office:value-type="string" office:string-value="{dateString}" text:name="DateRange"/>'
        content = reDate.sub(dateSub, content)

        for timeIndex in range(len(daysTimes[dayIndex])):

            thisDayTime = daysTimes[dayIndex][timeIndex]
            # make a simple form of time for naming the files
            timeSimple = (
                "-"
                + thisDayTime.replace(":", "").replace(" AM", "").replace(" PM", "")
                + "-"
            )

            timeSub = f'<text:user-field-decl office:value-type="string" office:string-value="{thisDayTime}" text:name="Time"/>'
            content = reTime.sub(timeSub, content)

            # overwrite the content file
            fp = open("content.xml", "w")
            fp.write(content)
            fp.close()

            # overwrite the metadata file with document title
            dateStmp = workingDate.strftime("%b-%d-%Y")
            titleSub = f"<dc:title>Devotion {days[dayIndex][0][0:3]} {thisDayTime} {dateStmp}</dc:title>"
            meta = reTitle.sub(titleSub, meta)

            # overwrite the meta file
            fp = open("meta.xml", "w")
            fp.write(meta)
            fp.close()

            # Create new output file and open as zipfile
            outputFile = (
                "../"
                + monthName
                + "/"
                + days[dayIndex][0][0:3]
                + timeSimple
                + dateStmp
                + "-"
                + inputFiles[weekIndex]
            )
            # like: Devotion-1-Tue-1000-Apr-01-2018.odt
            zf = zipfile.ZipFile(outputFile, "w")

            # write files, subdirs and files in subdirs to this zip file
            for f in os.listdir("."):
                zf.write(f)
                mode = os.stat(f).st_mode
                if stat.S_ISDIR(mode):
                    for g in os.listdir(f):
                        zf.write(f + "/" + g)

            zf.close()
            print(f"Wrote {outputFile}")

        # We reach the end of times for a given day, now advance the date.
        # This will allow tracking when we bump into next month.
        # This will automatically take care of the week jumps too.
        workingDate += timedelta(days=cycle[dayIndex])

    # and we are on to the next week
    weekIndex += 1
    os.chdir("..")
    shutil.rmtree(subdir)  # clean up unpacked files to prepare for next


# Now, convert the ODT files to PDF files
os.chdir(monthName)
# start the converter server
subprocess.run("python c:/python36/Scripts/unoconv --listener &", shell=True)
time.sleep(20)
# make one pass through everything
for fn in glob.glob("*.odt"):
    subprocess.run(f"python c:/python36/Scripts/unoconv -f pdf {fn}", shell=True)

# now cycle through looking for missing pdf files, because unoconv can fail.
missing = 1
while missing > 0:
    missing = 0
    for fn in glob.glob("*.odt"):
        pdfn = fn.replace(".odt", ".pdf")
        if not os.path.exists(pdfn):
            subprocess.run(
                f"python c:/python36/Scripts/unoconv -f pdf {fn}", shell=True
            )
            missing += 1
    print("missing ", missing)
