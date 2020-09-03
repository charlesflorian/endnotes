import sys
import os
import re

"""
This is a quick and dirty script to remove the <w:hyperlink> tags from an .xml file. This is not the best
long-term solution, but it will work for the time being.
"""

if len(sys.argv) < 1:
    print("You must pass a file name")
    sys.exit(1)

if not os.path.isfile(sys.argv[1]):
    print(f"File '{sys.argv[1]}' does not exist")
    sys.exit(1)

filename = sys.argv[1]

with open(filename, "r") as f:
    out = re.sub(r"</?w:hyperlink[^>]*>", "", f.read())


with open(filename, "w") as f:
    f.write(out)
