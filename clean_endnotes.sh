#!/bin/env bash

function usage {
    echo "Usage:"
    echo "$ bash clean_endnotes.sh file.docx"
} 

if [ $# -lt 1 ]; then
    usage
    exit 1
fi

file=$1
ext=${file##*.}
base=${file%.*}

if [ ! -f "$file" ]; then
    echo "File $file does not exist" >&2
    exit 1
fi

dir="$base"_unzipped

unzip "$file" -d "$dir"

if [ -f "$dir/word/endnotes.xml" ]; then
pushd $dir
    pushd word
        python3 ../../remove_links.py endnotes.xml
    popd
    zip -r "../$base-no-links.docx" *
popd
rm -rf "$dir"
else
    echo "There are no endnotes in $file."
fi