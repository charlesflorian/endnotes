# Endnote fixer

This is a very dirty and simple way to do some manipulation of endnotes in .docx files. This uses some of the work
done on https://github.com/python-openxml/python-docx.git. However, it uses an old pull request from 2014 which was
never merged, so that submodule commit has been added here.

Furthermore, it python-docx (at least in 2014) does not handle hyperlinks properly. This has been addressed in a very
dirty way: A run in a .docx file is a stretch of text, and looks like:
```
<w:r>
    <w:t>hello</w:t>
</w:r>
```
while a hyperlink is the same (or more than one run), but wrapped in a `<w:hyperlink>` tag. The solution are the
scripts `clean_endnotes.sh` and `remove_links.py`, which collectively remove all of the hyperlink tags.

## Usage

First (and optionally), one can remove hyperlinks from an existing documents with
```
$ bash clean_endnotes.sh file.docx
```
which will produce a file called `file-no-links.docx`.

Finally, one should run
```
$ python3 endnotes.py file-no-links.docx
```
The default behavior of `endnotes.py` is to take the (at most) first 10 words leading up to the reference in the main text,
and to add those to the start of the endnote reference.

## Options

`endnotes.py` has a few options. These are:

`-n num` - modify the number of words to keep leading up to the reference

`-o file` - choose an output file. The default is `file-endnotes.docx` if the input is `file.docx`.

`--chapter num` - Specialized. This will select only the references that are in Chapter `num`. Note that this is very article-specific, and probably won't work for yours.

`--sort` - Sorts the references. Default behavior is that the references occur in the resulting document in the order that they
occurred in the original document; this will sort them alphabetically instead.