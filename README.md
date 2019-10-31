# extract-swf-from-ppt

Use LibreOffice's unoconv to convert PowerPoint 97-2003 format files into PPTX, extract Flash files (*.swf)
from inside the presentation and dump these to a folder.

See also the [pptx-already-converted](https://github.com/TestValleySchool/extract-swf-from-ppt/tree/pptx-already-converted) branch for an approach that uses a VBA macro within PowerPoint to perform the first step of converting to *.pptx.

## Requirements

Install tqdm using Pip (or on CentOS 7, `yum install python36-tqdm`)
Install csplitb using Pip
