#!/usr/bin/env python3
#
# Extract SWF from PPT
#
# Use LibreOffice's unoconv to convert PowerPoint 97-2003 format files into PPTX, extract Flash files (*.swf)
# from inside the presentation and dump these to a folder.

# Copyright 2019 Test Valley School.
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

import argparse
import os
import subprocess
from tqdm import tqdm
import zipfile
#require csplitb from Pip

argparser = argparse.ArgumentParser(description='Convert PowerPoint 97-2003 fiels and extract Flash files (*.swf) from inside them.')
argparser.add_argument('-i', '--input-dir', dest='input_dir', help='The directory containing source files', required=True)
argparser.add_argument('-o', '--output-dir', dest='output_dir', help='The directory where output files should be written.', required=True)

args = argparser.parse_args()

if not os.path.isdir(args.input_dir):
    raise ValueError("The input directory specified is not a directory.")

# check output path
if not os.path.exists(args.output_dir):
    raise ValueError("The output directory specified does not exist.")

if not os.path.isdir(args.output_dir):
    raise ValueError("The output directory specified is not a directory.")

#print("Converting to PPTX...")

#for src_filename in tqdm(os.listdir(args.input_dir)):
    #if src_filename.endswith(".ppt"):
        #print(src_filename)
        #print(['unoconv', '-f', 'pptx', '-o', os.path.abspath(args.output_dir), src_filename])
        #subprocess.run(['unoconv', '-f', 'pptx', '-o', os.path.abspath(args.output_dir), src_filename], cwd=args.input_dir)

print("Extracting PPTX to ZIP...")

# loop over output folder and investigate zips for bin files
for pptx_file in tqdm([os.path.join(dp, f) for dp, dn, fn in os.walk(args.output_dir) for f in fn]):
    print(pptx_file)
    if pptx_file.endswith(".pptx"):
        zip_pptx = zipfile.ZipFile(pptx_file, 'r')

        anim_count = 0
        for entry_info in zip_pptx.infolist():
            #print(entry_info.filename)
            if entry_info.filename.endswith('.bin'):
                zip_pptx.extract(entry_info, path=args.output_dir)
                anim_dest=os.path.join(pptx_file + ".Animations", entry_info.filename)
                print(os.path.join(args.output_dir, entry_info.filename))
                os.renames(old=os.path.join(args.output_dir, entry_info.filename), new=anim_dest)
                print("Extracted bin animation ", anim_dest)

                subprocess.run(['csplitb', '--prefix', pptx_file + str(anim_count), '--suffix', '.swf', '--number', '2', '465753', anim_dest], cwd=os.getcwd())

                print("Ran csplitb for ", anim_dest)

                anim_count += 1

                for anim_dest_file in os.listdir(os.path.dirname(pptx_file)):
                    if anim_dest_file.endswith(".swf"):
                        print("Move ", anim_dest_file, "into place")
                        os.renames(old=os.path.join(os.path.dirname(pptx_file), anim_dest_file), new=os.path.join(pptx_file + ".Animations", anim_dest_file))