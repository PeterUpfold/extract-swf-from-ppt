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

print("Converting to PPTX...")

for src_filename in tqdm(os.listdir(args.input_dir)):
    if src_filename.endswith(".ppt"):
        print(src_filename)
        print(['unoconv', '-f', 'pptx', '-o', os.path.abspath(args.output_dir), src_filename])
        subprocess.run(['unoconv', '-f', 'pptx', '-o', os.path.abspath(args.output_dir), src_filename], cwd=args.input_dir)

print("Extracting PPTX to ZIP...")

# loop over output folder and investigate zips for bin files
for pptx_file in tqdm(os.listdir(args.output_dir)):
    print(pptx_file)
    if pptx_file.endswith(".pptx"):
        zip_pptx = zipfile.ZipFile(os.path.join(args.output_dir, pptx_file), 'r')

        for entry_info in zip_pptx.infolist():
            #print(entry_info.filename)
            if entry_info.filename.endswith('.bin'):
                zip_pptx.extract(entry_info, path=args.output_dir)
                anim_dest=os.path.join(args.output_dir, pptx_file + ".Animations", entry_info.filename)
                os.renames(old=os.path.join(args.output_dir, entry_info.filename), new=anim_dest)
                print("Extracted bin animation ", anim_dest)

                subprocess.run(['csplitb', '--prefix', pptx_file, '--suffix', '.swf', '--number', '2', '465753', anim_dest])

                for anim_dest_file in os.listdir(os.getcwd()):
                    if anim_dest_file.endswith(".swf"):
                        print("Move ", anim_dest_file, "into place")
                        os.rename(src=os.path.join(os.getcwd(), anim_dest_file), dst=os.path.join(args.output_dir, pptx_file + ".Animations", anim_dest_file))