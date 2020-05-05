#!/usr/bin/env python3
# Author: Leon Kuchenbecker <leon.kuchenbecker@uni-tuebingen.de>

import argparse
import pandas as pd
import sys
import tempfile
import subprocess
import os
import shutil

from collections import defaultdict

####################################################################################################

parser = argparse.ArgumentParser()
parser.add_argument('infile', type = str, metavar = '<sheetpath>', help = 'Path to the Excel file')
parser.add_argument('textemplate', type = str, metavar = '<texpath>', help = 'Path to the LaTeX template')
parser.add_argument('tutname', type = str, metavar = '<tutorname>', help = 'Name of the correcting tutor')
parser.add_argument('sheet', type = int, help = 'Sheet number to process')

####################################################################################################

def openLaTeX():
    """ Create a temporary directory with a TeX file. Writes header and settings to the file. Returns file handle. """
    tdir = tempfile.mkdtemp()
    out = open(tdir + '/out.tex', 'w')
    return tdir, out

####################################################################################################

def compileLaTeX(tdir):
    """ Compile the 'hmm.tex' file in the given directory """
    ret = subprocess.run(['pdflatex', '--interaction', 'batchmode', 'out'], cwd = tdir, stdout = subprocess.PIPE, stderr = subprocess.PIPE)
    if ret.returncode != 0:
        raise RuntimeError('An error occurred during LaTeX execution')
    ret = subprocess.run(['pdflatex', '--interaction', 'batchmode', 'out'], cwd = tdir, stdout = subprocess.PIPE, stderr = subprocess.PIPE)
    if ret.returncode != 0:
        raise RuntimeError('An error occurred during LaTeX execution')

####################################################################################################

def process(args):
    # Read Scores and comments
    scores = pd.read_excel(args.infile, sheet_name = 'Grading')

    # Read participants
    participants = pd.read_excel(args.infile, sheet_name = 'Participants').set_index('Username')

    # Read maximum scores
    sheet_meta       = pd.read_excel(args.infile, sheet_name = 'Sheets')
    max_scores       = { (row.Sheet, row.Task, row.Subtask) : row.MaxScore for _, row in sheet_meta.iterrows() }
    max_scores_sheet = sheet_meta.groupby('Sheet')['MaxScore'].sum() 

    # Select sheet
    scores = scores[scores.Sheet==args.sheet]

    if scores.empty:
        print("No matching grades found")
        return 1

    # Read LaTeX template
    with open(args.textemplate, 'r') as infile:
        template = infile.read()

    # Build dictionaries
    d = defaultdict(lambda : defaultdict( lambda : { 'score' : None, 'comments' : []} ) )

    for _,row in scores.iterrows():
        task = (row.Sheet, row.Task, row.Subtask)
        record = d[row.Username][task] 
        if row.Type.upper() == 'SCORE':
            if record['score'] is not None:
                raise RuntimeError(f'Duplicate score for identical task found (User: {row.Username}, Sheet: {row.Sheet}, Task: {row.Task}, Subtask: {row.Subtask}')
            try:
                record['max_score'] = max_scores[task]
            except KeyError:
                raise RuntimeError(f'Could not find maximum score for task {task}')
            record['score'] = row.Value
        elif row.Type.upper() == 'COMMENT':
            record['comments'].append(row.Value)
        else:
            raise RuntimeError(f'Invalid value type "{row.Type}".')


    for user, tasks in d.items():
        body=[]

        # Total score
        total_score = sum([ record['score'] for _, record in tasks.items() if record['score'] ])
        try:
            max_total_score = max_scores_sheet.loc[args.sheet]
        except KeyError:
            raise RuntimeError(f'Cannot calculate maximum total score for sheet {args.sheet}')

        # Participant name
        try:
            realname = participants.loc[user].Name
        except KeyError:
            raise RuntimeError(f'Cannot find real name for user "{user}".')


        tex=template.replace('§§sheetnr§§', str(args.sheet)).replace('§§fullname§§', str(realname)).replace('§§tutorname§§', args.tutname).replace('§§total§§', str(total_score)).replace('§§maxtotal§§', str(max_total_score))

        # Sort tasks and records
        sorted_tasks = sorted([ (task, record) for task, record in tasks.items() ], key = lambda x : x[0])
        cur_task = None
        for task, record in sorted_tasks:
            score_str = '{:.2f}'.format(record['score'])
            if task[1] != cur_task:
                cur_task = task[1]
                body.append(r'\newtask{' + str(cur_task) + r'}')
            body.append(f'\\scoreTask{{{task[0]}}}{{{task[1]}}}{{{task[2]}}}{{{score_str}}}{{{record["max_score"]}}}')
            if record['comments']:
                body.append(r'\beforeComments')
                for comment in record['comments']:
                    body.append(f'\\comment{{{comment}}}')
                body.append(r'\afterComments')

        # Write out and compile tex
        tex = tex.replace('§§body§§', '\n'.join(body))
        tdir, out = openLaTeX()
        out.write(tex)
        out.close()
        compileLaTeX(tdir)

        # Move output file in place
        outpath = f'sheet{args.sheet}.{user}.pdf'
        if os.path.exists(outpath):
            raise RuntimeError(f'Output path {outpath} exists! Aborting!')
        shutil.move(os.path.join(tdir, 'out.pdf'), outpath)
        print(f'[OK]\t{user}')

    return 0

####################################################################################################

def main(argv=sys.argv):
    args = parser.parse_args()
    try:
        sys.exit(main(args))
    except RuntimeError as e:
        print(f'\nERROR: {e}')
        sys.exit(1)
    except KeyboardInterrupt:
        print(f'\nAborting.')
        sys.exit(2)
