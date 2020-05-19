# Copyright (c) 2020 Leon Kuchenbecker <leon.kuchenbecker@uni-tuebingen.de>
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.

import argparse
import pandas as pd
import sys
import tempfile
import subprocess
import os
import shutil
import getpass

from collections import defaultdict
from assignmenttool.errors import AToolError

from assignmenttool import config, SMTPClient

####################################################################################################

def compileLaTeX(tex, pdflatex):
    """ Compile the 'hmm.tex' file in the given directory """
    tdir = tempfile.mkdtemp()
    out = open(tdir + '/out.tex', 'w')
    out.write(tex)
    out.close()
    ret = subprocess.run([pdflatex, '--interaction', 'batchmode', 'out'], cwd = tdir, stdout = subprocess.PIPE, stderr = subprocess.PIPE)
    if ret.returncode != 0:
        raise AToolError('An error occurred during LaTeX execution')
    ret = subprocess.run([pdflatex, '--interaction', 'batchmode', 'out'], cwd = tdir, stdout = subprocess.PIPE, stderr = subprocess.PIPE)
    if ret.returncode != 0:
        raise AToolError('An error occurred during LaTeX execution')
    with open(os.path.join(tdir, 'out.pdf'), 'rb') as infile:
        pdf = infile.read()
    shutil.rmtree(tdir)
    return pdf

####################################################################################################

def mail_feedback(config, participants, pdfs):
    """Send the feedback PDFs to the participants"""
    if config.mail_smtp_user and not config.mail_smtp_pass:
        config.smtp_password = getpass.getpass(f'Password for [{config.mail_smtp_user}@{config.mail_smtp_host}]: ')

    smtp = SMTPClient.SMTPClient(
            hostname = config.mail_smtp_host,
            port = config.mail_smtp_port,
            user = config.mail_smtp_user,
            password = config.smtp_password,
            tls = not config.mail_smtp_no_tls)

    for user, pdf in pdfs.items():
        # Lookup recipient name and email address
        try:
            name = participants.loc[user]['Name']
            email = participants.loc[user]['E-Mail']
        except KeyError:
            raise AToolError(f'Failed to look up name and email address for user "{user}".')
        smtp.sendMessage(
                sender       = (config.mail_sender_name, config.mail_sender_address),
                recipients   = (name, email),
                subject      = config.mail_subject.replace('§§username§§', user).replace('§§name§§', name).replace('§§sheetnr§§', str(config.sheet)).replace('§§tutorname§§', config.tutor_name),
                message_text = config.mail_template_text.replace('§§username§§', user).replace('§§name§§', name).replace('§§sheetnr§§', str(config.sheet)),
                attachments  = {
                    pdf['filename'] : pdf['data']
                    },
                bcc = config.mail_bcc
                )
        print(f'[OK]\t{user} -> {name} <{email}>')

####################################################################################################

def read_scores(infile):
    """Reads the scores from the Excel sheet while being as relaxed about the
    data type of the task / subtask columns as possible"""
    scores = pd.read_excel(infile, sheet_name = 'Grading', dtype={
        'Username' : str,
        'Sheet'    : 'Int64',
        'Task'     : 'Int64',
        'Subtask'  : 'Int64',
        'Type'     : str,
        'Value'    : object,
        })

    return scores

def process(config):
    # Read Scores and comments
    scores = read_scores(config.infile)

    # Read participants
    participants = pd.read_excel(config.infile, sheet_name = 'Participants').set_index('Username')

    # Read maximum scores
    sheet_meta       = pd.read_excel(config.infile, sheet_name = 'Sheets')
    max_scores       = { (row.Sheet, row.Task, row.Subtask) : row.MaxScore for _, row in sheet_meta.iterrows() }
    max_scores_sheet = sheet_meta.groupby('Sheet')['MaxScore'].sum()

    # Select sheet
    scores = scores[scores.Sheet==config.sheet]

    if scores.empty:
        print("No matching grades found")
        return 1

    # Read LaTeX template
    try:
        with open(config.tex_template, 'r') as infile:
            template = infile.read()
    except Exception as e:
        raise AToolError(f"Cannot open template file '{config.tex_template}': {e}")

    # Build dictionaries
    d = defaultdict(lambda : defaultdict( lambda : { 'score' : None, 'comments' : []} ) )
    sheet_comments = defaultdict(lambda : [])


    for _,row in scores.iterrows():
        if pd.isna(row.Sheet):
            raise AToolError('Failed to parse provided Excel file, "Grading" sheet contains empty value for "Sheet" column.')
        if row.Type.upper() == 'SCORE':
            # Check if the row checks out
            if pd.isna(row.Task) or pd.isna(row.Subtask) or pd.isna(row.Value):
                raise AToolError(f'Failed to parse provided Excel file, "Grading" sheet contains empty value for "Task", "Subtask" or "Value" column for sheet {row.Sheet}.')
            task = (row.Sheet, row.Task, row.Subtask)
            record = d[row.Username][task]
            # Check if a score occurs redundantly for the same task
            if record['score'] is not None:
                raise AToolError(f'Duplicate score for identical task found (User: {row.Username}, Sheet: {row.Sheet}, Task: {row.Task}, Subtask: {row.Subtask}')
            # Obtain maximum attainable score
            try:
                record['max_score'] = max_scores[task]
            except KeyError:
                raise AToolError(f'Could not find maximum score for task {task}')
            record['score'] = row.Value
        elif row.Type.upper() == 'COMMENT':
            # Check if this is a comment that applies to the entire sheet
            if pd.isna(row.Task) and pd.isna(row.Subtask):
                sheet_comments[row.Username].append(row.Value)
            elif pd.isna(row.Task) or pd.isna(row.Subtask):
                raise AToolError(f'Failed to parse provided Excel file, "Grading" sheet contains empty value for "Task" or "Subtask" but not for both.')
            # ... or if it's a comment that applies to a specific task
            else:
                task = (row.Sheet, row.Task, row.Subtask)
                record = d[row.Username][task]
                record['comments'].append(row.Value)
        else:
            raise AToolError(f'Invalid value type "{row.Type}".')

    mail_todo = {}
    for user, tasks in d.items():
        body=[]

        # Total score
        total_score = sum([ record['score'] for _, record in tasks.items() if record['score'] ])
        try:
            max_total_score = max_scores_sheet.loc[config.sheet]
        except KeyError:
            raise AToolError(f'Cannot calculate maximum total score for sheet {config.sheet}')

        # Participant name
        try:
            realname = participants.loc[user].Name
        except KeyError:
            raise AToolError(f'Cannot find real name for user "{user}".')


        tex=template.replace('§§sheetnr§§', str(config.sheet)).replace('§§fullname§§', str(realname)).replace('§§tutorname§§', config.tutor_name).replace('§§total§§', str(total_score)).replace('§§maxtotal§§', str(max_total_score))

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

        # Global comments
        global_ = []
        if user in sheet_comments:
            global_.append('\\beforeGlobalComments')
            global_comments = sheet_comments[user]
            for comment in global_comments:
                global_.append(f'\\globalComment{{{comment}}}')
            global_.append('\\afterGlobalComments')

        # Write out and compile tex
        tex = tex.replace('§§global§§', '\n'.join(global_))
        tex = tex.replace('§§body§§', '\n'.join(body))
        tex = tex.replace('§§tasks§§', '\n'.join(body))
        pdf = compileLaTeX(tex, config.pdflatex)

        # Move output file in place
        outpath = config.pdf_filename.replace('§§username§§', user).replace('§§name§§', realname).replace('§§sheetnr§§', str(config.sheet))

        # Store locally unless in mail only mode
        if not (config.no_local_file and config.mail):
            if os.path.exists(outpath):
                raise AToolError(f"Output path '{outpath}' exists! Aborting!")
            with open(outpath, 'wb') as outfile:
                outfile.write(pdf)
            print(f'[OK]\t{user}')

        # Prepare email if requested to do so
        if config.mail:
            mail_todo[user] = {
                    'filename' : os.path.basename(outpath),
                    'data' : pdf
                    }

    # Send out prepared emails
    if mail_todo:
        mail_feedback(config, participants, mail_todo)

    return 0

####################################################################################################

def main(argv=sys.argv):
    try:
        cfg = config.get_config()
        process(cfg)
    except AToolError as e:
        print(f'\nERROR: {e}')
        sys.exit(1)
    except KeyboardInterrupt:
        print(f'\nAborting.')
        sys.exit(2)
    except Exception as e:
        print(f'\nUnknown error: {e}')
        sys.exit(9)
