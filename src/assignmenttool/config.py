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
import configparser
import subprocess
import os

from assignmenttool.errors import AToolError

def config_from_cli():
    """Handle configuration passed through the command line"""

    parser = argparse.ArgumentParser()
    parser.add_argument('infile', type = str, metavar = '<sheetpath>', help = 'Path to the Excel file.')
    parser.add_argument('sheet', type = int, help = 'Sheet number to process')

    general = parser.add_argument_group('general settings')
    general.add_argument('--tex-template', type = str, metavar = '<texpath>', help = 'Path to the LaTeX template.')
    general.add_argument('--tutor-name', type = str, metavar = '<tutorname>', help = 'Name of the correcting tutor. Will be blank if not specified but used in the template.')
    general.add_argument('--pdflatex', type = str, metavar = '<pdflatexpath>', help = 'pdflatex command to use (default: pdflatex).', default='pdflatex')
    general.add_argument('--pdf-filename', type = str, metavar = '<name>', help = 'Output filename of the PDF feedback. May contain variables §§username§§, §§name§§ and §§sheetnr§§.')
    general.add_argument('--no-local-file', action = 'store_true', help = 'If --mail is specified, do not store PDFs locally.')
    parser.add_argument_group(general)

    mail = parser.add_argument_group('mail releated settings')
    mail.add_argument('--mail', action = 'store_true', help = 'Send out the feedback PDFs to the participants via email.')
    mail.add_argument('--mail-smtp-host', type = str, help = 'Hostname of the SMTP server to use for mail submission.')
    mail.add_argument('--mail-smtp-port', type = str, help = 'Hostname of the SMTP server to use for mail submission. Default: 587.')
    mail.add_argument('--mail-smtp-user', type = str, help = 'Username to use to authenticate at the SMTP server.')
    mail.add_argument('--mail-smtp-no-tls', action = 'store_true', help = 'Use SMTP without TLS.')
    mail.add_argument('--mail-sender-name', type = str, help = 'Sender name to use when sending out mails.')
    mail.add_argument('--mail-sender-address', type = str, help = 'Sender address to use when sending out mails.')
    mail.add_argument('--mail-bcc', type = str, nargs = '+', help = 'BCC recipient to add to every sent out email.')
    mail.add_argument('--mail-subject', type = str, help = 'Subject for outgoing emails to participants. May contain variables §§username§§, §§name§§ and §§sheetnr§§.')
    mail.add_argument('--mail-template', type = str, help = 'Path to the email body template for outgoing emails to participants. The template itself may contain variables §§username§§, §§name§§, §§sheetnr§§ and §§tutorname§§.')

    dev = parser.add_argument_group('developer settings')
    dev.add_argument('--debug', action = 'store_true', help = 'Do not remove temporary LaTeX build folder. Print path instead.')

    parser.add_argument_group(mail)
    config = parser.parse_args()

    return config

def read_rc(config):
    """Complements config specified by `conf` with values from the RC file(s) if not set in `conf`"""

    configfile = configparser.ConfigParser()
    # Two config files to read: '.assignmentrc' in $HOME and 'assignment.rc' in $PWD
    configfile.read([os.path.join(os.path.expanduser("~"), '.assignmentrc'), 'assignment.rc'])

    # Single value arguments
    for cli_arg, conf_group, conf_key in [
            ('tex_template', 'General', 'TexTemplate'),
            ('tutor_name', 'General', 'TutorName'),
            ('pdflatex', 'General', 'PDFLaTeX'),
            ('pdf_filename', 'General', 'PDFFilename'),
            ('mail_smtp_host', 'Mail', 'SMTPHost'),
            ('mail_smtp_port', 'Mail', 'SMTPPort'),
            ('mail_smtp_user', 'Mail', 'SMTPUser'),
            ('mail_smtp_pass', 'Mail', 'SMTPPass'),
            ('mail_bcc', 'Mail', 'BCC'),
            ('mail_subject', 'Mail', 'Subject'),
            ('mail_template', 'Mail', 'Template'),
            ('mail_sender_name', 'Mail', 'SenderName'),
            ('mail_sender_address', 'Mail', 'SenderAddress'),
            ]:
        if vars(config)[cli_arg] is None:
            try:
                vars(config)[cli_arg] = configfile[conf_group][conf_key]
            except KeyError:
                pass

    # Boolean arguments
    for cli_arg, conf_group, conf_key in [
            ('no_local_file', 'General', 'NoLocalFile'),
            ('debug', 'General', 'Debug'),
            ('mail_smtp_no_tls', 'Mail', 'NoTLS'),
            ]:
        if vars(config)[cli_arg] is False:
            try:
                vars(config)[cli_arg] = True if configfile[conf_group][conf_key] in ['yes', 'y', 'Yes', 'YES', 'True', 'TRUE', 'true', '1'] else False
            except KeyError:
                pass


    # List arguments
    for cli_arg, conf_group, conf_key in [
            ('mail_bcc', 'Mail', 'SMTPBCC')
            ]:
        if vars(config)[cli_arg] is None:
            try:
                vars(config)[cli_arg] = [ e.strip() for e in configfile[conf_group][conf_key].split(',') ]
            except KeyError:
                pass

def get_config():
    """Obtains and checks runtime configuration from CLI and RC file(s)"""

    # Read CLI config
    config = config_from_cli()
    vars(config)['mail_smtp_pass'] = None

    # Complement with RC config
    read_rc(config)

    # Check if template was specified
    if not config.tex_template:
        raise AToolError('No LaTeX template was specified. Use --tex-template or specify the path in th RC file.')

    # Check if pdflatex works
    ret = subprocess.run([config.pdflatex, '--version'], stdout = subprocess.PIPE, stderr = subprocess.PIPE)
    if ret.returncode != 0:
        raise AToolError(f"Failed to execute pdflatex ({config.pdflatex}). Please specify the path to the pdflatex binary using --pdflatex")

    # Default to blank tutor name
    if config.tutor_name is None:
        config.tutor_name = ''

    # Default to 587 as SMTP submission port
    if not config.mail_smtp_port:
        config.mail_smtp_port = 587

    if not config.pdf_filename:
        config.pdf_filename='Exercise§§sheetnr§§.§§username§§.feedback.pdf'

    # Check if mail config is present if requested
    if config.mail:
        for param in [
                '--mail-smtp-host',
                '--mail-smtp-user',
                '--mail-sender-name',
                '--mail-sender-address',
                '--mail-subject',
                '--mail-template',
                ]:
            if not vars(config)[param[2:].replace('-','_')]:
                raise AToolError(f'When using --mail, {param} must be specified.')
        # Read the mail template
        try:
            with open(config.mail_template, 'r') as infile:
                config.mail_template_text = infile.read().replace('§§tutorname§§', config.tutor_name)
        except Exception as e:
            raise AToolError(f"Failed to open mail template '{config.mail_template}': {e}")

    return config
