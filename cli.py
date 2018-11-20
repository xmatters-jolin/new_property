"""Command line and argument processor

.. _Following Google Python Style Guide:
   http://google.github.io/styleguide/pyguide.htm
"""

import sys
import time
import json
import argparse
import getpass
from datetime import datetime

from requests import auth

import config
import np_logger
import processor


def process_sites(args):
    """Called when command line specifies sites"""
    np_logger.get_logger().debug('Processing Sites only')
    processor.process(['sites'])
    return

def process_admins(args):
    """Called when command line specifies admins"""
    np_logger.get_logger().debug('Processing Admins only')
    processor.process(['admins'])
    return

def process_groups(args):
    """Called when command line specifies groups"""
    np_logger.get_logger().debug('Processing Groups only')
    processor.process(['groups'])
    return

def process_all(args):
    """Called when command line specifies all operations"""
    np_logger.get_logger().debug('Processing Sites, Admins, and Groups')
    processor.process(['sites','admin','groups'])
    return

class _CLIError(Exception):
    """Generic exception to raise and log different fatal errors."""
    def __init__(self, msg, rc=config.ERR_CLI_EXCEPTION):
        super(_CLIError).__init__(type(self))
        self.result_code = rc
        self.msg = "E: %s" % msg

    def __str__(self):
        return self.msg

    def __unicode__(self):
        return self.msg

class __Password(argparse.Action):
    """Container to get and/or hold incoming password"""
    def __call__(self, parser, namespace, values, option_string): # pylint: disable=signature-differs
        if values is None:
            values = getpass.getpass()
        setattr(namespace, self.dest, values)

def process_command_line(argv=None, prog_doc=''): # pylint: disable=too-many-branches,too-many-statements
    """Evaluates and responds to passed in command line arguments"""
    logger = None

    if argv is None:
        argv = sys.argv
    else:
        sys.argv.extend(argv)

    program_version_message = '%%(prog)s %s (%s)' % (
        "v%s" % config.VERSION, config.UPDATED)
    program_license = """%s

  Created by %s on %s.
  Copyright %s

  Licensed under the %s
  %s

  Distributed on an "AS IS" basis without warranties
  or conditions of any kind, either express or implied.

USAGE
""" % (prog_doc.split("\n")[1], config.AUTHOR, config.DATE,
       config.COPYRIGHT, config.LICENSE, config.LICENSE_REF)

    try:
        # Setup argument parser
        parser = argparse.ArgumentParser(
            description=program_license,
            formatter_class=argparse.RawDescriptionHelpFormatter)
        subparsers = parser.add_subparsers(dest='command_name')
        # Add common arguments
        parser.add_argument("-c", "--console", dest="noisy",
                            action='store_true',
                            help=(
                                "If specified, will echo all log output to "
                                "the console at the requested verbosity based "
                                "on the -v option"))
        parser.add_argument("-d", "--defaults", dest="defaults_filename",
                            default="defaults.json",
                            help=(
                                "Specifes the name of the file containing "
                                "default settings [default: %(default)s]"))
        parser.add_argument("-f", "--pfile", dest="properties_filename",
                            default=None,
                            help=(
                                  "If not specified in the defaults file, use "
                                  "this for the input file .xlsx file. "
                                  "[default: %(default)s]"))
        parser.add_argument("-i", "--itype", dest="instance_type",
                            default="np",
                            choices=['np', 'prod'],
                            help=(
                                  "Specifies whether we are updating the "
                                  "Production (prod) or Non-Production "
                                  "(np) instance. "
                                  "[default: %(default)s]"))
        parser.add_argument("-l", "--lfile", dest="log_filename",
                            default=None,
                            help=(
                                "If not specified in the defaults file, use "
                                "-l to specify the base name of the log file. "
                                "The name will have a timestamp and .log "
                                "appended to the end."))
        parser.add_argument("-o", "--odir", dest="out_directory",
                            default=None,
                            help=(
                                  "If not specified in the defaults file, use -o"
                                  " to specify the file system location where "
                                  "the output files will be written."))
        parser.add_argument('-p', action=__Password, nargs='?',
                            dest='password', default=None,
                            help=(
                                  "If not specified in the defaults file, use -p"
                                  " to specify a password either on the command"
                                  " line, or be prompted"))
        parser.add_argument("-s", "--supervisors", dest="supervisors",
                            default=None,
                            help=(
                                  "If not specified in the defaults file, use "
                                  "this for the xMatters User IDs of the "
                                  "default Supervisor(s) for added users. "
                                  "This is a comma-separated list of values, "
                                  "e.g. mySuper.one,mySuper.two "
                                  "[default: %(default)s]"))
        parser.add_argument("-U", "--udf", dest="udf_name",
                          default=None,
                          help=("If not specified in the defaults file, use "
                                "this for the User Defined Field. "
                                "[default: %(default)s]"))
        parser.add_argument("-u", "--user", dest="user",
                            default=None,
                            help=("If not specified in the defaults file, use "
                                  "-u to specify the xmatters user id that has"
                                  " permissions to get Event and Notification "
                                  "data."))
        parser.add_argument("-V", "--version",
                          action='version', version=program_version_message)
        parser.add_argument("-v", dest="verbose",
                          action="count", default=0,
                          help=(
                                "set verbosity level.  Each occurrence of v "
                                "increases the logging level.  By default it "
                                "is ERRORs only, a single v (-v) means add "
                                "WARNING logging, a double v (-vv) means add "
                                "INFO logging, and a tripple v (-vvv) means "
                                "add DEBUG logging [default: %(default)s]"))
        parser.add_argument("-x", "--xmodurl", dest="xmod_url",
                            default=None,
                            help=("If not specified in the defaults file, use "
                                  "-i to specify the base URL of your xmatters"
                                  " instance.  For example, 'https://myco.host"
                                  "ed.xmatters.com' without quotes."))
        #Add in event command parsers
        sites_parser = subparsers.add_parser(
            'sites', description=("Processes only the 'Sites' worksheet"),
            help=("Use this command in order to only read and process Sites."))
        sites_parser.set_defaults(func=process_sites)
        admins_parser = subparsers.add_parser(
             'admins', description=("Processes only the 'Admins' worksheet"),
             help=("Use this command in order to only read and process Admins."))
        admins_parser.set_defaults(func=process_admins)
        groups_parser = subparsers.add_parser(
             'groups', description=("Processes only the 'Groups'"),
             help=("Use this command in order to only process Groups."))
        groups_parser.set_defaults(func=process_groups)
        all_parser = subparsers.add_parser(
            'all', description=("Processes Sites, Admins, and Groups"),
            help=("Use this command in order to process all worksheets "
                  "from the infput file: Sites, Admins, Groups."))
        all_parser.set_defaults(func=process_all)

        # Process arguments
        args = parser.parse_args()

        # Dereference the arguments into the configuration object
        user = None
        password = None
        if args.properties_filename:
            config.properties_filename = args.properties_filename
        if args.instance_type:
            config.instance_type = args.instance_type
        if args.log_filename:
            config.log_filename = args.log_filename
        if args.out_directory:
            config.out_directory = args.out_directory
        if args.noisy > 0:
            config.noisy = args.noisy
        if args.password:
            password = args.password
        if args.user:
            user = args.user
        if args.verbose > 0:
            config.verbosity = args.verbose
        if args.xmod_url:
            config.xmod_url = args.xmod_url
        if args.udf_name:
            config.udf_name = args.udf_name
        if args.supervisors:
            config.supervisors = args.supervisors.split(',')

        # Try to read in the defaults from defaults.json
        try:
            with open(args.defaults_filename) as defaults:
                cfg = json.load(defaults)
        except FileNotFoundError:
            raise(_CLIError(
                config.ERR_CLI_MISSING_DEFAULTS_MSG % args.defaults_filename,
                config.ERR_CLI_MISSING_DEFAULTS_CODE))

        # Process the defaults
        if user is None and 'user' in cfg:
            user = cfg['user']
        if password is None and 'password' in cfg:
            password = cfg['password']
        if config.dir_sep is None and 'dirSep' in cfg:
            config.dir_sep = cfg['dirSep']
        if config.properties_filename is None and 'propertiesFilename' in cfg:
            config.properties_filename = cfg['propertiesFilename']
        if config.log_filename is None and 'logFilename' in cfg:
            config.log_filename = cfg['logFilename']
        if config.out_directory is None and 'outDirectory' in cfg:
            config.out_directory = cfg['outDirectory']
        if config.xmod_url is None and 'xmodURL' in cfg:
            config.xmod_url = cfg['xmodURL']
        if config.verbosity == 0 and 'verbosity' in cfg:
            if cfg['verbosity'] in [1, 2, 3]:
                config.verbosity = cfg['verbosity']
        if config.udf_name is None and 'udfName' in cfg:
            config.udf_name = cfg['udfName']
        if config.supervisors is None and 'supervisors' in cfg:
            config.supervisors = cfg['supervisors'].split(',')
        if 'instance' in cfg:
            config.non_prod = True if cfg['instance'] == 'np' else False
        config.command_name = args.command_name

        # Fix file names
        time_str = time.strftime("-%Y%m%d-%H%M")
        if config.log_filename:
            config.log_filename = (
                config.out_directory + config.dir_sep +
                config.log_filename + time_str + '.log')

        # Initialize logging
        logger = np_logger.get_logger()
        logger.info("Four Seassons Property Processor Started.")
        logger.debug("After parser.parse_args(), command_name=%s",
                    args.command_name)

        # Final verification of arguments
        if config.xmod_url:
            logger.info("xmatters Instance URL is: %s", config.xmod_url)
        else:
            raise(_CLIError(config.ERR_CLI_MISSING_XMOD_URL_MSG,
                            config.ERR_CLI_MISSING_XMOD_URL_CODE))
        if user:
            logger.info("User is: %s", user)
        else:
            raise(_CLIError(config.ERR_CLI_MISSING_USER_MSG,
                            config.ERR_CLI_MISSING_USER_CODE))
        if password:
            logger.info("Password was provided.")
        else:
            raise(_CLIError(config.ERR_CLI_MISSING_PASSWORD_MSG,
                            config.ERR_CLI_MISSING_PASSWORD_CODE))
        if config.out_directory:
            logger.info("Output directory is: %s", config.out_directory)
        else:
            raise(_CLIError(config.ERR_CLI_MISSING_OUTPUT_DIR_MSG,
                            config.ERR_CLI_MISSING_OUTPUT_DIR_CODE))
        if config.properties_filename:
            logger.info("Properties input filename is: %s",
                        config.properties_filename)
        else:
            raise(_CLIError(config.ERR_CLI_MISSING_PROPERTIES_FILENAME_MSG,
                            config.ERR_CLI_MISSING_PROPERTIES_FILENAME_CODE))
        if config.supervisors:
            logger.info("Default Admin/Users Supervisor(s): %s",
                        config.supervisors)
        else:
            raise(_CLIError(config.ERR_CLI_MISSING_SUPERVISORS_MSG,
                            config.ERR_CLI_MISSING_SUPERVISORS_CODE))
        if args.command_name:
            logger.info("About to begin processing command(s): %s",
                        config.command_name)
        else:
            raise(_CLIError(config.ERR_CLI_MISSING_COMMAND_MSG,
                            config.ERR_CLI_MISSING_COMMAND_CODE))

        # Setup the basic auth object for subsequent REST calls
        config.basic_auth = auth.HTTPBasicAuth(user, password)

        # Make sure we have a func None == all
        if args.func is None:
            args.func = process_all

        return args

    except KeyboardInterrupt:
        ### handle keyboard interrupt ###
        sys.exit(0)

    except _CLIError as cli_except:
        if config.DEBUG or config.TESTRUN:
            raise cli_except # pylint: disable=raising-bad-type
        msg = config.program_name + ": Command Line Error - " + cli_except.msg + " (" + str(cli_except.result_code) + ")"
        if logger:
            logger.error(msg)
        else:
            sys.stderr.write(msg+"\n")
        indent = (len(config.program_name) + 30) * " "
        sys.stderr.write(indent + "  for help use --help\n")
        sys.exit(cli_except.result_code)

    except Exception as exc: # pylint: disable=broad-except
        if config.DEBUG or config.TESTRUN:
            raise exc # pylint: disable=raising-bad-type
        sys.stderr.write(
            config.program_name + ": Unexpected exception " + repr(exc) + "\n")
        indent = (len(config.program_name) + 30) * " "
        sys.stderr.write(indent + "  For assistance use --help\n")
        sys.exit(config.ERR_CLI_EXCEPTION)

def main():
    """ By convention and for completeness """
    pass

if __name__ == '__main__':
    main()
