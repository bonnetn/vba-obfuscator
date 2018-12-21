#!/usr/bin/env python3
import argparse
import logging
import sys

from obfuscator.log import configure_logging
from obfuscator.modifier.base import Pipe
from obfuscator.modifier.comments import StripComments
from obfuscator.modifier.functions_vars import RandomizeNames
from obfuscator.modifier.misc import RemoveEmptyLines
from obfuscator.modifier.numbers import ReplaceIntegersWithAddition, ReplaceIntegersWithXor
from obfuscator.modifier.strings import CryptStrings, SplitStrings
from obfuscator.msdocument import MSDocument

if __name__ == "__main__":
    configure_logging()

    LOG = logging.getLogger(__name__)
    LOG.info("VBA obfuscator - Thomas LEROY & Nicolas BONNET")

    parser = argparse.ArgumentParser(description='Obfuscate a VBA file.')
    parser.add_argument('input_file', type=str, action='store',
                        help='path of the file to obfuscate')
    parser.add_argument('--output_file', type=str, action='store',
                        help='output file (if no file is supplied, stdout will be used)')
    args = parser.parse_args()

    LOG.info("Loading the document...")
    doc = MSDocument(args.input_file)

    LOG.info("Obfuscating the code...")
    Pipe(doc).run(
        SplitStrings(),
        CryptStrings(),
        RandomizeNames(),
        ReplaceIntegersWithAddition(),
        ReplaceIntegersWithXor(),
        StripComments(),
        RemoveEmptyLines(),
    )

    LOG.info("Done!")

    if args.output_file:
        with open(args.output_file, "w") as f:
            f.write(doc.code)
    else:
        sys.stdout.write(doc.code)
