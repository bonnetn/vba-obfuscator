#!/usr/bin/env python3

import logging

from obfuscator.log import configure_logging
from obfuscator.modifier.base import Pipe
from obfuscator.modifier.comments import StripComments
from obfuscator.modifier.misc import RemoveEmptyLines
from obfuscator.modifier.numbers import ObfuscateIntegers
from obfuscator.modifier.strings import CryptStrings, SplitStrings
from obfuscator.modifier.functions_vars import RandomizeNames
from obfuscator.msdocument import MSDocument

VBA_PATH = "example_macro/download_payload.vba"

if __name__ == "__main__":
    configure_logging()

    LOG = logging.getLogger(__name__)
    LOG.info("VBA obfuscator - Thomas LEROY & Nicolas BONNET")

    LOG.info("Loading the document...")
    doc = MSDocument(VBA_PATH)

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

    print("=" * 100)
    print(doc.code)
    with open("output.vba", "w") as f:
        f.write(doc.code)
