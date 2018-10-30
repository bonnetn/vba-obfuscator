import logging

from obfuscator.log import configure_logging
from obfuscator.modifier.base import Pipe
from obfuscator.modifier.func import RandomizeFunctionNames
from obfuscator.modifier.strings import CryptStrings, SplitStrings
from obfuscator.modifier.var import RandomizeVariableNames
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
        RandomizeFunctionNames(),
        RandomizeVariableNames(),
    )

    LOG.info("Done!")
    print(doc.code)
