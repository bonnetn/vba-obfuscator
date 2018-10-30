import logging

from obfuscator.log import configure_logging
from obfuscator.modifier.base import Pipe
from obfuscator.modifier.crypt_strings import CryptStrings
from obfuscator.modifier.randomize_function_names import RandomizeFunctionNames
from obfuscator.modifier.randomize_variable_names import RandomizeVariableNames
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
        RandomizeVariableNames(),
        CryptStrings(),
        RandomizeFunctionNames(),
    )

    LOG.info("Done!")
    print(doc.code)
