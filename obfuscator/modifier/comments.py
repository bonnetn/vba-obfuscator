import re

from obfuscator.modifier.base import Modifier
from obfuscator.msdocument import MSDocument


class StripComments(Modifier):
    def run(self, doc: MSDocument) -> None:
        # Remove comments
        pattern = r"([^'\"\n]*)'.*$"
        doc.code = re.sub(pattern, r"\1", doc.code, flags=re.MULTILINE)

        # Remove empty lines
        pattern = r"(\n\s*)+\n"
        doc.code = re.sub(pattern, r"\n", doc.code, flags=re.MULTILINE)

        # Remove first empty line
        pattern = r"^(\s*\n)"
        doc.code = re.sub(pattern, r"", doc.code, flags=re.MULTILINE)
