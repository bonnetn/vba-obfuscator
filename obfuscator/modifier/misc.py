import re

from obfuscator.modifier.base import Modifier
from obfuscator.msdocument import MSDocument


class RemoveEmptyLines(Modifier):
    def run(self, doc: MSDocument) -> None:
        doc.code = re.sub(r'\n+', '\n', doc.code)
