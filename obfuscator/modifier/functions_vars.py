import logging

from pygments import highlight
from pygments.formatter import Formatter
from pygments.lexers.dotnet import VbNetLexer
from pygments.token import Token

import obfuscator.modifier.base
import obfuscator.msdocument
from obfuscator.util import get_random_string_of_random_length

LOG = logging.getLogger(__name__)


class RandomizeNames(obfuscator.modifier.base.Modifier):
    def run(self, doc: obfuscator.msdocument.MSDocument) -> None:
        doc.code = highlight(doc.code, VbNetLexer(), _RandomizeNamesFormatter())


def _get_or_create(i: str, dct: dict) -> str:
    if i not in dct:
        dct[i] = get_random_string_of_random_length()
        LOG.debug("Randomized {}.".format(i))

    return dct[i]


def _blacklisted_functions(f):
    d = {}
    for name in f:
        d[name] = name
    return d


class _RandomizeNamesFormatter(Formatter):
    def __init__(self):
        self.names = _blacklisted_functions({
            "AutoOpen", "Workbook_Open", "Shell", "Array", "LBound", "Auto_Open", "ActiveDocument", "Variables", "Chr",
            "UBound", "Err", "Raise", "vbObjectError", "Asc", "H40", "H10", "HF", "Dim"
        })

    def format(self, tokensource, outfile):

        for ttype, value in tokensource:
            if ttype == Token.Name.Function:
                outfile.write(self._get_name(value))
            elif ttype == Token.Name:
                outfile.write(self._get_name(value))
            else:
                outfile.write(value)

    def _get_name(self, name: str) -> str:
        return _get_or_create(name, self.names)
