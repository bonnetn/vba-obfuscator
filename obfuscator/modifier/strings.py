import base64
import logging
import random
from typing import List

from pygments import highlight
from pygments.formatter import Formatter
from pygments.lexers.dotnet import VbNetLexer
from pygments.token import Token

from obfuscator.modifier.base import Modifier
from obfuscator.msdocument import MSDocument
from obfuscator.util import get_random_string, split_var_declaration_from_code

LOG = logging.getLogger(__name__)
VBA_XOR_FUNCTION = split_var_declaration_from_code("""
Private Function unxor(ciphertext As Variant, start As Integer)
    Dim cleartext As String
    Dim key() As Byte
    key = Base64Decode(ActiveDocument.Variables("{}"))
    cleartext = ""
    
    For i = LBound(ciphertext) To UBound(ciphertext)
        cleartext = cleartext & Chr(key(i+start) Xor ciphertext(i))
    Next
    unxor = cleartext

End Function
""")
with open("base64.vbs") as f:
    VBA_BASE64_FUNCTION = split_var_declaration_from_code(f.read())


class SplitStrings(Modifier):
    def run(self, doc: MSDocument) -> None:
        doc.code = highlight(doc.code, VbNetLexer(), _SplitStrings())


class CryptStrings(Modifier):
    def run(self, doc: MSDocument) -> None:
        LOG.debug('Generating document variable name.')

        formatter = _EncryptStrings()
        doc.code = highlight(doc.code, VbNetLexer(), formatter)

        document_var = get_random_string(16)

        code_prefix, code_suffix = split_var_declaration_from_code(doc.code)

        # Merge the codes: we must keep the global variables declarations on top.
        doc.code = code_prefix + VBA_BASE64_FUNCTION[0] + VBA_XOR_FUNCTION[0] + \
                   code_suffix + VBA_BASE64_FUNCTION[1] + VBA_XOR_FUNCTION[1].format(document_var)

        b64 = base64.b64encode(bytes(formatter.crypt_key)).decode()
        MAX_LENGTH = 512
        printable_b64 = [b64[i:i + MAX_LENGTH] for i in range(0, len(b64), MAX_LENGTH)]
        printable_b64 = '" & _\n"'.join(printable_b64)
        LOG.info('''Paste this in your VBA editor to add the Document Variable:
ActiveDocument.Variables.Add Name:="{}", Value:="{}"'''.format(document_var, printable_b64))

        doc.code = '"Use this line to add the document variable to you file and then remove these comments."\n' + \
                   'ActiveDocument.Variables.Add Name:="{}", Value:="{}"\n'.format(document_var,
                                                                                   printable_b64) + doc.code

        doc.doc_var[document_var] = b64


def _get_random_key(n: int) -> List[int]:
    return [random.randint(0, 255) for _ in range(n)]


def _xor(t):
    a, b = t
    return a ^ b


def _xor_crypt(msg, key):
    msg = map(ord, msg)
    str_key = zip(msg, key)
    return map(_xor, str_key)


def _to_vba_array(arr):
    arr = map(str, arr)
    numbers = ",".join(arr)
    return "Array({})".format(numbers)


def _split_string(s: str) -> int:
    """
    Split a string in two. This function will never split an escaped double quote ("") in half.
    :param s:
    :return:
    """
    split_possibilities = len(s) - 1
    impossible_split_pos = s.count('"') // 2
    if split_possibilities - impossible_split_pos <= 0:
        return -1

    split_possibilities -= impossible_split_pos

    i = 0
    pos = random.randint(1, split_possibilities)
    while pos > 0:
        if s[i] == '"':
            i = i + 1
        i = i + 1
        pos = pos - 1

    return i


class _StringFormatter(Formatter):
    def __init__(self, **options):
        super().__init__(**options)
        self.lastval = ""
        self.lasttype = None

    def _run_on_string(self, s: str) -> str:
        raise NotImplementedError()

    def format(self, tokensource, outfile):
        skip_line = False
        for ttype, value in tokensource:
            if self.lasttype:
                if self.lasttype == ttype:
                    self.lastval += value
                else:
                    if ttype == Token.Keyword and value == "Const":  #  Skip the line if it is a const.
                        skip_line = True
                    if "\n" in self.lastval:
                        skip_line = False

                    # Crypt strings unless we are skipping the line.
                    if self.lasttype == Token.Literal.String:
                        if skip_line:
                            outfile.write(self.lastval)
                        else:
                            outfile.write(self._run_on_string(self.lastval[1:-1]))
                    else:
                        outfile.write(self.lastval)
                    self.lastval = value
            else:
                self.lastval = value
            self.lasttype = ttype

        outfile.write(value)


class _SplitStrings(_StringFormatter):
    def __init__(self):
        super().__init__()
        self.crypt_key = []

    def _split_string(self, s: str) -> str:
        if len(s) > 8:
            s = s.strip('"')
            pos = _split_string(s)
            splitted_string = '"{}" & "{}"'.format(s[:pos], s[pos:])
            LOG.debug("Splitted '{}' in two.".format(s))
            return splitted_string
        else:
            return s

    def _run_on_string(self, s: str):
        return self._split_string(s)


class _EncryptStrings(_StringFormatter):
    def __init__(self):
        super().__init__()
        self.crypt_key = []

    def _obfuscate_string(self, s: str) -> str:
        LOG.debug("Generating XOR key for '{}'.".format(s))
        start = len(self.crypt_key)
        key = _get_random_key(len(s))
        self.crypt_key += key
        LOG.debug("XOR key will be at [{}; {}].".format(start, len(key)))

        ciphertext = _xor_crypt(s, key)
        array = _to_vba_array(ciphertext)
        LOG.debug("Encrypted string to VBA Array -> {}.".format(array))

        return 'unxor({},{})'.format(array, start)

    def _run_on_string(self, s: str):
        return self._obfuscate_string(s)
