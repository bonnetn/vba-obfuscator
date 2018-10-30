import re
from typing import Iterable

from obfuscator.modifier.base import Modifier
from obfuscator.msdocument import MSDocument
from obfuscator.util import get_random_string

FIND_STRINGS_REGEX = r'"((?:[^"]|"")*)"'
VBA_XOR_FUNCTION = """
Private Function unxor(ciphertext As Variant, key As String)
    Dim cleartext As String
    cleartext = ""
    
    For i = LBound(ciphertext) To UBound(ciphertext)
        keyChar = Mid(key, i + 1, 1)
        cleartext = cleartext & Chr(Asc(keyChar) Xor ciphertext(i))
    Next
    unxor = cleartext

End Function


"""


class CryptStrings(Modifier):
    def run(self, doc: MSDocument) -> None:
        print("Encrypthing strings:")
        for str_found in _get_all_strings(doc.code):
            key = get_random_string(len(str_found))
            array = _to_vba_array(_xor_crypt(str_found, key))
            unxor_eq = 'unxor({},"{}")'.format(array, key)
            doc.code = re.sub(re.escape('"{}"'.format(str_found)), unxor_eq, doc.code, 1)
            print('- Replaced "{}" by "{}".'.format(str_found, unxor_eq))


def _get_all_strings(content: str) -> Iterable[str]:
    strings_found = re.finditer(FIND_STRINGS_REGEX, content)
    return map(lambda v: v.group(1), strings_found)


def _xor(t):
    a, b = t
    return a ^ b


def _xor_crypt(msg, key):
    msg = map(ord, msg)
    key = map(ord, key)
    str_key = zip(msg, key)
    return map(_xor, str_key)


def _to_vba_array(arr):
    arr = map(str, arr)
    numbers = ",".join(arr)
    return "Array({})".format(numbers)
