import logging
import random
import re
from typing import Iterable

from obfuscator.modifier.base import Modifier
from obfuscator.msdocument import MSDocument
from obfuscator.util import get_random_string

LOG = logging.getLogger(__name__)
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


class SplitStrings(Modifier):
    def run(self, doc: MSDocument) -> None:
        for str_found in _get_all_strings(doc.code):
            if len(str_found) > 8:
                pos = _split_string(str_found)
                splitted_string = '"{}"&"{}"'.format(str_found[:pos], str_found[pos:])
                doc.code = re.sub(re.escape('"{}"'.format(str_found)), splitted_string, doc.code, 1)
                LOG.debug("Splitted '{}' in two.".format(str_found))

        doc.code = doc.code


class CryptStrings(Modifier):
    def run(self, doc: MSDocument) -> None:
        for str_found in _get_all_strings(doc.code):
            key = get_random_string(len(str_found))
            array = _to_vba_array(_xor_crypt(str_found, key))
            unxor_eq = 'unxor({},"{}")'.format(array, key)
            doc.code = re.sub(re.escape('"{}"'.format(str_found)), unxor_eq, doc.code, 1)
            LOG.debug("Encrypted '{}'.".format(str_found))

        doc.code = VBA_XOR_FUNCTION + doc.code


def _get_all_strings(content: str) -> Iterable[str]:
    strings_found = re.finditer(FIND_STRINGS_REGEX, content)
    return set(map(lambda v: v.group(1), strings_found))


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
