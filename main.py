import re
import secrets
import string
from typing import Iterable

VBA_PATH = "example_macro/download_payload.vba"
FIND_VAR_NAMES_REGEX = r"\s*Dim\s+\b(\w+)\b"
FIND_STRINGS_REGEX = r'"(([^"]|"")*)"'
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


def read_file(path: str) -> str:
    with open(path, "r") as f:
        return f.read()


def get_all_strings_iter(content: str) -> Iterable[str]:
    strings_found = re.finditer(FIND_STRINGS_REGEX, content)
    return map(lambda v: v.group(1), strings_found)


def get_all_variables_iter(content: str) -> Iterable[str]:
    var_names = re.finditer(FIND_VAR_NAMES_REGEX, content)
    return map(lambda v: v.group(1), var_names)


def get_random_string(size: int) -> str:
    return ''.join(secrets.choice(string.ascii_letters + string.digits) for _ in range(size))


def replace_var_name(content: str, var_name: str, new_name: str) -> str:
    var_name = re.escape(var_name)
    new_name = re.escape(new_name)
    pattern = r"\b{}\b".format(var_name)
    result = re.sub(pattern, new_name, content)
    return result


def xor(t):
    a, b = t
    return a ^ b


def xor_crypt(str, key):
    str = map(ord, str)
    key = map(ord, key)
    str_key = zip(str, key)
    return map(xor, str_key)


def to_vba_array(arr):
    arr = map(str, arr)
    numbers = ",".join(arr)
    return "Array({})".format(numbers)


if __name__ == "__main__":
    vba_content = read_file(VBA_PATH)

    print("Replacing variable names...")
    for var in get_all_variables_iter(vba_content):
        var_length = 7 + secrets.randbelow(16)
    new_name = get_random_string(var_length)
    vba_content = replace_var_name(vba_content, var, new_name)
    print("- Replace '{}' with '{}'.".format(var, new_name))

    print("Obfuscating strings.")

    for str_found in get_all_strings_iter(vba_content):
        print("String found: {}".format(str_found))
        key = get_random_string(len(str_found))
        array = to_vba_array(xor_crypt(str_found, key))
        code = 'unxor({}, "{}")'.format(array, key)

        print("->", code)

    print("=" * 100)
    print(replace_var_name(vba_content, "exec", "test"))
