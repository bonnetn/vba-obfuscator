import re
import secrets
import string
from typing import Iterable

VBA_PATH = "download_payload.vba"
FIND_VAR_NAMES_REGEX = r"\s*Dim (\w+) "


def read_file(path: str) -> str:
    with open(path, "r") as f:
        return f.read()


def get_all_variables_iter(content: str) -> Iterable[str]:
    var_names = re.finditer(FIND_VAR_NAMES_REGEX, content)
    return map(lambda v: v.group(1), var_names)


def get_random_variable_name(size: int) -> str:
    return ''.join(secrets.choice(string.ascii_letters + string.digits) for _ in range(size))


def replace_var_name(content: str, var_name: str, new_name: str) -> str:
    var_name = re.escape(var_name)
    new_name = re.escape(new_name)
    # We have to use named groups here. If we had use \1 and \2, numbers in var_name could mess the regex.
    pattern = r"(?P<prefix>\s){}(?P<suffix>\s)".format(var_name)
    repl = r"\g<prefix>{}\g<suffix>".format(new_name)
    result = re.sub(pattern, repl, content)
    return result


if __name__ == "__main__":
    vba_content = read_file(VBA_PATH)
    for var in get_all_variables_iter(vba_content):
        var_length = 7 + secrets.randbelow(16)
        vba_content = replace_var_name(vba_content, var, get_random_variable_name(var_length))

    print(replace_var_name(vba_content, "exec", "test"))
