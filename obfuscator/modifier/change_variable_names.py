import re
import secrets
from typing import Iterable

import obfuscator.msdocument
import obfuscator.modifier.base
from obfuscator.util import get_random_string

FIND_VAR_NAMES_REGEX = r"\s*Dim\s+\b(\w+)\b"


class ChangeVariableNames(obfuscator.modifier.base.Modifier):
    def run(self, doc: obfuscator.msdocument.MSDocument) -> None:
        print("Obfuscating variable names:")
        for var in _get_variable_names(doc.code):
            var_length = 7 + secrets.randbelow(16)
            new_name = get_random_string(var_length)
            doc.code = _change_variable_name(doc.code, var, new_name)
            print("- Replaced '{}' by '{}'.".format(var, new_name))
        print("Done!")


def _get_variable_names(content: str) -> Iterable[str]:
    var_names = re.finditer(FIND_VAR_NAMES_REGEX, content)
    return map(lambda v: v.group(1), var_names)


def _change_variable_name(content: str, var_name: str, new_name: str) -> str:
    var_name = re.escape(var_name)
    new_name = re.escape(new_name)
    pattern = r"\b{}\b".format(var_name)
    result = re.sub(pattern, new_name, content)
    return result
