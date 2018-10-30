import logging
import re
from typing import Iterable

import obfuscator.modifier.base
import obfuscator.msdocument
from obfuscator.util import get_random_string_of_random_length, replace_whole_word

LOG = logging.getLogger(__name__)
FIND_VAR_NAMES_REGEX = r"\s*Dim\s+\b(\w+)\b"


class RandomizeVariableNames(obfuscator.modifier.base.Modifier):
    def run(self, doc: obfuscator.msdocument.MSDocument) -> None:
        for var in _get_variable_names(doc.code):
            new_name = get_random_string_of_random_length()
            doc.code = replace_whole_word(doc.code, var, new_name)
            LOG.debug("Replaced '{}' with '{}'.".format(var, new_name))


def _get_variable_names(content: str) -> Iterable[str]:
    var_names = re.finditer(FIND_VAR_NAMES_REGEX, content)
    return map(lambda v: v.group(1), var_names)
