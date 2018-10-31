import logging
import re
from typing import Iterable

import obfuscator.modifier.base
import obfuscator.msdocument
from obfuscator.util import get_random_string_of_random_length, replace_whole_word

LOG = logging.getLogger(__name__)
FIND_DEFINITION_VAR_NAMES_REGEX = r"\b(\w+)\b\s+\bAs\b"  #  Find definitions of variables like "XXX As String"
FIND_ASSIGNMENT_VAR_NAMES_REGEX = r"\b(\w+)\b\s*="  #  Find assignment of variables like "XXX = 42"


class RandomizeVariableNames(obfuscator.modifier.base.Modifier):
    def run(self, doc: obfuscator.msdocument.MSDocument) -> None:
        for var in _get_variable_names(doc.code):
            new_name = get_random_string_of_random_length()
            doc.code = replace_whole_word(doc.code, var, new_name)
            LOG.debug("Randomized '{}'.".format(var))


def _get_variable_names(content: str) -> Iterable[str]:
    assignment_vars = re.finditer(FIND_ASSIGNMENT_VAR_NAMES_REGEX, content)
    assignment_vars = set(map(lambda v: v.group(1), assignment_vars))

    definition_vars = re.finditer(FIND_DEFINITION_VAR_NAMES_REGEX, content)
    definition_vars = set(map(lambda v: v.group(1), definition_vars))

    return assignment_vars | definition_vars
